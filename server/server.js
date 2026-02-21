const express = require("express");
const path = require("path");
const cheerio = require("cheerio");

const app = express();
const PORT = process.env.PORT || 3000;
const COCO_URL = "https://miau.my-x.hu/myx-free/coco/beker_y0.php";
const COCO_ENGINE_URL = "https://miau.my-x.hu/myx-free/coco/engine3.php";
const COCO_FETCH_TIMEOUT_MS = 60000;
const COCO_FETCH_RETRIES = 3;
const COCO_RETRY_DELAY_MS = 1500;

app.use(express.json({ limit: "5mb" }));
app.use(express.urlencoded({ extended: true, limit: "5mb" }));
app.use(express.static(path.join(__dirname, "..", "frontend")));

app.get("/health", (_req, res) => {
  res.json({ ok: true });
});

app.get("/api/coco-health", async (_req, res) => {
  try {
    const start = Date.now();
    const r = await fetchWithRetry(
      COCO_URL,
      {
        method: "GET",
        headers: { "User-Agent": "Mozilla/5.0 WHR-OAM-COCO-Proxy" }
      },
      COCO_FETCH_TIMEOUT_MS,
      2
    );
    const ms = Date.now() - start;
    res.json({ ok: r.ok, status: r.status, latencyMs: ms, url: COCO_URL });
  } catch (error) {
    res.status(502).json({ ok: false, message: `COCO unreachable: ${error.message}` });
  }
});

app.post("/api/coco-y0", async (req, res) => {
  try {
    const { matrix, objectNames, attributeNames } = req.body || {};
    if (!Array.isArray(matrix) || !Array.isArray(objectNames) || !Array.isArray(attributeNames)) {
      return res.status(400).json({ ok: false, message: "matrix, objectNames, attributeNames are required arrays." });
    }

    const matrixText = buildMatrixText(matrix, objectNames, attributeNames);
    const proxyResult = await tryRunCoco(matrixText, objectNames, attributeNames);

    return res.json({
      ok: true,
      automated: proxyResult.automated,
      estimations: proxyResult.estimations,
      message: proxyResult.message,
      matrixText,
      rawHtml: proxyResult.rawHtml || ""
    });
  } catch (error) {
    return res.status(500).json({
      ok: false,
      automated: false,
      message: `Proxy error: ${error.message}`,
      estimations: []
    });
  }
});

app.get("*", (_req, res) => {
  res.sendFile(path.join(__dirname, "..", "frontend", "index.html"));
});

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});

function buildMatrixText(matrix, objectNames, attributeNames) {
  void objectNames;
  void attributeNames;
  // COCO parser is sensitive: space-separated numbers + CR row breaks works reliably.
  const rows = matrix.map((row) => row.map((v) => (Number.isFinite(v) ? String(v) : "")).join(" "));
  return rows.join("\r");
}

async function tryRunCoco(matrixText, objectNames, attributeNames) {
  let lastReason = "";
  try {
    // 1) Open COCO page to collect hidden form values.
    const pageResp = await fetchWithRetry(
      COCO_URL,
      {
        method: "GET",
        headers: { "User-Agent": "Mozilla/5.0 WHR-OAM-COCO-Proxy" }
      },
      COCO_FETCH_TIMEOUT_MS,
      COCO_FETCH_RETRIES
    );
    const pageHtml = await pageResp.text();
    const formMeta = parseFormMeta(pageHtml);

    // 2) Use exact COCO field names from beker_y0.php form.
    const payloadCandidates = buildPayloadCandidates(formMeta.hidden, matrixText, objectNames, attributeNames);
    for (const payload of payloadCandidates) {
      const targetCandidates = [formMeta.action || COCO_ENGINE_URL, COCO_ENGINE_URL];
      for (const targetUrl of targetCandidates) {
        try {
          const postResp = await fetchWithRetry(
            targetUrl,
            {
              method: "POST",
              headers: {
                "Content-Type": "application/x-www-form-urlencoded",
                "User-Agent": "Mozilla/5.0 WHR-OAM-COCO-Proxy",
                Referer: COCO_URL
              },
              body: new URLSearchParams(payload)
            },
            COCO_FETCH_TIMEOUT_MS,
            COCO_FETCH_RETRIES
          );
          const html = await postResp.text();
          const parsed = parseEstimationsFromHtml(html, objectNames);
          if (parsed.ok) {
            return {
              automated: true,
              estimations: parsed.estimations,
              message: "COCO Y0 automation succeeded.",
              rawHtml: html
            };
          }
          // Some COCO responses contain only a summary page. Follow "Open url" detail page and parse there.
          const detailUrl = extractOpenUrl(html, targetUrl);
          if (detailUrl) {
            try {
              const detailResp = await fetchWithRetry(
                detailUrl,
                {
                  method: "GET",
                  headers: { "User-Agent": "Mozilla/5.0 WHR-OAM-COCO-Proxy", Referer: COCO_URL }
                },
                COCO_FETCH_TIMEOUT_MS,
                COCO_FETCH_RETRIES
              );
              const detailHtml = await detailResp.text();
              const parsedDetail = parseEstimationsFromHtml(detailHtml, objectNames);
              if (parsedDetail.ok) {
                return {
                  automated: true,
                  estimations: parsedDetail.estimations,
                  message: "COCO Y0 automation succeeded (detail page parse).",
                  rawHtml: detailHtml
                };
              }
            } catch (_detailErr) {
              // Ignore and continue payload fallbacks.
            }
          }
          lastReason = "COCO responded but estimation rows could not be parsed reliably.";
        } catch (postErr) {
          lastReason = `POST failed (${targetUrl}): ${postErr.message}`;
        }
      }
    }
  } catch (_error) {
    lastReason = _error.message;
    return {
      automated: false,
      estimations: [],
      message: `COCO network fetch failed (${lastReason}). Use manual fallback (matrix/object/attribute + paste estimations).`,
      rawHtml: ""
    };
  }

  // 3) Fallback for manual run.
  return {
    automated: false,
    estimations: [],
    message: `COCO could not be executed automatically (${lastReason || "unknown reason"}). Use matrix export + manual estimation paste fallback.`,
    rawHtml: ""
  };
}

function parseFormMeta(html) {
  const $ = cheerio.load(html);
  const form = $("form").first();
  const action = form.attr("action")
    ? new URL(form.attr("action"), COCO_URL).toString()
    : COCO_URL;
  const hidden = {};
  form.find("input[type='hidden']").each((_i, el) => {
    const name = $(el).attr("name");
    const value = $(el).attr("value") || "";
    if (name) hidden[name] = value;
  });
  return { action, hidden };
}

function buildPayloadCandidates(hiddenFields, matrixText, objectNames, attributeNames) {
  const now = Date.now();
  return [
    {
      ...hiddenFields,
      job: `whr_${now}`,
      matrix: matrixText,
      stair: "50",
      modell: "Y0",
      object: objectNames.join("\n"),
      attribute: attributeNames.join("\n"),
      button2: "Futtatás"
    },
    {
      ...hiddenFields,
      job: `whr_${now}`,
      matrix: matrixText,
      stair: String(Math.max(2, objectNames.length)),
      modell: "Y0",
      object: objectNames.join("\r\n"),
      attribute: attributeNames.join("\r\n"),
      button2: "Futtatás"
    }
  ];
}

function parseEstimationsFromHtml(html, objectNames) {
  const $ = cheerio.load(html);
  const normalize = (x) => String(x || "").trim().toLowerCase();
  const wanted = new Map(objectNames.map((n, i) => [normalize(n), i]));
  const out = new Array(objectNames.length).fill(null);
  const isPlausibleEstimation = (v) => Number.isFinite(v) && v > 0 && v < 10000;

  // COCO output rows typically have: [objectName, estimation, ...]
  $("tr").each((_i, tr) => {
    const tds = $(tr).find("td");
    if (tds.length < 2) return;
    const first = $(tds[0]).text().trim();
    const idx = wanted.get(normalize(first));
    if (idx === undefined) return;

    for (let c = 1; c < tds.length; c += 1) {
      const txt = $(tds[c]).text().trim().replace(",", ".");
      const num = Number(txt);
      if (isPlausibleEstimation(num)) {
        out[idx] = num;
        break;
      }
    }
  });

  const okCount = out.filter((x) => isPlausibleEstimation(x)).length;
  // Require a high match ratio; otherwise treat as parse failure and force manual fallback.
  if (okCount >= Math.max(3, Math.floor(objectNames.length * 0.8))) {
    return { ok: true, estimations: out };
  }
  return { ok: false, estimations: [] };
}

function extractOpenUrl(html, baseUrl) {
  const $ = cheerio.load(html);
  const a = $("a")
    .filter((_i, el) => /open url/i.test($(el).text() || ""))
    .first();
  if (!a.length) return null;
  const href = a.attr("href");
  if (!href) return null;
  try {
    return new URL(href, baseUrl || COCO_URL).toString();
  } catch (_e) {
    return null;
  }
}

async function fetchWithTimeout(url, options, timeoutMs) {
  const controller = new AbortController();
  const id = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetch(url, { ...options, signal: controller.signal });
  } finally {
    clearTimeout(id);
  }
}

async function fetchWithRetry(url, options, timeoutMs, retries) {
  let lastError = null;
  for (let attempt = 1; attempt <= retries; attempt += 1) {
    try {
      const response = await fetchWithTimeout(url, options, timeoutMs);
      if (!response.ok && response.status >= 500) {
        throw new Error(`HTTP ${response.status}`);
      }
      return response;
    } catch (err) {
      lastError = err;
      if (attempt < retries) {
        await sleep(COCO_RETRY_DELAY_MS * attempt);
      }
    }
  }
  throw lastError || new Error("Unknown fetch error");
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
