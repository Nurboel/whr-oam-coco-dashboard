const state = {
  rows: [],
  cocoRows: [],
  matrix: [],
  matrixText: "",
  objectNamesText: "",
  attributeNamesText: "",
  mapping: {},
  results: []
};

const requiredAliases = {
  country: ["country name", "country", "countryname"],
  ladder: ["ladder score", "ladder", "life ladder"],
  gdp: ["explained by: log gdp per capita", "log gdp per capita", "gdp"],
  soc: ["explained by: social support", "social support"],
  life: ["explained by: healthy life expectancy", "healthy life expectancy"],
  free: ["explained by: freedom to make life choices", "freedom to make life choices"],
  gen: ["explained by: generosity", "generosity"],
  corr: ["explained by: perceptions of corruption", "perceptions of corruption", "corruption"],
  dystopia: ["dystopia + residual", "dystopia residual", "dystopia"]
};

const el = {
  fileInput: document.getElementById("fileInput"),
  processBtn: document.getElementById("processBtn"),
  uploadStatus: document.getElementById("uploadStatus"),
  mappingBody: document.querySelector("#mappingTable tbody"),
  rawBody: document.querySelector("#rawTable tbody"),
  matrixBody: document.querySelector("#matrixTable tbody"),
  matrixTextArea: document.getElementById("matrixTextArea"),
  objectNamesTextArea: document.getElementById("objectNamesTextArea"),
  attributeNamesTextArea: document.getElementById("attributeNamesTextArea"),
  exportMatrixBtn: document.getElementById("exportMatrixBtn"),
  manualEstimationInput: document.getElementById("manualEstimationInput"),
  applyManualBtn: document.getElementById("applyManualBtn"),
  correlations: document.getElementById("correlations"),
  worldMaps: document.getElementById("worldMaps"),
  resultBody: document.querySelector("#resultTable tbody"),
  countrySearch: document.getElementById("countrySearch"),
  clearSearchBtn: document.getElementById("clearSearchBtn"),
  countryDrawer: document.getElementById("countryDrawer"),
  drawerTitle: document.getElementById("drawerTitle"),
  drawerBody: document.getElementById("drawerBody"),
  closeDrawerBtn: document.getElementById("closeDrawerBtn"),
  exportAllMapsBtn: document.getElementById("exportAllMapsBtn"),
  exportCsvBtn: document.getElementById("exportCsvBtn"),
  exportXlsxBtn: document.getElementById("exportXlsxBtn")
};

function normalizeHeader(value) {
  return String(value || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function parseNum(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (value === null || value === undefined) return null;
  const cleaned = String(value).trim().replace(/\s+/g, "").replace(",", ".");
  if (!cleaned) return null;
  const out = Number(cleaned);
  return Number.isFinite(out) ? out : null;
}

function round3(v) {
  if (v === null) return null;
  return Math.round(v * 1000) / 1000;
}

function matchColumn(headers, aliases) {
  const hs = headers.map((h) => ({ raw: h, n: normalizeHeader(h) }));
  const as = aliases.map((a) => normalizeHeader(a));

  for (const a of as) {
    const exact = hs.find((h) => h.n === a);
    if (exact) return exact.raw;
  }
  for (const a of as) {
    const fuzzy = hs.find((h) => h.n.includes(a) || a.includes(h.n));
    if (fuzzy) return fuzzy.raw;
  }
  return null;
}

function buildMapping(headers) {
  const map = {};
  Object.entries(requiredAliases).forEach(([k, aliases]) => {
    map[k] = matchColumn(headers, aliases);
  });
  return map;
}

function mappingScore(mapping) {
  const required = ["country", "ladder", "gdp", "soc", "life", "free", "gen", "corr"];
  return required.filter((k) => !!mapping[k]).length;
}

function pickBestDataSheet(workbook) {
  let best = null;
  workbook.SheetNames.forEach((name) => {
    const ws = workbook.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: null });
    if (!rows.length) return;
    const mapping = buildMapping(Object.keys(rows[0]));
    const score = mappingScore(mapping);
    if (!best || score > best.score) {
      best = { name, rows, mapping, score };
    }
  });
  return best;
}

function extractRank2Estimations(workbook) {
  const rank2Name = workbook.SheetNames.find((n) => normalizeHeader(n) === "rank2");
  if (!rank2Name) return new Map();
  const ws = workbook.Sheets[rank2Name];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  if (!rows.length) return new Map();

  let headerRow = -1;
  let countryCol = -1;
  let estCol = -1;
  for (let r = 0; r < rows.length; r += 1) {
    const row = rows[r];
    if (!Array.isArray(row)) continue;
    const normalized = row.map((x) => normalizeHeader(x));
    const cIdx = normalized.findIndex((v) => v === "countryattributename" || v === "countryname");
    const eIdx = normalized.findIndex((v) => v === "est" || v === "estimation");
    if (cIdx >= 0 && eIdx >= 0) {
      headerRow = r;
      countryCol = cIdx;
      estCol = eIdx;
      break;
    }
  }
  if (headerRow < 0) return new Map();

  const out = new Map();
  for (let r = headerRow + 1; r < rows.length; r += 1) {
    const row = rows[r];
    if (!Array.isArray(row)) continue;
    const country = row[countryCol] === null || row[countryCol] === undefined ? "" : String(row[countryCol]).trim();
    const est = parseNum(row[estCol]);
    if (!country || est === null) continue;
    out.set(country, est);
  }
  return out;
}

function rankEqDescending(rows, field) {
  const points = rows
    .map((r, i) => ({ i, v: r[field] }))
    .filter((x) => x.v !== null)
    .sort((a, b) => b.v - a.v);

  const ranks = new Array(rows.length).fill(null);
  let lastVal = null;
  let lastRank = null;
  points.forEach((p, idx) => {
    const rank = idx + 1;
    if (lastVal !== null && p.v === lastVal) {
      ranks[p.i] = lastRank;
    } else {
      ranks[p.i] = rank;
      lastVal = p.v;
      lastRank = rank;
    }
  });
  return ranks;
}

function rankEqAscending(rows, field) {
  const points = rows
    .map((r, i) => ({ i, v: r[field] }))
    .filter((x) => x.v !== null)
    .sort((a, b) => a.v - b.v);

  const ranks = new Array(rows.length).fill(null);
  let lastVal = null;
  let lastRank = null;
  points.forEach((p, idx) => {
    const rank = idx + 1;
    if (lastVal !== null && p.v === lastVal) {
      ranks[p.i] = lastRank;
    } else {
      ranks[p.i] = rank;
      lastVal = p.v;
      lastRank = rank;
    }
  });
  return ranks;
}

function pearson(rows, xField, yField) {
  const pairs = rows
    .filter((r) => r[xField] !== null && r[yField] !== null)
    .map((r) => [r[xField], r[yField]]);
  const N = pairs.length;
  if (N < 2) return { r: null, N };

  const mx = pairs.reduce((s, p) => s + p[0], 0) / N;
  const my = pairs.reduce((s, p) => s + p[1], 0) / N;
  let num = 0;
  let dx2 = 0;
  let dy2 = 0;
  pairs.forEach(([x, y]) => {
    const dx = x - mx;
    const dy = y - my;
    num += dx * dy;
    dx2 += dx * dx;
    dy2 += dy * dy;
  });
  const den = Math.sqrt(dx2 * dy2);
  return { r: den === 0 ? null : num / den, N };
}

function format3(v) {
  return v === null ? "" : Number(v).toFixed(3);
}

function format2(v) {
  return v === null ? "" : Number(v).toFixed(2);
}

function format4(v) {
  return v === null ? "NA" : Number(v).toFixed(4);
}

function buildCocoPayloadTexts() {
  // Match backend/COCO-safe formatting for manual paste.
  state.matrixText = state.matrix.map((row) => row.join(" ")).join("\r");
  state.objectNamesText = state.cocoRows.map((r) => r.country).join("\n");
  state.attributeNamesText = [
    "Ladder score",
    "Explained by: Log GDP per capita",
    "Explained by: Social support",
    "Explained by: Healthy life expectancy",
    "Explained by: Freedom to make life choices",
    "Explained by: Generosity",
    "Explained by: Perceptions of corruption",
    "Dystopia + residual",
    "Y"
  ].join("\n");
}

function renderCocoPayloadTexts() {
  el.matrixTextArea.value = state.matrixText;
  el.objectNamesTextArea.value = state.objectNamesText;
  el.attributeNamesTextArea.value = state.attributeNamesText;
}

function buildRankedCocoMatrix() {
  // Raw numeric columns are converted to Excel-style RANK.EQ before COCO.
  // COCO receives only rank values + constant Y=1000.
  const ladRank = rankEqDescending(state.cocoRows, "ladder");
  const gdpRank = rankEqDescending(state.cocoRows, "gdp");
  const socRank = rankEqDescending(state.cocoRows, "soc");
  const lifeRank = rankEqDescending(state.cocoRows, "life");
  const freeRank = rankEqDescending(state.cocoRows, "free");
  const genRank = rankEqDescending(state.cocoRows, "gen");
  // Rank2-compatible handling for corruption.
  const corrRank = rankEqAscending(state.cocoRows, "corr");
  const dystRank = rankEqDescending(state.cocoRows, "dystopia");

  state.matrix = state.cocoRows.map((_r, i) => [
    ladRank[i],
    gdpRank[i],
    socRank[i],
    lifeRank[i],
    freeRank[i],
    genRank[i],
    corrRank[i],
    dystRank[i],
    1000
  ]);
}

function renderMapping() {
  const labels = {
    country: "Country name",
    ladder: "Ladder score",
    gdp: "Explained by: Log GDP per capita",
    soc: "Explained by: Social support",
    life: "Explained by: Healthy life expectancy",
    free: "Explained by: Freedom to make life choices",
    gen: "Explained by: Generosity",
    corr: "Explained by: Perceptions of corruption",
    dystopia: "Dystopia + residual"
  };
  el.mappingBody.innerHTML = "";
  Object.keys(labels).forEach((k) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${labels[k]}</td><td>${state.mapping[k] || "NOT FOUND"}</td>`;
    el.mappingBody.appendChild(tr);
  });
}

function renderRawTable() {
  el.rawBody.innerHTML = "";
  state.rows.forEach((r) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${r.country}</td>
      <td>${format3(r.ladder)}</td>
      <td>${format3(r.gdp)}</td>
      <td>${format3(r.soc)}</td>
      <td>${format3(r.life)}</td>
      <td>${format3(r.free)}</td>
      <td>${format3(r.gen)}</td>
      <td>${format3(r.corr)}</td>
      <td>${format3(r.dystopia)}</td>
    `;
    el.rawBody.appendChild(tr);
  });
}

function renderMatrixTable() {
  el.matrixBody.innerHTML = "";
  state.cocoRows.forEach((r, i) => {
    const m = state.matrix[i] || [];
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${r.country}</td>
      <td>${m[0] ?? ""}</td>
      <td>${m[1] ?? ""}</td>
      <td>${m[2] ?? ""}</td>
      <td>${m[3] ?? ""}</td>
      <td>${m[4] ?? ""}</td>
      <td>${m[5] ?? ""}</td>
      <td>${m[6] ?? ""}</td>
      <td>${m[7] ?? ""}</td>
      <td>${m[8] ?? ""}</td>
    `;
    el.matrixBody.appendChild(tr);
  });
}

function renderResults() {
  el.resultBody.innerHTML = "";
  const q = (el.countrySearch?.value || "").trim().toLowerCase();
  const filtered = state.results
    .map((r, i) => ({ r, i }))
    .filter((x) => !q || x.r.country.toLowerCase().includes(q));

  filtered.forEach(({ r, i }) => {
    const tr = document.createElement("tr");
    tr.setAttribute("data-row-index", String(i));
    tr.innerHTML = `
      <td>${r.country}</td>
      <td>${format3(r.ladder)}</td>
      <td>${format3(r.gdp)}</td>
      <td>${format3(r.soc)}</td>
      <td>${format3(r.life)}</td>
      <td>${format3(r.free)}</td>
      <td>${format3(r.gen)}</td>
      <td>${format3(r.corr)}</td>
      <td>${format3(r.dystopia)}</td>
      <td>${r.naive2Score ?? ""}</td>
      <td>${r.naive2Rank ?? ""}</td>
      <td>${format3(r.naive1Score)}</td>
      <td>${r.naive1Rank ?? ""}</td>
      <td>${r.objectiveRank ?? ""}</td>
      <td>${r.delta1 ?? ""}</td>
      <td>${r.delta2 ?? ""}</td>
      <td>${format3(r.estimation)}</td>
      <td>1000</td>
      <td>${format3(r.cocoDelta)}</td>
      <td>${format2(r.cocoDeltaOverTruth)}</td>
      <td><button type="button" class="explain-btn" data-row-index="${i}">Explain</button></td>
    `;
    el.resultBody.appendChild(tr);
  });

  el.resultBody.querySelectorAll(".explain-btn").forEach((btn) => {
    btn.addEventListener("click", (ev) => {
      ev.stopPropagation();
      const idx = Number(btn.getAttribute("data-row-index"));
      if (Number.isFinite(idx) && state.results[idx]) openCountryDrawer(state.results[idx]);
    });
  });

  el.resultBody.querySelectorAll("tr[data-row-index]").forEach((tr) => {
    tr.addEventListener("dblclick", () => {
      const idx = Number(tr.getAttribute("data-row-index"));
      if (Number.isFinite(idx) && state.results[idx]) openCountryDrawer(state.results[idx]);
    });
  });
}

function renderValidationAndCorrelations() {
  const c1 = pearson(state.results, "naive2Rank", "objectiveRank");
  const c2 = pearson(state.results, "ladder", "estimation");
  el.correlations.innerHTML = "";
  [
    { label: "corrRanks = corr(naive2Rank, objectiveRank)", data: c1 },
    { label: "corrScores = corr(LadderScore, Estimation)", data: c2 }
  ].forEach((item) => {
    const div = document.createElement("div");
    div.className = "corr-item";
    div.innerHTML = `<div>${item.label}</div><div>r = ${format4(item.data.r)}</div><div>N = ${item.data.N}</div>`;
    el.correlations.appendChild(div);
  });
}

function quantileBuckets(values) {
  const sorted = [...values].sort((a, b) => a - b);
  if (!sorted.length) return { q1: 0, q2: 0 };
  const q1 = sorted[Math.floor((sorted.length - 1) * (1 / 3))];
  const q2 = sorted[Math.floor((sorted.length - 1) * (2 / 3))];
  return { q1, q2 };
}

function quantileBuckets5(values) {
  const sorted = [...values].sort((a, b) => a - b);
  if (!sorted.length) return { q1: 0, q2: 0, q3: 0, q4: 0 };
  const pick = (p) => sorted[Math.floor((sorted.length - 1) * p)];
  return { q1: pick(0.2), q2: pick(0.4), q3: pick(0.6), q4: pick(0.8) };
}

function normCountry(value) {
  return String(value || "").toLowerCase().replace(/[^a-z]/g, "");
}

const PLOTLY_COUNTRY_ALIAS = {
  "taiwanprovinceofchina": "Taiwan",
  "unitedstates": "United States",
  "unitedarabemirates": "United Arab Emirates",
  "czechia": "Czech Republic",
  "vietnam": "Vietnam",
  "russianfederation": "Russia"
};

function toPlotlyCountryName(country) {
  const key = normCountry(country);
  return PLOTLY_COUNTRY_ALIAS[key] || country;
}

function getBucketInfo(rows, field, mode) {
  const vals = rows
    .filter((r) => Number.isFinite(r[field]))
    .map((r) => (mode === "delta_abs" ? Math.abs(r[field]) : r[field]));
  return quantileBuckets5(vals);
}

function bucketClass(value, mode, q1, q2, q3, q4) {
  if (!Number.isFinite(value)) return "na";
  if (mode === "high_good") {
    const v = value;
    if (v >= q4) return "best";
    if (v >= q3) return "second";
    if (v >= q2) return "mid";
    if (v >= q1) return "low";
    return "worst";
  }
  if (mode === "low_good") {
    const v = value;
    if (v <= q1) return "best";
    if (v <= q2) return "second";
    if (v <= q3) return "mid";
    if (v <= q4) return "low";
    return "worst";
  }
  const v = Math.abs(value);
  if (v <= q1) return "best";
  if (v <= q2) return "second";
  if (v <= q3) return "mid";
  if (v <= q4) return "low";
  return "worst";
}

async function renderWorldMaps() {
  const defs = [
    { title: "1) Ladder score map", field: "ladder", mode: "high_good" },
    { title: "2) Naive1 map (more is better)", field: "naive1Score", mode: "high_good" },
    { title: "3) Naive2 map (less is better)", field: "naive2Score", mode: "low_good" },
    { title: "4) Delta1 map (more is better)", field: "delta1", mode: "high_good" },
    { title: "5) Delta2 map (more is better)", field: "delta2", mode: "high_good" }
  ];

  const colors = {
    best: "#86efac",   // light green
    second: "#15803d", // dark green
    mid: "#facc15",    // yellow
    low: "#f97316",    // orange
    worst: "#dc2626",  // red
    na: "#cbd5e1"      // gray no data
  };

  el.worldMaps.innerHTML = "";
  defs.forEach((def, idx) => {
    const card = document.createElement("div");
    card.className = "world-map-card";
    const plotId = `world-map-${idx}`;
      card.innerHTML = `
        <h3>${def.title}</h3>
      <div id="${plotId}" class="world-map-plot"></div>
      <div class="legend">
        <span><i class="swatch" style="background:${colors.best}"></i>Light green (best)</span>
        <span><i class="swatch" style="background:${colors.second}"></i>Dark green (2nd)</span>
        <span><i class="swatch" style="background:${colors.mid}"></i>Yellow</span>
        <span><i class="swatch" style="background:${colors.low}"></i>Orange</span>
        <span><i class="swatch" style="background:${colors.worst}"></i>Red (worst)</span>
        <span><i class="swatch" style="background:${colors.na}"></i>No data</span>
      </div>
    `;
    el.worldMaps.appendChild(card);

    const info = getBucketInfo(state.results, def.field, def.mode);
    const countries = [];
    const z = [];
    const text = [];
    state.results.forEach((r) => {
      const cls = bucketClass(r[def.field], def.mode, info.q1, info.q2, info.q3, info.q4);
      const code =
        cls === "best" ? 5 :
        cls === "second" ? 4 :
        cls === "mid" ? 3 :
        cls === "low" ? 2 :
        cls === "worst" ? 1 : 0;
      countries.push(toPlotlyCountryName(r.country));
      z.push(code);
      text.push(`${r.country}: ${Number.isFinite(r[def.field]) ? r[def.field] : "no data"}`);
    });

    Plotly.newPlot(
      plotId,
      [
        {
          type: "choropleth",
          locationmode: "country names",
          locations: countries,
          z,
          text,
          hovertemplate: "%{text}<extra></extra>",
          zmin: 0,
          zmax: 5,
          colorscale: [
            [0.00, colors.na],
            [0.16, colors.na],
            [0.17, colors.worst],
            [0.33, colors.worst],
            [0.34, colors.low],
            [0.50, colors.low],
            [0.51, colors.mid],
            [0.67, colors.mid],
            [0.68, colors.second],
            [0.84, colors.second],
            [0.85, colors.best],
            [1.00, colors.best]
          ],
          marker: { line: { color: "#64748b", width: 0.4 } },
          showscale: false
        }
      ],
      {
        margin: { l: 0, r: 0, t: 0, b: 0 },
        geo: {
          projection: { type: "natural earth" },
          showframe: false,
          showcoastlines: false,
          showland: true,
          landcolor: colors.na,
          showcountries: true,
          countrycolor: "#94a3b8",
          bgcolor: "#f7fbff"
        },
        paper_bgcolor: "#f7fbff"
      },
      { displayModeBar: false, responsive: true }
    );
  });
}

function buildExplanation(row) {
  return [
    `naive2 = ladRank + gdpRank + socRank + lifeRank + freeRank + genRank + corrRank + dystRank`,
    `naive2 = ${row.ladRank ?? "NA"} + ${row.gdpRank ?? "NA"} + ${row.socRank ?? "NA"} + ${row.lifeRank ?? "NA"} + ${row.freeRank ?? "NA"} + ${row.genRank ?? "NA"} + ${row.corrRank ?? "NA"} + ${row.dystRank ?? "NA"} = ${row.naive2Score ?? "NA"}`,
    `naive2_rank = RANK.EQ(naive2, asc) = ${row.naive2Rank ?? "NA"}`,
    `naive1 = AVG(gdp, soc, life, free, gen, corr) = ${format3(row.naive1Score) || "NA"}`,
    `naive1_rank = RANK.EQ(naive1, desc) = ${row.naive1Rank ?? "NA"}`,
    `objectiveRank = RANK.EQ(estimation, desc) = ${row.objectiveRank ?? "NA"}`,
    `delta1 = naive2_rank - objectiveRank = ${row.naive2Rank ?? "NA"} - ${row.objectiveRank ?? "NA"} = ${row.delta1 ?? "NA"}`,
    `delta2 = naive1_rank - objectiveRank = ${row.naive1Rank ?? "NA"} - ${row.objectiveRank ?? "NA"} = ${row.delta2 ?? "NA"}`,
    `COCO_Delta = 1000 - estimation = ${format3(row.cocoDelta) || "NA"}`,
    `COCO_Delta/Teny (%) = ${format2(row.cocoDeltaOverTruth) || "NA"}`
  ].join("\n");
}

function openCountryDrawer(row) {
  if (!el.countryDrawer || !el.drawerTitle || !el.drawerBody) return;
  el.drawerTitle.textContent = row.country;
  el.drawerBody.innerHTML = `
    <div class="drawer-grid">
      <div><strong>Ladder:</strong> ${format3(row.ladder) || "NA"}</div>
      <div><strong>Estimation:</strong> ${format3(row.estimation) || "NA"}</div>
      <div><strong>naive2 / rank:</strong> ${row.naive2Score ?? "NA"} / ${row.naive2Rank ?? "NA"}</div>
      <div><strong>naive1 / rank:</strong> ${format3(row.naive1Score) || "NA"} / ${row.naive1Rank ?? "NA"}</div>
      <div><strong>objectiveRank:</strong> ${row.objectiveRank ?? "NA"}</div>
      <div><strong>delta1 / delta2:</strong> ${row.delta1 ?? "NA"} / ${row.delta2 ?? "NA"}</div>
      <div><strong>Status:</strong> ${row.status}</div>
    </div>
    <h4>Calculation breakdown</h4>
    <pre class="drawer-pre">${buildExplanation(row)}</pre>
  `;
  el.countryDrawer.classList.remove("hidden-drawer");
  el.countryDrawer.setAttribute("aria-hidden", "false");
}

function closeCountryDrawer() {
  if (!el.countryDrawer) return;
  el.countryDrawer.classList.add("hidden-drawer");
  el.countryDrawer.setAttribute("aria-hidden", "true");
}

async function exportMapPng(plotNode, fileName) {
  const dataUrl = await Plotly.toImage(plotNode, {
    format: "png",
    width: 1400,
    height: 800,
    scale: 2
  });
  const a = document.createElement("a");
  a.href = dataUrl;
  a.download = fileName;
  a.click();
}

async function exportAllMapsPng() {
  const maps = Array.from(document.querySelectorAll(".world-map-plot"));
  if (!maps.length) {
    alert("No rendered maps yet. Generate results first.");
    return;
  }
  for (let i = 0; i < maps.length; i += 1) {
    await exportMapPng(maps[i], `whr_map_${i + 1}.png`);
    await new Promise((resolve) => setTimeout(resolve, 120));
  }
}

function computeResultsFromEstimations(estimationMap) {
  const rows = state.rows.map((r) => ({
    country: r.country,
    ladder: r.ladder,
    gdp: r.gdp,
    soc: r.soc,
    life: r.life,
    free: r.free,
    gen: r.gen,
    corr: r.corr,
    dystopia: r.dystopia,
    estimation: estimationMap.has(r.country) ? estimationMap.get(r.country) : null,
    ladRank: null,
    gdpRank: null,
    socRank: null,
    lifeRank: null,
    freeRank: null,
    genRank: null,
    corrRank: null,
    dystRank: null,
    naive2Score: null,
    naive2Rank: null,
    naive1Score: null,
    naive1Rank: null,
    objectiveRank: null,
    delta1: null,
    delta2: null,
    status: "NO DATA"
  }));

  const ladRank = rankEqDescending(rows, "ladder");
  const gdpRank = rankEqDescending(rows, "gdp");
  const socRank = rankEqDescending(rows, "soc");
  const lifeRank = rankEqDescending(rows, "life");
  const freeRank = rankEqDescending(rows, "free");
  const genRank = rankEqDescending(rows, "gen");
  // Rank2 compatibility: corruption is ranked ascending (lower is better).
  const corrRank = rankEqAscending(rows, "corr");
  const dystRank = rankEqDescending(rows, "dystopia");

  rows.forEach((r, i) => {
    r.ladRank = ladRank[i];
    r.gdpRank = gdpRank[i];
    r.socRank = socRank[i];
    r.lifeRank = lifeRank[i];
    r.freeRank = freeRank[i];
    r.genRank = genRank[i];
    r.corrRank = corrRank[i];
    r.dystRank = dystRank[i];

    const rankFields = [r.ladRank, r.gdpRank, r.socRank, r.lifeRank, r.freeRank, r.genRank, r.corrRank, r.dystRank];
    r.naive2Score = rankFields.every((v) => v !== null) ? rankFields.reduce((a, b) => a + b, 0) : null;

    const avgFields = [r.gdp, r.soc, r.life, r.free, r.gen, r.corr];
    r.naive1Score = avgFields.every((v) => v !== null) ? round3(avgFields.reduce((a, b) => a + b, 0) / 6) : null;
  });

  const naive2Ranks = rankEqAscending(rows, "naive2Score");
  const naive1Ranks = rankEqDescending(rows, "naive1Score");
  const objectiveRanks = rankEqDescending(rows, "estimation");

  rows.forEach((r, i) => {
    r.naive2Rank = naive2Ranks[i];
    r.naive1Rank = naive1Ranks[i];
    r.objectiveRank = objectiveRanks[i];

    if (r.naive2Rank !== null && r.objectiveRank !== null) r.delta1 = r.naive2Rank - r.objectiveRank;
    if (r.naive1Rank !== null && r.objectiveRank !== null) r.delta2 = r.naive1Rank - r.objectiveRank;
    if (r.estimation !== null) {
      r.cocoDelta = 1000 - r.estimation;
      r.cocoDeltaOverTruth = (r.cocoDelta / 1000) * 100;
    } else {
      r.cocoDelta = null;
      r.cocoDeltaOverTruth = null;
    }

    r.status = r.ladder !== null && r.estimation !== null ? "OK" : "NO DATA";
  });

  state.results = rows;
  renderResults();
  renderValidationAndCorrelations();
  renderWorldMaps().catch(() => {
    // Keep text maps if geo rendering fails.
  });
}

function parseManualEstimations(text, expectedCount) {
  const out = [];
  text
    .split(/\r?\n/)
    .map((x) => x.trim())
    .filter(Boolean)
    .forEach((line) => {
      const nums = line.match(/-?\d+(?:[.,]\d+)?/g);
      if (!nums) return;
      nums.forEach((n) => {
        const v = Number(n.replace(",", "."));
        if (Number.isFinite(v)) out.push(v);
      });
    });
  return out.slice(0, expectedCount);
}

function initTabs() {
  const screenCards = Array.from(document.querySelectorAll("[data-screen]"));
  const tabButtons = Array.from(document.querySelectorAll("[data-screen-btn]"));
  const helpBtn = document.getElementById("floatingHelpBtn");
  const helpBox = document.getElementById("floatingHelpText");
  const helpText = {
    upload: "Upload: Select your WHR Excel file. This section detects columns and shows your cleaned raw data preview.",
    coco: "COCO Input: Shows the ranked matrix, object names, and attribute names sent to COCO Y0. Use fallback paste here if auto-run fails.",
    results: "Results: Displays COCO estimation and Rank2-style outputs (naive1/naive2 ranks, objective rank, delta1, delta2) plus correlations.",
    maps: "Maps: Displays 5 colored world maps (Ladder, Naive1, Naive2, Delta1, Delta2) with green/yellow/orange/red performance colors."
  };
  let currentScreen = "upload";

  const setScreen = (screen) => {
    currentScreen = screen;
    screenCards.forEach((card) => {
      card.classList.toggle("hidden-screen", card.getAttribute("data-screen") !== screen);
    });
    tabButtons.forEach((btn) => {
      btn.classList.toggle("active", btn.getAttribute("data-screen-btn") === screen);
    });
    // Plotly maps may need resize after panel becomes visible.
    setTimeout(() => {
      document.querySelectorAll(".world-map-plot").forEach((node) => {
        if (window.Plotly) Plotly.Plots.resize(node);
      });
    }, 60);
    if (helpBox && !helpBox.classList.contains("hidden-help")) {
      helpBox.textContent = helpText[currentScreen] || "No help text available.";
    }
  };

  const toggleHelp = () => {
    if (!helpBox) return;
    if (!helpBox.classList.contains("hidden-help")) {
      helpBox.classList.add("hidden-help");
      helpBox.textContent = "";
      return;
    }
    helpBox.textContent = helpText[currentScreen] || "No help text available.";
    helpBox.classList.remove("hidden-help");
  };

  tabButtons.forEach((btn) => {
    btn.addEventListener("click", () => setScreen(btn.getAttribute("data-screen-btn")));
  });
  if (helpBtn) helpBtn.addEventListener("click", toggleHelp);
  setScreen("upload");
}

async function callCocoProxy() {
  const payload = {
    matrix: state.matrix,
    objectNames: state.cocoRows.map((r) => r.country),
    attributeNames: [
      "Ladder score",
      "Explained by: Log GDP per capita",
      "Explained by: Social support",
      "Explained by: Healthy life expectancy",
      "Explained by: Freedom to make life choices",
      "Explained by: Generosity",
      "Explained by: Perceptions of corruption",
      "Dystopia + residual",
      "Y"
    ]
  };
  try {
    const res = await fetch("/api/coco-y0", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });
    return res.json();
  } catch (err) {
    return {
      ok: false,
      automated: false,
      message: `Backend not reachable (${err.message}). Start server and open app from http://localhost:3000.`,
      estimations: []
    };
  }
}

function exportTextFile(filename, content) {
  const blob = new Blob([content], { type: "text/plain;charset=utf-8;" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
}

function exportCsv() {
  if (!state.results.length) return;
  const header = [
    "COCO:Y0",
    "Ladder score",
    "Explained by: Log GDP per capita",
    "Explained by: Social support",
    "Explained by: Healthy life expectancy",
    "Explained by: Freedom to make life choices",
    "Explained by: Generosity",
    "Explained by: Perceptions of corruption",
    "Dystopia + residual",
    "naive2",
    "naive2_rank",
    "naive1",
    "naive1_rank",
    "objectiveRank",
    "delta1",
    "delta2",
    "Becslés",
    "Tény+0",
    "COCO_Delta",
    "COCO_Delta/Tény"
  ];
  const lines = [header.join(",")];
  state.results.forEach((r) => {
    const row = [
      `"${String(r.country).replace(/"/g, "\"\"")}"`,
      r.ladder ?? "",
      r.gdp ?? "",
      r.soc ?? "",
      r.life ?? "",
      r.free ?? "",
      r.gen ?? "",
      r.corr ?? "",
      r.dystopia ?? "",
      r.naive2Score ?? "",
      r.naive2Rank ?? "",
      r.naive1Score ?? "",
      r.naive1Rank ?? "",
      r.objectiveRank ?? "",
      r.delta1 ?? "",
      r.delta2 ?? "",
      r.estimation ?? "",
      1000,
      r.cocoDelta ?? "",
      r.cocoDeltaOverTruth ?? ""
    ];
    lines.push(row.join(","));
  });
  exportTextFile("whr_oam_coco_results.csv", lines.join("\n"));
}

function exportXlsx() {
  if (!state.results.length) return;
  const rows = state.results.map((r) => ({
    "COCO:Y0": r.country,
    "Ladder score": r.ladder,
    "Explained by: Log GDP per capita": r.gdp,
    "Explained by: Social support": r.soc,
    "Explained by: Healthy life expectancy": r.life,
    "Explained by: Freedom to make life choices": r.free,
    "Explained by: Generosity": r.gen,
    "Explained by: Perceptions of corruption": r.corr,
    "Dystopia + residual": r.dystopia,
    naive2: r.naive2Score,
    naive2_rank: r.naive2Rank,
    naive1: r.naive1Score,
    naive1_rank: r.naive1Rank,
    objectiveRank: r.objectiveRank,
    delta1: r.delta1,
    delta2: r.delta2,
    "Becslés": r.estimation,
    "Tény+0": 1000,
    "COCO_Delta": r.cocoDelta,
    "COCO_Delta/Tény": r.cocoDeltaOverTruth
  }));
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Results");
  XLSX.writeFile(wb, "whr_oam_coco_results.xlsx");
}

async function processFile() {
  const file = el.fileInput.files[0];
  if (!file) {
    alert("Please choose an Excel file first.");
    return;
  }
  el.uploadStatus.textContent = "Reading Excel...";

  const reader = new FileReader();
  reader.onload = async (e) => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: "array" });
      const best = pickBestDataSheet(wb);
      if (!best || best.score < 7) {
        el.uploadStatus.textContent = "Could not find a valid raw data sheet with required columns.";
        return;
      }
      state.mapping = best.mapping;
      renderMapping();
      el.uploadStatus.textContent = `Detected input sheet: ${best.name}`;

      state.rows = best.rows
        .map((row) => {
          const countryRaw = state.mapping.country ? row[state.mapping.country] : null;
          const country = countryRaw === null || countryRaw === undefined ? "" : String(countryRaw).trim();
          if (!country) return null;
          return {
            country,
            ladder: parseNum(state.mapping.ladder ? row[state.mapping.ladder] : null),
            gdp: parseNum(state.mapping.gdp ? row[state.mapping.gdp] : null),
            soc: parseNum(state.mapping.soc ? row[state.mapping.soc] : null),
            life: parseNum(state.mapping.life ? row[state.mapping.life] : null),
            free: parseNum(state.mapping.free ? row[state.mapping.free] : null),
            gen: parseNum(state.mapping.gen ? row[state.mapping.gen] : null),
            corr: parseNum(state.mapping.corr ? row[state.mapping.corr] : null),
            dystopia: parseNum(state.mapping.dystopia ? row[state.mapping.dystopia] : null)
          };
        })
        .filter(Boolean);

      renderRawTable();

      // Rank2-compatible COCO subset: rows with complete 8 ranked attributes.
      state.cocoRows = state.rows.filter((r) =>
        [r.ladder, r.gdp, r.soc, r.life, r.free, r.gen, r.corr, r.dystopia].every((v) => v !== null)
      );

      buildRankedCocoMatrix();
      buildCocoPayloadTexts();
      renderMatrixTable();
      renderCocoPayloadTexts();
      el.uploadStatus.textContent = "Raw data ranked successfully. Sending ranked matrix to COCO Y0 proxy...";

      // If the workbook contains Rank2 with estimations, use it directly to match Excel exactly.
      const rank2Estimations = extractRank2Estimations(wb);
      if (rank2Estimations.size > 0) {
        computeResultsFromEstimations(rank2Estimations);
        el.uploadStatus.textContent = `Processed. Used Rank2 sheet estimations from workbook (${rank2Estimations.size} rows).`;
        return;
      }

      const cocoResponse = await callCocoProxy();
      if (!cocoResponse.ok) {
        el.uploadStatus.textContent = `COCO proxy error: ${cocoResponse.message}`;
        return;
      }

      if (!cocoResponse.automated) {
        state.matrixText = cocoResponse.matrixText || state.matrixText;
        renderCocoPayloadTexts();
        el.uploadStatus.textContent = `${cocoResponse.message} Use manual fallback.`;
        return;
      }

      const estimationMap = new Map();
      state.cocoRows.forEach((r, i) => {
        const est = parseNum(cocoResponse.estimations[i]);
        if (est !== null && est > 0 && est < 10000) estimationMap.set(r.country, est);
      });
      if (!estimationMap.size) {
        computeResultsFromEstimations(new Map());
        el.uploadStatus.textContent = "COCO response did not contain valid estimations. Use manual fallback paste.";
        return;
      }
      computeResultsFromEstimations(estimationMap);
      el.uploadStatus.textContent = "COCO automation completed.";
    } catch (err) {
      el.uploadStatus.textContent = `Error: ${err.message}`;
    }
  };
  reader.readAsArrayBuffer(file);
}

el.processBtn.addEventListener("click", processFile);
el.exportMatrixBtn.addEventListener("click", () => {
  const text = state.matrixText || "";
  exportTextFile("coco_rank_matrix.txt", text);
});
el.applyManualBtn.addEventListener("click", () => {
  if (!state.cocoRows.length) {
    alert("No COCO matrix rows available.");
    return;
  }
  const vals = parseManualEstimations(el.manualEstimationInput.value, state.cocoRows.length);
  if (!vals.length) {
    alert("No numeric estimations found.");
    return;
  }
  const estimationMap = new Map();
  state.cocoRows.forEach((r, i) => {
    const v = i < vals.length ? vals[i] : null;
    if (v !== null) estimationMap.set(r.country, v);
  });
  computeResultsFromEstimations(estimationMap);
  el.uploadStatus.textContent = "Manual estimations applied.";
});
el.exportCsvBtn.addEventListener("click", exportCsv);
el.exportXlsxBtn.addEventListener("click", exportXlsx);
if (el.countrySearch) {
  el.countrySearch.addEventListener("input", renderResults);
}
if (el.clearSearchBtn) {
  el.clearSearchBtn.addEventListener("click", () => {
    if (el.countrySearch) el.countrySearch.value = "";
    renderResults();
  });
}
if (el.closeDrawerBtn) {
  el.closeDrawerBtn.addEventListener("click", closeCountryDrawer);
}
if (el.exportAllMapsBtn) {
  el.exportAllMapsBtn.addEventListener("click", () => {
    exportAllMapsPng().catch((err) => {
      alert(`Map export failed: ${err.message}`);
    });
  });
}
initTabs();
