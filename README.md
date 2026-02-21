WHR OAM + COCO Y0 Dashboard

This project is a web-based implementation of the WHR (World Happiness Report) + OAM workflow integrated with COCO Y0 estimation.

It replaces the original Excel-based pipeline with a structured, reproducible system that standardizes preprocessing, ranking, matrix generation, estimation, and post-analysis.

The goal is to preserve the analytical logic of the Excel model while improving transparency, repeatability, and methodological control.

Purpose

The traditional workflow required:

Manual ranking using Excel functions

Manual construction of the COCO input matrix

Copy-pasting data into COCO

Manual recalculation of indicators and deltas

This dashboard automates that entire process while keeping every transformation explicit and traceable.

It enables consistent analysis across different WHR datasets without changing the underlying logic.

Core Workflow

Upload WHR raw Excel data

Automatic sheet detection and fuzzy column matching

Numeric normalization (including decimal comma handling)

Competition ranking (Excel RANK.EQ logic)

COCO input matrix generation

COCO Y0 estimation (automatic via backend proxy or manual fallback)

Post-estimation indicator computation

Correlation validation

Table and map visualization

Export of analytical results

Analytical Outputs

After estimation, the system computes:

Objective rank (based on COCO estimation)

Naive1 score and rank (average of explained-by indicators)

Naive2 score and rank (sum of ranked attributes)

Delta1 and Delta2 (distance from objective rank)

COCO gap indicators

Pearson correlation coefficients

This reproduces the same analytical structure used in the Excel pipeline, but in a consistent web-based environment.

Architecture
Frontend

HTML

CSS

Vanilla JavaScript

SheetJS (Excel parsing)

Responsible for:

Data ingestion

Ranking logic

Matrix construction

Visualization

Export functionality

Backend

Node.js

Express

Acts as a proxy layer for COCO Y0 to bypass browser-level CORS restrictions.

Research Value

This implementation:

Standardizes preprocessing rules

Eliminates manual spreadsheet errors

Ensures raw → rank → matrix → estimation traceability

Supports repeatable cross-dataset experimentation

Makes analytical assumptions explicit

It is designed as a methodological tool rather than a general consumer application.

Running Locally
git clone https://github.com/Nurboel/whr-oam-coco-dashboard.git
cd whr-oam-coco-dashboard
npm install
node server.js

Then open the frontend in your browser.

License

MIT License
