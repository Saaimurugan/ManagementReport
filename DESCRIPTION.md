# Description

**Technology Leadership Dashboard** — a single-file, offline HTML dashboard that turns a JIRA-style work-item export into an executive view of delivery health.

It presents an at-a-glance Executive Summary (On Track / At Risk / Critical), key metrics, and a per-project health table, plus two drill-down views: by hierarchy (Project → Epic → Story → Task/Bug) and by Sprint. Every level shows three rolled-up RAG health indicators — Efforts, Timeline, and Quality — where Quality is measured as open bugs per ticket against a configurable allowance.

The dashboard runs entirely in the browser with no server or dependencies: `dashboard_standalone.html` embeds the dataset directly. Data flows from `JIRA.xlsx` → `jira_data.json` (via `generate_dashboard_data.py`) → `dashboard_standalone.html` (via `create_standalone_dashboard.py`).
