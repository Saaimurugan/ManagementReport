# Technology Leadership Dashboard 

A self-contained, single-file HTML dashboard for monitoring JIRA-style project delivery across projects, epics, stories, sprints, and tasks. It turns a flat export of work items into an executive-friendly view with health indicators, drill-down tables, and charts — no web server, database, or build tooling required.

Open `dashboard_standalone.html` in any modern browser and everything works offline; the data is embedded directly in the file.

## Highlights

- **One-glance Executive Summary** with an overall RAG (Red / Amber / Green) verdict and counts of On Track, At Risk, and Critical projects.
- **Three-metric health model — Efforts, Timeline, Quality (E / T / Q)** — shown as colored chips at every level and rolled up from tasks to stories, epics, projects, and sprints.
- **Detailed Task View**: drill down Project → Epic → Story → Task/Bug.
- **Sprint View**: the same drill-down grouped by Sprint → Project → Epic → Story → Task/Bug.
- **Project Health Status** table with per-project Efforts, Timeline, and Quality status.
- **Interactive filters** (Project, Status, Priority, Quality/bugs, free-text search) and charts.
- **Fully offline**: the standalone file embeds the data, so it can be emailed or opened from disk.

## Files

| File | Purpose |
| --- | --- |
| `dashboard_standalone.html` | **The deliverable.** Open this in a browser. Data is embedded, works offline. |
| `dashboard.html` | The dashboard template. Loads data at runtime via `fetch('jira_data.json')`. Edit this to change layout/logic. |
| `jira_data.json` | The work-item dataset consumed by the dashboard. |
| `generate_dashboard_data.py` | Converts the Excel source (`JIRA.xlsx`) into `jira_data.json`. |
| `create_standalone_dashboard.py` | Embeds `jira_data.json` into `dashboard.html` to produce `dashboard_standalone.html`. |
| `JIRA.xlsx` | Source spreadsheet of work items. |

## Data pipeline

```
JIRA.xlsx  --(generate_dashboard_data.py)-->  jira_data.json  --(create_standalone_dashboard.py)-->  dashboard_standalone.html
```

1. **Excel → JSON**: `generate_dashboard_data.py` reads `JIRA.xlsx`, drops the unused `Cost` column, formats dates, and writes `jira_data.json`.
2. **JSON → Standalone**: `create_standalone_dashboard.py` replaces the `fetch('jira_data.json')` loader in `dashboard.html` with the data embedded inline, producing `dashboard_standalone.html`.

### Regenerating the dashboard

Requires Python 3 with `pandas` and `openpyxl` (only for step 1).

```bash
# 1. (Only if the Excel source changed) rebuild the JSON dataset
python generate_dashboard_data.py

# 2. Rebuild the standalone dashboard
python create_standalone_dashboard.py
```

> Note: the standalone build runs the dashboard's initialization synchronously. Any top-level `const`/`let` used during startup must be declared at the top of the script block.

## Data model

Each record in `jira_data.json` is a work item with these fields:

| Field | Meaning |
| --- | --- |
| `Type` | `Story`, `Task`, or `Bug` |
| `Project`, `EPIC`, `Story`, `Task` | Hierarchy labels |
| `Sprint` | Sprint the item belongs to (e.g. `Sprint 3`) |
| `Quality` | **Number of open bugs** on the item (integer). Lower is better. |
| `Budget Planned`, `Budget Consumed`, `Budget Remaining` | Effort/budget figures. `Budget Remaining` is derived at load. |
| `Saving Planned`, `Saving Achived`, `Saving Pending` | Savings figures. `Saving Pending` is derived at load. |
| `Story Points` | Estimated effort in points |
| `Date` | Due date (`YYYY-MM-DD`) |
| `Priority` | `High`, `Medium`, `Low` |
| `Assignee` | Owner |
| `Status` | `To Do`, `In Progress`, `In Review`, `Done` |
| `Dependencies` | Optional dependency reference |

## Health model (E / T / Q)

Every task carries three independent RAG metrics; parents (Story, Epic, Project, Sprint) roll up the **worst** status of their children.

- **E — Efforts**: budget/effort consumed vs planned. Over budget → red; near the limit (≥90%) → amber; otherwise green.
- **T — Timeline**: schedule adherence. Done or comfortably ahead of the due date → green; not started or due within 7 days → amber; overdue and not done → red.
- **Q — Quality**: open bug count vs the allowance. Within allowance → green; over allowance → red.

### Bug allowance

Quality is a bug count with a configurable allowance defined once as `ALLOWED_BUGS` near the top of the script in `dashboard.html`:

```js
const ALLOWED_BUGS = 1; // allowed bugs per ticket
```

- A ticket **within** the allowance (≤ `ALLOWED_BUGS`) is green.
- A ticket **over** the allowance (> `ALLOWED_BUGS`) is red.

Change this single value to enforce a different limit; it flows through the task chips, rollups, filters, chart, and project health.

### Project-level status (Executive Summary & Project Health)

A project's overall health is the worst of its Efforts and Quality health:

- **Efforts** (efficiency = completion % − budget-spent %): red if `< -50`, amber if `< -25`, else green.
- **Quality** (average bugs per ticket vs allowance): red if `> 1.5 × ALLOWED_BUGS`, amber if `> ALLOWED_BUGS`, else green.

The Executive Summary tallies how many projects are On Track (green), At Risk (amber), and Critical (red).

## Using the dashboard

- **Filters**: narrow by Project, Status, Priority, Quality (bug count), or free-text search. Everything (metrics, health, tables, charts) updates together. **Reset Filters** clears them.
- **Drill-down**: the Detailed Task View and Sprint View start fully collapsed. Click any row to expand a level, or use **Expand All / Collapse All** (each section has its own button).
- **Color key**: green = on track, amber = at risk, red = critical. E/T/Q chips appear on every group and task row.

## Customizing

- **Layout, styling, and logic**: edit `dashboard.html`, then run `create_standalone_dashboard.py` to rebuild the standalone file.
- **Thresholds**: health cutoffs live inline in `updateExecSummary`, `calculateRAGStatus`, and the task-level `calc*` helpers; the bug allowance is the `ALLOWED_BUGS` constant.
- **Data**: update `JIRA.xlsx` and rerun both scripts, or edit `jira_data.json` directly and rerun `create_standalone_dashboard.py`.
