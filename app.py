"""
Flask UI for JIRA Dashboard Generator
--------------------------------------
1. Upload a JIRA Excel file
2. Convert it to jira_data.json  (logic from generate_dashboard_data.py)
3. Embed the JSON into dashboard_standalone.html (logic from create_standalone_dashboard.py)
4. Serve the finished dashboard for download / inline preview
"""

import json
import os
import re
import traceback
import logging

import pandas as pd
from flask import (
    Flask,
    Response,
    flash,
    redirect,
    render_template_string,
    request,
    send_file,
    url_for,
)
from werkzeug.utils import secure_filename

# Configure logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

try:
    from create_template import build_template
except ImportError as e:
    logger.error(f"Failed to import create_template: {e}")
    build_template = None

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
ALLOWED_EXTENSIONS = {"xlsx", "xls"}

app = Flask(__name__)
app.secret_key = "jira-dashboard-secret-key-change-me"
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

TEMPLATE_PATH = os.path.join(BASE_DIR, "JIRA_template.xlsx")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Pre-generate the template if it doesn't exist yet
if not os.path.exists(TEMPLATE_PATH):
    try:
        if build_template:
            build_template(TEMPLATE_PATH)
            logger.info(f"Template created at {TEMPLATE_PATH}")
        else:
            logger.warning("build_template function not available, template will be created on-demand")
    except Exception as e:
        logger.error(f"Failed to create template: {e}")
        pass  # non-fatal; download route will surface the error

# ---------------------------------------------------------------------------
# HTML template (inline – no separate templates folder needed)
# ---------------------------------------------------------------------------
PAGE_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>JIRA Dashboard Generator</title>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      background: #090c10;
      color: #c9d1d9;
      min-height: 100vh;
      display: grid;
      place-items: center;
      padding: 2rem 1rem;
    }

    /* ── Page shell ─────────────────────────────────────────────────── */
    .shell {
      width: 100%;
      max-width: 680px;
      display: flex;
      flex-direction: column;
      gap: 0;
    }

    /* ── Top header bar ─────────────────────────────────────────────── */
    .top-bar {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 1.1rem 1.75rem;
      background: #161b22;
      border: 1px solid #30363d;
      border-bottom: none;
      border-radius: 14px 14px 0 0;
    }
    .brand { display: flex; align-items: center; gap: .75rem; }
    .brand-icon {
      width: 38px; height: 38px; border-radius: 9px;
      background: linear-gradient(135deg, #1f6feb 0%, #388bfd 100%);
      display: flex; align-items: center; justify-content: center;
      box-shadow: 0 2px 8px rgba(31,111,235,.45);
      flex-shrink: 0;
    }
    .brand h1 { font-size: 1.15rem; font-weight: 700; color: #e6edf3; letter-spacing: -.01em; }
    .brand p  { font-size: .78rem; color: #6e7681; margin-top: 1px; }
    .badge {
      font-size: .7rem; font-weight: 600; padding: .25rem .6rem;
      background: #1c2333; border: 1px solid #30363d;
      border-radius: 20px; color: #58a6ff;
    }

    /* ── Step tracker ───────────────────────────────────────────────── */
    .stepper {
      display: flex;
      align-items: center;
      padding: 1.1rem 1.75rem;
      background: #0d1117;
      border-left: 1px solid #30363d;
      border-right: 1px solid #30363d;
      gap: 0;
    }
    .step-item {
      display: flex;
      align-items: center;
      flex: 1;
      gap: 0;
    }
    .step-item:last-child { flex: 0; }
    .step-content {
      display: flex;
      flex-direction: column;
      align-items: center;
      gap: .3rem;
      min-width: 64px;
    }
    .step-dot {
      width: 32px; height: 32px; border-radius: 50%;
      background: #161b22; border: 2px solid #30363d;
      display: flex; align-items: center; justify-content: center;
      font-size: .78rem; font-weight: 700; color: #484f58;
      transition: all .25s;
    }
    .step-item.active .step-dot {
      background: #1f6feb; border-color: #388bfd;
      color: #fff; box-shadow: 0 0 0 4px rgba(31,111,235,.2);
    }
    .step-item.done .step-dot {
      background: #238636; border-color: #2ea043; color: #fff;
    }
    .step-label {
      font-size: .7rem; font-weight: 600; color: #484f58;
      white-space: nowrap; letter-spacing: .02em; text-transform: uppercase;
    }
    .step-item.active .step-label { color: #58a6ff; }
    .step-item.done  .step-label  { color: #3fb950; }
    .step-connector {
      flex: 1; height: 2px;
      background: #21262d;
      margin: 0 .5rem;
      margin-bottom: 1.4rem;
      transition: background .25s;
    }
    .step-connector.done { background: #2ea043; }

    /* ── Main card body ─────────────────────────────────────────────── */
    .card-body {
      background: #161b22;
      border: 1px solid #30363d;
      border-top: none;
      border-radius: 0 0 14px 14px;
      padding: 2rem 1.75rem 1.75rem;
      box-shadow: 0 12px 40px rgba(0,0,0,.5);
    }

    /* ── Upload zone ─────────────────────────────────────────────────── */
    label.upload-area {
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      gap: .55rem;
      border: 2px dashed #30363d;
      border-radius: 10px;
      padding: 2.5rem 1.5rem;
      cursor: pointer;
      transition: border-color .2s, background .2s;
      margin-bottom: 1.25rem;
      background: #0d1117;
    }
    label.upload-area:hover,
    label.upload-area.drag-over { border-color: #388bfd; background: #0c1929; }
    .upload-icon {
      width: 52px; height: 52px; border-radius: 12px;
      background: #1c2333; border: 1px solid #30363d;
      display: flex; align-items: center; justify-content: center;
      font-size: 1.5rem;
    }
    .upload-title { font-size: .95rem; font-weight: 600; color: #e6edf3; }
    .upload-sub   { font-size: .8rem; color: #6e7681; }
    .file-chosen  { font-size: .85rem; color: #58a6ff; word-break: break-all; text-align: center; }
    #file-input   { display: none; }

    /* ── Buttons ─────────────────────────────────────────────────────── */
    .btn {
      display: flex; align-items: center; justify-content: center;
      gap: .5rem; width: 100%; padding: .75rem 1rem;
      border: none; border-radius: 8px;
      font-size: .95rem; font-weight: 600; cursor: pointer;
      transition: filter .15s, transform .1s;
      text-decoration: none;
    }
    .btn:active { transform: scale(.98); }
    .btn-primary { background: linear-gradient(135deg,#1f6feb,#388bfd); color: #fff; }
    .btn-primary:hover { filter: brightness(1.12); }
    .btn-primary:disabled { opacity: .45; cursor: not-allowed; filter: none; }
    .btn-success { background: linear-gradient(135deg,#238636,#2ea043); color: #fff; }
    .btn-success:hover { filter: brightness(1.12); }
    .btn-ghost {
      background: #21262d; color: #8b949e;
      border: 1px solid #30363d;
    }
    .btn-ghost:hover { background: #292e36; color: #c9d1d9; }
    .btn-outline {
      background: transparent; color: #58a6ff;
      border: 1px solid #388bfd;
    }
    .btn-outline:hover { background: rgba(56,139,253,.1); }
    .btn-row { display: flex; gap: .75rem; margin-top: .75rem; }
    .btn-row .btn { flex: 1; }

    /* ── Alerts ──────────────────────────────────────────────────────── */
    .alert {
      border-radius: 8px; padding: .85rem 1rem;
      margin-bottom: 1.25rem; font-size: .875rem;
      display: flex; align-items: flex-start; gap: .6rem;
    }
    .alert-danger  { background: #2d1517; border: 1px solid #6e2428; color: #ff7b72; }
    .alert-success { background: #0d2d1a; border: 1px solid #1a4731; color: #3fb950; }
    .alert-info    { background: #0c1929; border: 1px solid #1b3152; color: #79c0ff; }

    /* ── Stats row (step 2) ──────────────────────────────────────────── */
    .stat-card {
      background: #0d1117; border: 1px solid #30363d;
      border-radius: 10px; padding: 1.1rem 1.4rem;
      margin-bottom: 1.25rem;
      display: flex; align-items: center; gap: 1rem;
    }
    .stat-icon {
      font-size: 1.6rem; width: 48px; height: 48px;
      background: #161b22; border-radius: 10px;
      display: flex; align-items: center; justify-content: center;
      flex-shrink: 0; border: 1px solid #30363d;
    }
    .stat-label { font-size: .75rem; color: #6e7681; text-transform: uppercase;
                  letter-spacing: .06em; font-weight: 600; }
    .stat-value { font-size: 1.5rem; font-weight: 700; color: #e6edf3;
                  line-height: 1.1; margin-top: .1rem; }
    .stat-sub   { font-size: .78rem; color: #3fb950; margin-top: .15rem; }

    /* ── Result box (step 3) ─────────────────────────────────────────── */
    .result-box {
      background: #0d1117; border: 1px solid #1a4731;
      border-radius: 10px; padding: 1.25rem 1.4rem;
      margin-bottom: 1.25rem;
    }
    .result-box h3 { font-size: .95rem; font-weight: 700; color: #e6edf3; margin-bottom: .4rem; }
    .result-box p  { font-size: .83rem; color: #6e7681; margin-bottom: 1.1rem; line-height: 1.5; }

    /* ── Progress bar ────────────────────────────────────────────────── */
    .progress-wrap { display: none; margin-bottom: 1.1rem; }
    .progress-track {
      height: 5px; border-radius: 3px;
      background: #21262d; overflow: hidden;
    }
    .progress-fill {
      height: 100%; width: 0%;
      background: linear-gradient(90deg,#1f6feb,#58a6ff);
      border-radius: 3px; transition: width .45s ease;
    }
    .progress-label { font-size: .75rem; color: #6e7681; margin-top: .4rem; }

    /* ── Template hint ───────────────────────────────────────────────── */
    .template-hint {
      display: flex; align-items: center; justify-content: center;
      gap: .4rem; margin-top: 1rem;
      font-size: .8rem; color: #6e7681;
    }
    .template-hint a {
      color: #58a6ff; text-decoration: none; font-weight: 500;
      display: inline-flex; align-items: center; gap: .25rem;
    }
    .template-hint a:hover { text-decoration: underline; }

    /* ── Divider ─────────────────────────────────────────────────────── */
    .divider {
      border: none; border-top: 1px solid #21262d;
      margin: 1.25rem 0;
    }

    /* ── Footer ──────────────────────────────────────────────────────── */
    footer {
      text-align: center; margin-top: 1.25rem;
      font-size: .72rem; color: #30363d; letter-spacing: .03em;
    }
  </style>
</head>
<body>
<div class="shell">

  <!-- ── Top bar ── -->
  <div class="top-bar">
    <div class="brand">
      <div class="brand-icon">
        <svg width="22" height="22" viewBox="0 0 24 24" fill="none">
          <path d="M3 3h7v7H3zM14 3h7v7h-7zM14 14h7v7h-7zM3 14h7v7H3z"
                stroke="#fff" stroke-width="1.8" stroke-linejoin="round"/>
        </svg>
      </div>
      <div>
        <h1>JIRA Dashboard Generator</h1>
        <p>Upload Excel → Convert → Generate Standalone Report</p>
      </div>
    </div>
    <span class="badge">v1.0</span>
  </div>

  <!-- ── Step tracker ── -->
  <div class="stepper">
    <div class="step-item {{ 'done' if step > 1 else 'active' if step == 1 else '' }}">
      <div class="step-content">
        <div class="step-dot">{% if step > 1 %}✓{% else %}1{% endif %}</div>
        <div class="step-label">Upload</div>
      </div>
    </div>
    <div class="step-connector {{ 'done' if step > 1 else '' }}"></div>
    <div class="step-item {{ 'done' if step > 2 else 'active' if step == 2 else '' }}">
      <div class="step-content">
        <div class="step-dot">{% if step > 2 %}✓{% else %}2{% endif %}</div>
        <div class="step-label">Convert</div>
      </div>
    </div>
    <div class="step-connector {{ 'done' if step > 2 else '' }}"></div>
    <div class="step-item {{ 'done' if step > 3 else 'active' if step == 3 else '' }}">
      <div class="step-content">
        <div class="step-dot">{% if step > 3 %}✓{% else %}3{% endif %}</div>
        <div class="step-label">Generate</div>
      </div>
    </div>
  </div>

  <!-- ── Card body ── -->
  <div class="card-body">

    <!-- Alerts -->
    {% for category, message in messages %}
    <div class="alert alert-{{ category }}">
      <span>{% if category == 'danger' %}⚠{% elif category == 'success' %}✓{% else %}ℹ{% endif %}</span>
      <span>{{ message }}</span>
    </div>
    {% endfor %}

    {% if step == 1 %}
    <!-- STEP 1: Upload -->
    <form id="upload-form" method="post" action="{{ url_for('upload') }}" enctype="multipart/form-data">
      <label class="upload-area" for="file-input" id="drop-zone">
        <div class="upload-icon">�</div>
        <span class="upload-title" id="chosen-file">Drop your JIRA Excel file here</span>
        <span class="upload-sub">or click to browse &nbsp;·&nbsp; .xlsx / .xls</span>
      </label>
      <input id="file-input" type="file" name="excel_file" accept=".xlsx,.xls" required />

      <div class="progress-wrap" id="progress-wrap">
        <div class="progress-track"><div class="progress-fill" id="progress-fill"></div></div>
        <div class="progress-label" id="progress-label">Uploading…</div>
      </div>

      <button type="submit" class="btn btn-primary" id="submit-btn">
        <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor">
          <path d="M8 1a.75.75 0 0 1 .75.75v5.5h2.72a.25.25 0 0 1 .177.427l-3.396 3.396a.25.25 0 0 1-.353 0L4.503 7.677A.25.25 0 0 1 4.68 7.25H7.25V1.75A.75.75 0 0 1 8 1zM1.5 12a.75.75 0 0 1 .75-.75h11.5a.75.75 0 0 1 0 1.5H2.25A.75.75 0 0 1 1.5 12z"/>
        </svg>
        Upload &amp; Convert to JSON
      </button>
    </form>

    <div class="template-hint">
      <span>New to this?</span>
      <a href="{{ url_for('download_template') }}">
        <svg width="13" height="13" viewBox="0 0 16 16" fill="currentColor">
          <path d="M8 1a.75.75 0 0 1 .75.75v5.5h2.72a.25.25 0 0 1 .177.427l-3.396 3.396a.25.25 0 0 1-.353 0L4.503 7.677A.25.25 0 0 1 4.68 7.25H7.25V1.75A.75.75 0 0 1 8 1zM1.5 12a.75.75 0 0 1 .75-.75h11.5a.75.75 0 0 1 0 1.5H2.25A.75.75 0 0 1 1.5 12z"/>
        </svg>
        Download sample Excel template
      </a>
    </div>

    {% elif step == 2 %}
    <!-- STEP 2: Convert done, generate dashboard -->
    <div class="stat-card">
      <div class="stat-icon">🗂️</div>
      <div>
        <div class="stat-label">Records converted</div>
        <div class="stat-value">{{ record_count }}</div>
        <div class="stat-sub">✓ jira_data.json saved</div>
      </div>
    </div>

    <form method="post" action="{{ url_for('generate') }}">
      <button type="submit" class="btn btn-primary">
        <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor">
          <path d="M2 2.5A2.5 2.5 0 0 1 4.5 0h7A2.5 2.5 0 0 1 14 2.5v10.042a.75.75 0 0 1-1.225.575l-2.275-1.74-2.275 1.74a.75.75 0 0 1-.9 0L5.05 11.377l-2.275 1.74A.75.75 0 0 1 1.5 12.5V2.5zm2.5-1A1 1 0 0 0 3.5 2.5v8.653l1.525-1.166a.75.75 0 0 1 .9 0L8.2 11.727l2.275-1.74a.75.75 0 0 1 .9 0l1.525 1.166V2.5a1 1 0 0 0-1-1h-7z"/>
        </svg>
        Generate Standalone Dashboard
      </button>
    </form>

    <hr class="divider" />
    <form method="get" action="{{ url_for('index') }}">
      <button type="submit" class="btn btn-ghost">← Upload a different file</button>
    </form>

    {% elif step == 3 %}
    <!-- STEP 3: Done -->
    <div class="result-box">
      <h3>🎉 Dashboard ready</h3>
      <p>Your self-contained HTML report is ready. Download it and open in any browser — no server or separate JSON file needed.</p>
      <a href="{{ url_for('download') }}" class="btn btn-success" style="text-decoration:none">
        <svg width="15" height="15" viewBox="0 0 16 16" fill="currentColor">
          <path d="M8 1a.75.75 0 0 1 .75.75v5.5h2.72a.25.25 0 0 1 .177.427l-3.396 3.396a.25.25 0 0 1-.353 0L4.503 7.677A.25.25 0 0 1 4.68 7.25H7.25V1.75A.75.75 0 0 1 8 1zM1.5 12a.75.75 0 0 1 .75-.75h11.5a.75.75 0 0 1 0 1.5H2.25A.75.75 0 0 1 1.5 12z"/>
        </svg>
        Download dashboard_standalone.html
      </a>
      <div class="btn-row">
        <a href="{{ url_for('preview') }}" target="_blank" class="btn btn-outline" style="text-decoration:none">
          🔍 Preview in browser
        </a>
      </div>
    </div>

    <form method="get" action="{{ url_for('index') }}">
      <button type="submit" class="btn btn-ghost">← Start over</button>
    </form>
    {% endif %}

  </div><!-- /card-body -->
</div><!-- /shell -->

<footer>JIRA Management Report &nbsp;·&nbsp; Flask UI</footer>

<script>
  // File picker feedback
  const fileInput = document.getElementById('file-input');
  const fileLabel = document.getElementById('chosen-file');
  if (fileInput) {
    fileInput.addEventListener('change', () => {
      if (fileInput.files.length) {
        fileLabel.textContent = fileInput.files[0].name;
        fileLabel.style.color = '#58a6ff';
      } else {
        fileLabel.textContent = 'Drop your JIRA Excel file here';
        fileLabel.style.color = '';
      }
    });
  }

  // Drag-and-drop
  const dropZone = document.getElementById('drop-zone');
  if (dropZone) {
    ['dragenter', 'dragover'].forEach(evt =>
      dropZone.addEventListener(evt, e => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
      })
    );
    ['dragleave', 'drop'].forEach(evt =>
      dropZone.addEventListener(evt, e => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
      })
    );
    dropZone.addEventListener('drop', e => {
      const files = e.dataTransfer.files;
      if (files.length && fileInput) {
        try { fileInput.files = files; } catch(_) {}
        fileLabel.textContent = files[0].name;
        fileLabel.style.color = '#58a6ff';
      }
    });
  }

  // Upload progress animation
  const uploadForm = document.getElementById('upload-form');
  if (uploadForm) {
    uploadForm.addEventListener('submit', () => {
      const wrap = document.getElementById('progress-wrap');
      const fill = document.getElementById('progress-fill');
      const lbl  = document.getElementById('progress-label');
      const btn  = document.getElementById('submit-btn');
      if (wrap) wrap.style.display = 'block';
      if (btn)  btn.disabled = true;
      [
        [350,  25, 'Uploading file…'],
        [800,  55, 'Converting to JSON…'],
        [1300, 80, 'Almost done…'],
      ].forEach(([delay, pct, text]) => {
        setTimeout(() => {
          if (fill) fill.style.width = pct + '%';
          if (lbl)  lbl.textContent  = text;
        }, delay);
      });
    });
  }
</script>
</body>
</html>
"""

# ---------------------------------------------------------------------------
# Core processing helpers
# ---------------------------------------------------------------------------

def excel_to_json(excel_path: str, json_path: str) -> int:
    """Replicates generate_dashboard_data.py logic."""
    try:
        logger.debug(f"Reading Excel file: {excel_path}")
        df = pd.read_excel(excel_path)
        logger.debug(f"Excel file loaded with {len(df)} rows and columns: {list(df.columns)}")
        
        df = df.drop(columns=["Cost"], errors="ignore")
        
        # Handle date conversion more safely
        if "Date" in df.columns:
            df["Date"] = pd.to_datetime(df["Date"], errors='coerce').dt.strftime("%Y-%m-%d")
        
        data = df.to_dict(orient="records")
        logger.debug(f"Converting {len(data)} records to JSON")
        
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        
        logger.info(f"Successfully converted {len(data)} records to {json_path}")
        return len(data)
    except Exception as e:
        logger.error(f"Error in excel_to_json: {str(e)}")
        logger.error(f"Traceback: {traceback.format_exc()}")
        raise


def build_standalone(json_path: str, template_path: str, output_path: str) -> int:
    """Replicates create_standalone_dashboard.py logic."""
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    with open(template_path, "r", encoding="utf-8") as f:
        html_content = f.read()

    loader_pattern = re.compile(
        r"        // Load data from jira_data\.json\s*\n"
        r"        fetch\('jira_data\.json'\).*?"
        r"\}\);",
        re.DOTALL,
    )

    embedded_code = (
        "        // Load data (embedded)\n"
        "        allData = " + json.dumps(data) + ";\n"
        "                initializeDashboard();"
    )

    new_html, count = loader_pattern.subn(embedded_code, html_content)
    if count != 1:
        raise ValueError(
            f"Expected exactly 1 fetch loader in dashboard.html, found {count}."
        )

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(new_html)

    return len(data)


# ---------------------------------------------------------------------------
# Session-like state stored in app context (fine for single-user local tool)
# ---------------------------------------------------------------------------
_state: dict = {"step": 1, "record_count": 0, "messages": []}


def _render(step=None, messages=None, **kwargs):
    try:
        if step is not None:
            _state["step"] = step
        if messages is not None:
            _state["messages"] = messages
        
        context = {
            "step": _state.get("step", 1),
            "record_count": _state.get("record_count", 0),
            "messages": _state.get("messages", []),
            **kwargs
        }
        
        return render_template_string(PAGE_TEMPLATE, **context)
    except Exception as e:
        logger.error(f"Error in _render: {str(e)}")
        logger.error(f"Traceback: {traceback.format_exc()}")
        return f"Template rendering error: {str(e)}", 500


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/", methods=["GET"])
def index():
    try:
        _state["step"] = 1
        _state["messages"] = []
        return _render(step=1, messages=[])
    except Exception as e:
        logger.error(f"Error in index route: {str(e)}")
        return f"Application error: {str(e)}", 500


@app.route("/health")
def health_check():
    """Simple health check endpoint"""
    try:
        import pandas as pd
        import openpyxl
        
        health_info = {
            "status": "healthy",
            "upload_folder_exists": os.path.exists(UPLOAD_FOLDER),
            "template_exists": os.path.exists(TEMPLATE_PATH),
            "dashboard_template_exists": os.path.exists(os.path.join(BASE_DIR, "dashboard.html")),
            "base_dir": BASE_DIR,
            "pandas_version": pd.__version__,
            "openpyxl_available": True,
            "create_template_available": build_template is not None
        }
        
        # Test if we can create a simple DataFrame
        test_df = pd.DataFrame({"test": [1, 2, 3]})
        health_info["pandas_working"] = len(test_df) == 3
        
        return health_info
    except Exception as e:
        logger.error(f"Health check failed: {str(e)}")
        return {"status": "unhealthy", "error": str(e)}, 500


@app.route("/upload", methods=["POST"])
def upload():
    try:
        logger.debug("Upload route called")
        
        file = request.files.get("excel_file")
        if not file or file.filename == "":
            logger.warning("No file selected")
            return _render(step=1, messages=[("danger", "No file selected.")])

        logger.debug(f"File uploaded: {file.filename}")
        
        ext = file.filename.rsplit(".", 1)[-1].lower() if "." in file.filename else ""
        if ext not in ALLOWED_EXTENSIONS:
            logger.warning(f"Invalid file extension: {ext}")
            return _render(step=1, messages=[("danger", "Only .xlsx and .xls files are supported.")])

        filename = secure_filename(file.filename)
        excel_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        
        logger.debug(f"Saving file to: {excel_path}")
        file.save(excel_path)

        json_path = os.path.join(BASE_DIR, "jira_data.json")
        logger.debug(f"Converting to JSON: {json_path}")

        try:
            count = excel_to_json(excel_path, json_path)
            logger.info(f"Successfully processed {count} records")
        except Exception as e:
            err = traceback.format_exc()
            logger.error(f"Conversion failed: {err}")
            return _render(step=1, messages=[("danger", f"Conversion failed: {str(e)}")])

        _state["record_count"] = count
        return _render(step=2, messages=[])
        
    except Exception as e:
        logger.error(f"Unexpected error in upload route: {str(e)}")
        logger.error(f"Traceback: {traceback.format_exc()}")
        return _render(step=1, messages=[("danger", f"Upload failed: {str(e)}")])


@app.route("/generate", methods=["POST"])
def generate():
    json_path     = os.path.join(BASE_DIR, "jira_data.json")
    template_path = os.path.join(BASE_DIR, "dashboard.html")
    output_path   = os.path.join(BASE_DIR, "dashboard_standalone.html")

    if not os.path.exists(json_path):
        return _render(step=1, messages=[("danger", "jira_data.json not found. Please upload an Excel file first.")])
    if not os.path.exists(template_path):
        return _render(step=2, messages=[("danger", "dashboard.html template not found in project directory.")])

    try:
        build_standalone(json_path, template_path, output_path)
    except Exception:
        err = traceback.format_exc()
        return _render(step=2, messages=[("danger", f"Dashboard generation failed:\n{err}")])

    return _render(step=3, messages=[])


@app.route("/download")
def download():
    output_path = os.path.join(BASE_DIR, "dashboard_standalone.html")
    if not os.path.exists(output_path):
        return redirect(url_for("index"))
    return send_file(output_path, as_attachment=True, download_name="dashboard_standalone.html")


@app.route("/preview")
def preview():
    output_path = os.path.join(BASE_DIR, "dashboard_standalone.html")
    if not os.path.exists(output_path):
        return redirect(url_for("index"))
    with open(output_path, "r", encoding="utf-8") as f:
        content = f.read()
    return Response(content, mimetype="text/html")


@app.route("/template")
def download_template():
    try:
        if not os.path.exists(TEMPLATE_PATH):
            if build_template:
                try:
                    build_template(TEMPLATE_PATH)
                    logger.info(f"Template generated on-demand at {TEMPLATE_PATH}")
                except Exception as e:
                    logger.error(f"Failed to generate template: {e}")
                    err = traceback.format_exc()
                    return Response(f"Could not generate template: {str(e)}\n\nDetails:\n{err}", status=500, mimetype="text/plain")
            else:
                return Response("Template generation not available: create_template module not imported", status=500, mimetype="text/plain")
        
        return send_file(TEMPLATE_PATH, as_attachment=True, download_name="JIRA_template.xlsx")
    except Exception as e:
        logger.error(f"Error in template download: {str(e)}")
        return Response(f"Template download failed: {str(e)}", status=500, mimetype="text/plain")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

@app.errorhandler(500)
def internal_error(error):
    logger.error(f"Internal server error: {error}")
    return "Internal server error occurred. Please check the logs for details.", 500

@app.errorhandler(Exception)
def handle_exception(e):
    logger.error(f"Unhandled exception: {str(e)}")
    logger.error(f"Traceback: {traceback.format_exc()}")
    return f"An error occurred: {str(e)}", 500

if __name__ == "__main__":
    print("Starting JIRA Dashboard Generator at http://127.0.0.1:5000")
    app.run(debug=True, port=5000)
