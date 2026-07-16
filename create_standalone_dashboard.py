"""Generate dashboard_standalone.html from dashboard.html + jira_data.json.

The template (dashboard.html) loads data at runtime via fetch('jira_data.json').
This script embeds the JSON data directly into the HTML so the resulting
dashboard_standalone.html works by simply opening it in any browser, with no
web server or separate JSON file required.
"""
import json
import re

# Read the JSON data. Note: jira_data.json may contain NaN tokens which are not
# valid strict JSON, but Python's json module parses them by default.
with open('jira_data.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

# Read the HTML template
with open('dashboard.html', 'r', encoding='utf-8') as f:
    html_content = f.read()

# The fetch-based loader in the template (must match dashboard.html exactly).
loader_pattern = re.compile(
    r"        // Load data from jira_data\.json\s*\n"
    r"        fetch\('jira_data\.json'\).*?"
    r"\}\);",
    re.DOTALL,
)

# Embed the data directly. json.dumps emits NaN by default, which is valid
# JavaScript (the dashboard recomputes derived fields on load anyway).
embedded_code = (
    "        // Load data (embedded)\n"
    "        allData = " + json.dumps(data) + ";\n"
    "                initializeDashboard();"
)

new_html, count = loader_pattern.subn(embedded_code, html_content)

if count != 1:
    raise SystemExit(
        f"Expected exactly 1 fetch loader in dashboard.html, found {count}. "
        "Ensure dashboard.html contains the fetch('jira_data.json') loader."
    )

# Write the standalone HTML
with open('dashboard_standalone.html', 'w', encoding='utf-8') as f:
    f.write(new_html)

print(f"Successfully created dashboard_standalone.html with {len(data)} records embedded!")
print("You can now open it in any browser without needing the JSON file.")
