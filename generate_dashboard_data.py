import pandas as pd
import json

# Read the Excel file
df = pd.read_excel('JIRA.xlsx')

# Drop the Cost column (no longer used by the dashboard)
df = df.drop(columns=['Cost'], errors='ignore')

# Convert datetime to string for JSON serialization
df['Date'] = pd.to_datetime(df['Date']).dt.strftime('%Y-%m-%d')

# Convert to JSON
data = df.to_dict(orient='records')

# Write to JSON file
with open('jira_data.json', 'w') as f:
    json.dump(data, f, indent=2)

print(f"Successfully converted {len(data)} records to jira_data.json")
