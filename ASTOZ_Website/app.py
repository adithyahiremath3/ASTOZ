from flask import Flask, render_template
import pandas as pd
import os

app = Flask(__name__)

EXCEL_FILE = "ASTOZ.xlsx"  # Use raw string or forward slashes

# Ensure the file exists
if not os.path.exists(EXCEL_FILE):
    raise FileNotFoundError(f"File not found: {EXCEL_FILE}")

# Read sheet names safely
with pd.ExcelFile(EXCEL_FILE) as xls:
    sheet_names = xls.sheet_names

@app.route('/')
def index():
    return render_template('index.html', sheet_names=sheet_names)

@app.route('/table/<sheet_name>')
def show_table(sheet_name):
    if sheet_name not in sheet_names:
        return "Sheet not found", 404

    # Read Excel sheet
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, dtype=str)

    # Convert URLs to clickable links
    def make_clickable(val):
        if isinstance(val, str) and (val.startswith("http://") or val.startswith("https://")):
            return f'<a href="{val}" target="_blank">{val}</a>'
        return val

    df = df.applymap(make_clickable)  # ✅ Proper indentation

    # Data cleaning
    df = df.dropna(how='all')  # Remove empty rows
    df = df.dropna(axis=1, how='all')  # Remove empty columns
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)  # Trim spaces

    return render_template('table.html', sheet_name=sheet_name, tables=df.to_html(classes='data', index=False, escape=False))

if __name__ == '__main__':
    app.run(debug=True)
