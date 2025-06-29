# 📊 Excel Automation Bot

Automatically clean and analyze Excel (.xlsx) files using Python and openpyxl.

## Features
- Deletes empty rows
- Calculates totals across numeric columns
- Highlights totals above a threshold
- Processes all files in `input_files/`

## How to Use
1. Place all `.xlsx` files in `input_files/`
2. Run:
```bash
pip install -r requirements.txt
python automate_excel.py
