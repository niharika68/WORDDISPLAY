# Pharmacy Procurement Report Generator

A simple tool that automatically creates professional reports for pharmacy procurement data.

---

## What Does This Do?

This tool generates a complete report package with one click:

1. **Excel Spreadsheet** (`Enterprise_Report.xlsx`)
   - Two organized sheets: Summary and Orders
   - Professional formatting with green headers
   - Easy-to-read currency and number formatting
   - Filters enabled so you can sort and search

2. **Table Images** (`summary.png` and `orders.png`)
   - Screenshots of your data tables
   - Ready to paste into emails or presentations

3. **Word Document** (`Final_Report.docx`)
   - Professional report with title and timestamp
   - Contains both Summary and Orders tables
   - Ready to print or share

---

## Sample Data Included

The tool comes with realistic sample data for demonstration:

**Orders Data includes:**
- Hospital name
- Pharmacy supplier
- Drug name and pricing
- Number of units ordered
- Order date
- Whether it's been invoiced (Yes/No)
- Calculated order value

**Summary Data shows:**
- Monthly totals
- Total orders per month
- Total spending
- Savings compared to previous month

---

## How to Run

### First Time Setup

1. Make sure you have Python installed on your computer
2. Open Terminal (Mac) or Command Prompt (Windows)
3. Navigate to this folder
4. Run this command once to install required tools:
   ```
   poetry install
   ```

### Generate Reports

Run this command whenever you want to create new reports:

```
poetry run python report_generator.py
```

### Find Your Reports

After running, look in the `output` folder for:
- `Enterprise_Report.xlsx` - Open with Excel
- `Final_Report.docx` - Open with Word
- `summary.png` and `orders.png` - Image files

---

## Project Files Explained

| File | Purpose |
|------|---------|
| `data_source.py` | Creates the sample procurement data |
| `report_generator.py` | Builds the Excel, images, and Word files |
| `pyproject.toml` | Lists the required software packages |
| `output/` folder | Where your generated reports appear |

---

## Need Help?

**Reports not generating?**
- Make sure you ran `poetry install` first
- Check that you're in the correct folder

**Can't open the files?**
- Excel file needs Microsoft Excel or Google Sheets
- Word file needs Microsoft Word or Google Docs
- PNG files open in any image viewer

**Want different data?**
- Edit `data_source.py` to change the sample data
- A developer can help customize it for your real data

---

## Quick Reference

| Action | Command |
|--------|---------|
| Install (first time) | `poetry install` |
| Generate reports | `poetry run python report_generator.py` |
| Open output folder | Look in the `output` folder |

---

*Generated reports include sample data for demonstration purposes.*
