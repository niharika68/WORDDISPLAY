# Pharmacy Procurement Report Generator

A simple tool that automatically creates a professional Word document report for pharmacy procurement data.

---

## What Does This Do?

This tool generates a **Word Document Report** (`Final_Report.docx`) with one click.

### The Final Report Includes:

📄 **Summary Table** - Monthly overview of orders, spending, and savings

📄 **Orders Table** - Detailed list of all procurement orders with NDC codes

📊 **Visualizations:**
- **Top 5 NDC by Spend** - Bar chart of highest-cost drug codes
- **Monthly Savings** - Bar chart showing savings trends (green = saved, red = increased)
- **Top 5 NDC Distribution** - Pie chart of spend concentration

### Supporting Files Also Generated:

| File | Purpose |
|------|---------|
| `Enterprise_Report.xlsx` | Interactive Excel for further analysis |
| `chart_top_ndc_spend.png` | Top 5 NDC bar chart |
| `chart_savings_by_month.png` | Monthly savings chart |
| `chart_top_ndc_pie.png` | NDC spend pie chart |

---

## Sample Data Included

The tool comes with realistic sample data for demonstration:

**Orders Data includes:**
- Hospital name
- Pharmacy supplier
- Drug name with NDC code
- Price and units ordered
- Order date
- Invoice status (Yes/No)
- Calculated order value

**Summary Data shows:**
- Monthly totals
- Total orders per month
- Total spending
- Savings percentage and dollar amount

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

### Find Your Report

After running, open the `output` folder and double-click:
- **`Final_Report.docx`** - Your complete Word document report

---

## Project Files Explained

| File | Purpose |
|------|---------|
| `data_source.py` | Creates the sample procurement data |
| `report_generator.py` | Builds the Word report and charts |
| `pyproject.toml` | Lists the required software packages |
| `output/` folder | Where your generated report appears |

---

## Libraries Used

| Library | Purpose |
|---------|---------|
| **pandas** | Data manipulation and analysis |
| **python-docx** | Creates Word documents (.docx) |
| **xlsxwriter** | Creates Excel files with formatting |
| **matplotlib** | Generates charts and visualizations |
| **openpyxl** | Reads/writes Excel files |

All libraries are automatically installed when you run `poetry install`.

---

## Need Help?

**Reports not generating?**
- Make sure you ran `poetry install` first
- Check that you're in the correct folder

**Can't open the Word file?**
- Needs Microsoft Word, Google Docs, or LibreOffice

**Want different data?**
- Edit `data_source.py` to change the sample data
- A developer can help customize it for your real data

---

## Quick Reference

| Action | Command |
|--------|---------|
| Install (first time) | `poetry install` |
| Generate report | `poetry run python report_generator.py` |
| Find your report | Open `output/Final_Report.docx` |

---

*Generated reports include sample data for demonstration purposes.*
