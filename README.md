# Day 28 — Excel Charts & Dashboards

**Roadmap Phase:** Phase 2 — Advanced Excel + Python  
**Date:** May 2026  
**Libraries:** `openpyxl` 

---

## What This Project Does

Reads three sheets of financial data from an input Excel file and generates a fully charted output workbook with five chart sheets and an executive dashboard — all programmatically via Python.

---

## Files

| File | Purpose |
|------|---------|
| `day28_practice_data.xlsx` | Input — 3 sheets: Monthly Data, Dept Expenses, Stock Performance |
| `day28_charts_dashboard.xlsx` | Output — Bar, Line, Pie, Area charts + Executive Dashboard |
| `day28_input_data.py` | Generates the practice input file |
| `day28_charts_dashboard.py` | Main script — reads input, writes all charts and dashboard |

---

## Charts Built

- **Bar Chart** — Grouped column chart: Monthly Revenue vs Expenses
- **Line Chart** — Multi-series: 4 stock prices over 12 weeks
- **Pie Chart** — Department expense breakdown with % labels
- **Area Chart** — Revenue vs Expenses as filled area trend
- **Dashboard Sheet** — KPI tiles (Total Revenue, Expenses, Net Profit, Avg Margin) + 2 embedded charts

---

## Key Concepts Covered

- `Reference()` — mapping worksheet cells to chart data
- `add_data()` + `set_categories()` — attaching data and x-axis labels
- `chart.type`, `chart.grouping`, `chart.style` — chart configuration
- `ws.add_chart(chart, anchor)` — placing charts at specific cell positions
- Dashboard layout: KPI tiles + multiple charts on one sheet

---

## How to Run

```bash
python day28_input_data.py       # Create practice data file
python day28_charts_dashboard.py # Generate all charts and dashboard
```

---

## Portfolio Connection

This project demonstrates **automated financial reporting** — a core skill for data analyst and finance automation roles. The dashboard pattern (KPI tiles + charts from live data) is directly reused in **Portfolio Project 1: Financial Dashboard** (Day 78–79).
