"""
DAY 28 — Excel Charts & Dashboards with openpyxl
================================================
Topic  : BarChart, LineChart, PieChart, AreaChart, chart styling, dashboard layout
Input  : day28_practice_data.xlsx  (3 sheets of financial data)
Output : day28_charts_dashboard.xlsx
Rules  : NO pandas — pure openpyxl only
"""

# ─── IMPORTS ──────────────────────────────────────────────────────────────────
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter

# Chart classes — each handles a different chart type
from openpyxl.chart import BarChart, LineChart, PieChart, AreaChart, Reference, Series

# For styling the chart's visual series colors
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.label import DataLabelList

# For cell formatting on the dashboard sheet
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ─── CONCEPT 1: REFERENCE ─────────────────────────────────────────────────────
# Reference() tells openpyxl WHERE the chart data lives in the spreadsheet.
# It is the bridge between your cells and the chart object.
#
# Syntax:
#   Reference(worksheet, min_col, min_row, max_col, max_row)
#
# Example: Reference(ws, min_col=2, min_row=1, max_col=2, max_row=13)
#   → Points to column B, rows 1–13 on worksheet ws
# ─────────────────────────────────────────────────────────────────────────────

# ─── CONCEPT 2: SERIES ────────────────────────────────────────────────────────
# Series() represents ONE line / one set of bars in a chart.
# A chart can have MULTIPLE series (e.g., Revenue + Expenses side by side).
#
# Syntax:
#   Series(values_reference, title_from_data=True)
# ─────────────────────────────────────────────────────────────────────────────

# ─── CONCEPT 3: add_chart(chart, anchor) ──────────────────────────────────────
# anchor = cell address string like "A2" or "H5"
# This is the TOP-LEFT corner where the chart will be placed.
# ws.add_chart(chart, "H2")   → chart starts at cell H2
# ─────────────────────────────────────────────────────────────────────────────

# ─── LOAD INPUT FILE ──────────────────────────────────────────────────────────
wb_in  = load_workbook("day28_practice_data.xlsx")
ws_monthly = wb_in["Monthly Data"]      # revenue / expenses / profit
ws_dept    = wb_in["Dept Expenses"]     # department spending (pie chart)
ws_stock   = wb_in["Stock Performance"] # 4 stocks over 12 weeks (line chart)

# ─── OUTPUT WORKBOOK ──────────────────────────────────────────────────────────
wb_out = Workbook()

# ══════════════════════════════════════════════════════════════════════════════
# CHART 1 — GROUPED BAR CHART: Revenue vs Expenses (monthly)
# ══════════════════════════════════════════════════════════════════════════════

ws_bar = wb_out.active
ws_bar.title = "Bar Chart"

# Copy data from input so chart has a local data source
ws_bar.append(["Month", "Revenue", "Expenses", "Net Profit"])
for row in ws_monthly.iter_rows(min_row=2, max_row=13, values_only=True):
    month  = row[0]
    rev    = row[1]
    exp    = row[2]
    profit = row[1] - row[2]           # calculate profit since no pandas
    ws_bar.append([month, rev, exp, profit])

# ── Build the BarChart object ──────────────────────────────────────────────
bar_chart = BarChart()

# CONCEPT: chart.type
#   "col"  = vertical bars (columns) — most common for comparisons
#   "bar"  = horizontal bars
bar_chart.type = "col"

# CONCEPT: chart.grouping
#   "clustered"  = bars side by side (comparison)
#   "stacked"    = bars stacked on top of each other
#   "percentStacked" = stacked to 100%
bar_chart.grouping = "clustered"

bar_chart.title     = "Monthly Revenue vs Expenses (FY 2026)"
bar_chart.y_axis.title = "Amount (₹)"
bar_chart.x_axis.title = "Month"
bar_chart.style     = 10         # built-in Excel style 1–48
bar_chart.width     = 20         # width in cm
bar_chart.height    = 12         # height in cm

# CONCEPT: Reference for data (values)
#   min_col=2, max_col=3 → columns B and C (Revenue, Expenses)
#   min_row=1 → include row 1 as header/title
data_ref = Reference(ws_bar, min_col=2, max_col=3, min_row=1, max_row=13)

# CONCEPT: Reference for categories (x-axis labels)
cats_ref = Reference(ws_bar, min_col=1, min_row=2, max_row=13)

# CONCEPT: add_data — attaches the data reference to the chart
#   titles_from_data=True → first row of reference is used as series title
bar_chart.add_data(data_ref, titles_from_data=True)

# CONCEPT: set_categories — labels for the x-axis (months)
bar_chart.set_categories(cats_ref)

# Place chart on the sheet — anchor = top-left cell of chart
ws_bar.add_chart(bar_chart, "F2")

print("✓ Chart 1: Grouped Bar Chart created")


# ══════════════════════════════════════════════════════════════════════════════
# CHART 2 — LINE CHART: 4 Stock Prices over 12 Weeks
# ══════════════════════════════════════════════════════════════════════════════

ws_line = wb_out.create_sheet("Line Chart")

# Copy stock data
ws_line.append(["Week", "RELIANCE", "INFY", "TCS", "HDFC"])
for row in ws_stock.iter_rows(min_row=2, max_row=13, values_only=True):
    ws_line.append(list(row))

# ── Build LineChart ─────────────────────────────────────────────────────────
line_chart = LineChart()

# CONCEPT: chart.grouping for line chart
#   "standard"      = independent lines (default)
#   "stacked"       = stacked lines
#   "percentStacked"= stacked to 100%
line_chart.grouping = "standard"

# CONCEPT: smooth = True → curves instead of sharp corners
line_chart.smooth = True

line_chart.title        = "Stock Price Comparison – 12 Weeks"
line_chart.y_axis.title = "Price (₹)"
line_chart.x_axis.title = "Week"
line_chart.style        = 12
line_chart.width        = 22
line_chart.height       = 13

# Reference columns 2–5 (all 4 stocks), row 1 = header
stock_data_ref = Reference(ws_line, min_col=2, max_col=5, min_row=1, max_row=13)
stock_cats_ref = Reference(ws_line, min_col=1, min_row=2, max_row=13)

line_chart.add_data(stock_data_ref, titles_from_data=True)
line_chart.set_categories(stock_cats_ref)

ws_line.add_chart(line_chart, "G2")

print("✓ Chart 2: Multi-Series Line Chart created")


# ══════════════════════════════════════════════════════════════════════════════
# CHART 3 — PIE CHART: Department Expense Breakdown
# ══════════════════════════════════════════════════════════════════════════════

ws_pie = wb_out.create_sheet("Pie Chart")

ws_pie.append(["Department", "Actual Spend"])
for row in ws_dept.iter_rows(min_row=2, max_row=7, values_only=True):
    dept   = row[0]
    actual = row[2]
    ws_pie.append([dept, actual])

# ── Build PieChart ──────────────────────────────────────────────────────────
pie_chart = PieChart()

# CONCEPT: Pie chart only needs ONE data column (values) + ONE category column
pie_data_ref = Reference(ws_pie, min_col=2, min_row=1, max_row=7)  # includes header
pie_cats_ref = Reference(ws_pie, min_col=1, min_row=2, max_row=7)  # dept names

pie_chart.add_data(pie_data_ref, titles_from_data=True)
pie_chart.set_categories(pie_cats_ref)

# CONCEPT: dataLabels — show percentages on slices
pie_chart.dataLabels = DataLabelList()
pie_chart.dataLabels.showPercent  = True
pie_chart.dataLabels.showCatName  = False
pie_chart.dataLabels.showVal      = False

# CONCEPT: chart.title.tx for custom title styling
pie_chart.title  = "Department Expense Breakdown"
pie_chart.style  = 26         # darker style with distinct slice colors
pie_chart.width  = 18
pie_chart.height = 13

ws_pie.add_chart(pie_chart, "D2")

print("✓ Chart 3: Pie Chart created")


# ══════════════════════════════════════════════════════════════════════════════
# CHART 4 — AREA CHART: Revenue Trend (filled area under the line)
# ══════════════════════════════════════════════════════════════════════════════

ws_area = wb_out.create_sheet("Area Chart")

ws_area.append(["Month", "Revenue", "Expenses"])
for row in ws_monthly.iter_rows(min_row=2, max_row=13, values_only=True):
    ws_area.append([row[0], row[1], row[2]])

# ── Build AreaChart ─────────────────────────────────────────────────────────
area_chart = AreaChart()

# CONCEPT: AreaChart is a line chart with the area BELOW filled with color.
# Useful for showing cumulative/volume data over time.
# grouping="stacked" → areas stacked on top of each other (total view)
area_chart.grouping = "standard"

area_chart.title        = "Revenue vs Expenses – Area View"
area_chart.y_axis.title = "Amount (₹)"
area_chart.x_axis.title = "Month"
area_chart.style        = 14
area_chart.width        = 20
area_chart.height       = 12

area_data_ref = Reference(ws_area, min_col=2, max_col=3, min_row=1, max_row=13)
area_cats_ref = Reference(ws_area, min_col=1, min_row=2, max_row=13)

area_chart.add_data(area_data_ref, titles_from_data=True)
area_chart.set_categories(area_cats_ref)

ws_area.add_chart(area_chart, "F2")

print("✓ Chart 4: Area Chart created")


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 5 — EXECUTIVE DASHBOARD (KPI tiles + all charts embedded)
# ══════════════════════════════════════════════════════════════════════════════
# CONCEPT: A dashboard is a SINGLE sheet that combines:
#   • Summary KPI numbers (key metrics at a glance)
#   • Multiple charts positioned side by side
#   • No raw data tables — only insights

ws_dash = wb_out.create_sheet("Dashboard")

# ── Helper styles ────────────────────────────────────────────────────────────
def make_kpi_tile(ws, row, col, label, formula, bg_color):
    """Write a KPI label + value pair with colored background."""
    thin = Side(style="thin", color="FFFFFF")
    bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)

    label_cell = ws.cell(row=row,   column=col, value=label)
    value_cell = ws.cell(row=row+1, column=col, value=formula)

    label_cell.font      = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    label_cell.fill      = PatternFill("solid", start_color=bg_color)
    label_cell.alignment = Alignment(horizontal="center", vertical="center")
    label_cell.border    = bdr

    value_cell.font         = Font(bold=True, color="FFFFFF", name="Arial", size=13)
    value_cell.fill         = PatternFill("solid", start_color=bg_color)
    value_cell.alignment    = Alignment(horizontal="center", vertical="center")
    value_cell.number_format = '#,##0'
    value_cell.border       = bdr

# Dashboard Title
ws_dash.merge_cells("A1:P1")
title_cell = ws_dash["A1"]
title_cell.value     = "Financial Performance Dashboard — FY 2026"
title_cell.font      = Font(bold=True, color="FFFFFF", name="Arial", size=16)
title_cell.fill      = PatternFill("solid", start_color="1F4E79")
title_cell.alignment = Alignment(horizontal="center", vertical="center")
ws_dash.row_dimensions[1].height = 35

# ── Populate a local data table on dashboard (for KPI formulas) ──────────────
# Copy monthly data to columns R–V (hidden area, used by KPI formulas)
ws_dash.cell(row=1, column=18, value="Month")
ws_dash.cell(row=1, column=19, value="Revenue")
ws_dash.cell(row=1, column=20, value="Expenses")
ws_dash.cell(row=1, column=21, value="Profit")

monthly_rows = list(ws_monthly.iter_rows(min_row=2, max_row=13, values_only=True))
for i, row in enumerate(monthly_rows, 2):
    ws_dash.cell(row=i, column=18, value=row[0])
    ws_dash.cell(row=i, column=19, value=row[1])
    ws_dash.cell(row=i, column=20, value=row[2])
    ws_dash.cell(row=i, column=21, value=row[1] - row[2])

# ── KPI Tiles (row 3–4) ───────────────────────────────────────────────────────
make_kpi_tile(ws_dash, row=3, col=2,  label="Total Revenue (₹)",  formula="=SUM(S2:S13)", bg_color="2E75B6")
make_kpi_tile(ws_dash, row=3, col=5,  label="Total Expenses (₹)", formula="=SUM(T2:T13)", bg_color="C00000")
make_kpi_tile(ws_dash, row=3, col=8,  label="Net Profit (₹)",     formula="=SUM(U2:U13)", bg_color="375623")
make_kpi_tile(ws_dash, row=3, col=11, label="Avg Margin (%)",     formula="=AVERAGE(U2:U13)/AVERAGE(S2:S13)", bg_color="7030A0")

ws_dash.cell(row=4, column=11).number_format = "0.0%"

# Set column widths for KPI display
for col in range(1, 17):
    ws_dash.column_dimensions[get_column_letter(col)].width = 14
ws_dash.row_dimensions[3].height = 22
ws_dash.row_dimensions[4].height = 28

# ── Embed Bar Chart on Dashboard ─────────────────────────────────────────────
# CONCEPT: You can add a chart to ANY sheet — not just the one with the data.
# Create a fresh bar chart that references the dashboard's local data table.
dash_bar = BarChart()
dash_bar.type      = "col"
dash_bar.grouping  = "clustered"
dash_bar.title     = "Revenue vs Expenses"
dash_bar.style     = 10
dash_bar.width     = 20
dash_bar.height    = 12

# Reference dashboard's hidden data columns S and T (cols 19 and 20)
dash_bar_data = Reference(ws_dash, min_col=19, max_col=20, min_row=1, max_row=13)
dash_bar_cats = Reference(ws_dash, min_col=18, min_row=2, max_row=13)
dash_bar.add_data(dash_bar_data, titles_from_data=True)
dash_bar.set_categories(dash_bar_cats)
ws_dash.add_chart(dash_bar, "B6")

# ── Embed Area Chart on Dashboard ────────────────────────────────────────────
dash_area = AreaChart()
dash_area.grouping = "standard"
dash_area.title    = "Profit Trend"
dash_area.style    = 14
dash_area.width    = 20
dash_area.height   = 12

# Only profit column (col 21) for a clean trend view
dash_area_data = Reference(ws_dash, min_col=21, min_row=1, max_row=13)
dash_area_cats = Reference(ws_dash, min_col=18, min_row=2, max_row=13)
dash_area.add_data(dash_area_data, titles_from_data=True)
dash_area.set_categories(dash_area_cats)
ws_dash.add_chart(dash_area, "J6")

print("✓ Chart 5 + 6: Dashboard bar & area charts created")


# ══════════════════════════════════════════════════════════════════════════════
# FINALIZE — Remove default "Sheet" tab, save
# ══════════════════════════════════════════════════════════════════════════════

wb_out.save("day28_charts_dashboard.xlsx")
print("\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
print("OUTPUT SAVED: day28_charts_dashboard.xlsx")
print("Sheets:")
for s in wb_out.sheetnames:
    print(f"  • {s}")
print("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
