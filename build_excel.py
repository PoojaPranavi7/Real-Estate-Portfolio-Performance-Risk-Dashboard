from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


REPO_ROOT = Path(__file__).resolve().parent
DATA_DIR = REPO_ROOT / "data"
OUTPUT_PATH = REPO_ROOT / "excel" / "portfolio_dashboard.xlsx"


# Palette
DARK_NAVY = "1B2A4A"
HEADER_FONT_COLOR = "FFFFFF"
LIGHT_GRAY = "F5F5F5"
WHITE = "FFFFFF"
LIGHT_BLUE = "D9EAF7"
LIGHT_GREEN = "DDF3D8"
GREEN_FILL = "C6EFCE"
YELLOW_FILL = "FFEB9C"
RED_FILL = "FFC7CE"
ORANGE_FILL = "F4B942"
DARK_GREEN_FILL = "2E7D32"
MID_GREEN_FILL = "A9D18E"
GRAY_FILL = "D9D9D9"
DARK_ORANGE = "C55A11"
RED_TAB = "C00000"
DARK_RED_TEXT = "9C0006"

THIN_BORDER = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)

CARD_BORDER = Border(
    left=Side(style="medium", color="1B2A4A"),
    right=Side(style="medium", color="1B2A4A"),
    top=Side(style="medium", color="1B2A4A"),
    bottom=Side(style="medium", color="1B2A4A"),
)


def normalize_pct_value(value):
    if pd.isna(value):
        return value
    return value / 100 if value > 1.5 else value


def normalize_pct_series(series: pd.Series) -> pd.Series:
    return series.apply(normalize_pct_value)


def auto_fit_columns(ws, min_width=12, max_width=35):
    for col_idx, col_cells in enumerate(ws.columns, start=1):
        col_letter = get_column_letter(col_idx)
        max_len = 0
        for cell in col_cells:
            if type(cell).__name__ == "MergedCell":
                continue
            value = cell.value
            if value is None:
                continue
            if isinstance(value, pd.Timestamp):
                text = value.strftime("%b-%y")
            else:
                text = str(value)
            if len(text) > max_len:
                max_len = len(text)
        ws.column_dimensions[col_letter].width = max(min_width, min(max_width, max_len + 2))


def style_table_header(ws, header_row=1):
    header_fill = PatternFill(fill_type="solid", fgColor=DARK_NAVY)
    header_font = Font(name="Calibri", size=11, bold=True, color=HEADER_FONT_COLOR)
    for cell in ws[header_row]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER


def style_data_cells(ws, start_row, end_row, end_col):
    data_font = Font(name="Calibri", size=10)
    for row in ws.iter_rows(min_row=start_row, max_row=end_row, min_col=1, max_col=end_col):
        for cell in row:
            cell.font = data_font
            if cell.alignment.horizontal is None:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(vertical="center")
            cell.border = THIN_BORDER


def write_dataframe(ws, df: pd.DataFrame, start_row: int = 1, start_col: int = 1):
    for col_idx, col_name in enumerate(df.columns, start=start_col):
        ws.cell(row=start_row, column=col_idx, value=col_name)
    for row_idx, row in enumerate(df.itertuples(index=False), start=start_row + 1):
        for col_idx, value in enumerate(row, start=start_col):
            ws.cell(row=row_idx, column=col_idx, value=value)


def apply_border_range(ws, min_row, max_row, min_col, max_col, border):
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell.border = border


def style_custom_header(ws, row_idx, start_col, end_col):
    fill = PatternFill(fill_type="solid", fgColor=DARK_NAVY)
    font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row=row_idx, column=col)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER


def build_cover_tab(wb: Workbook):
    ws = wb.create_sheet("Cover")
    ws.sheet_properties.tabColor = DARK_NAVY

    # Fill canvas
    bg_fill = PatternFill(fill_type="solid", fgColor=DARK_NAVY)
    for r in range(1, 35):
        for c in range(1, 11):
            ws.cell(row=r, column=c).fill = bg_fill

    ws.merge_cells("B4:I6")
    ws["B4"] = "Real Estate Portfolio Performance & Risk Dashboard"
    ws["B4"].font = Font(name="Calibri", size=20, bold=True, color="FFFFFF")
    ws["B4"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("B8:I8")
    ws["B8"] = "Crystal View Capital / Osprey Management"
    ws["B8"].font = Font(name="Calibri", size=14, bold=False, color="D9D9D9")
    ws["B8"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("B11:I11")
    ws["B11"] = "Portfolio Period: January – December 2024"
    ws["B11"].font = Font(name="Calibri", size=12, color="FFFFFF")
    ws["B11"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("B13:I13")
    ws["B13"] = "12 Properties | 2 Asset Types | NV · AZ · NM · UT"
    ws["B13"].font = Font(name="Calibri", size=12, italic=True, color="B8C2D6")
    ws["B13"].alignment = Alignment(horizontal="center", vertical="center")

    for col in ("A", "J"):
        ws.column_dimensions[col].width = 4
    for col in ("B", "C", "D", "E", "F", "G", "H", "I"):
        ws.column_dimensions[col].width = 18
    ws.row_dimensions[4].height = 34
    ws.row_dimensions[5].height = 34
    ws.row_dimensions[6].height = 34


def build_properties_tab(wb: Workbook, properties_df: pd.DataFrame):
    ws = wb.create_sheet("Properties")
    ws.sheet_properties.tabColor = DARK_NAVY

    df = properties_df.rename(
        columns={
            "property_id": "Property ID",
            "property_name": "Property Name",
            "asset_type": "Asset Type",
            "city": "City",
            "state": "State",
            "total_units": "Total Units",
            "year_acquired": "Year Acquired",
            "purchase_price": "Purchase Price",
        }
    )

    write_dataframe(ws, df, start_row=1)
    style_table_header(ws)
    style_data_cells(ws, 2, ws.max_row, ws.max_column)
    ws.freeze_panes = "A2"

    for r in range(2, ws.max_row + 1):
        row_fill = PatternFill(fill_type="solid", fgColor=WHITE if r % 2 == 0 else LIGHT_GRAY)
        for c in range(1, ws.max_column + 1):
            ws.cell(row=r, column=c).fill = row_fill

        asset_cell = ws.cell(row=r, column=3)
        if asset_cell.value == "MHC":
            asset_cell.fill = PatternFill(fill_type="solid", fgColor=LIGHT_BLUE)
        elif asset_cell.value == "Self-Storage":
            asset_cell.fill = PatternFill(fill_type="solid", fgColor=LIGHT_GREEN)

        ws.cell(row=r, column=8).number_format = "$#,##0.00"

    auto_fit_columns(ws)


def build_occupancy_tab(wb: Workbook, properties_df: pd.DataFrame, occupancy_df: pd.DataFrame):
    ws = wb.create_sheet("Occupancy")
    ws.sheet_properties.tabColor = "4F81BD"

    occ = occupancy_df.copy()
    occ["month"] = pd.to_datetime(occ["month"])
    occ["occupancy_rate"] = normalize_pct_series(occ["occupancy_rate"])

    occ = occ.merge(
        properties_df[["property_id", "property_name", "asset_type"]],
        on="property_id",
        how="left",
    )

    def occupancy_flag(rate):
        if rate >= 0.93:
            return "On Target"
        if rate >= 0.88:
            return "Watch"
        return "Below Threshold"

    occ["Flag"] = occ["occupancy_rate"].apply(occupancy_flag)
    occ.sort_values(["month", "occupancy_rate"], ascending=[True, True], inplace=True)

    out = occ[
        [
            "property_id",
            "property_name",
            "asset_type",
            "month",
            "occupied_units",
            "total_units",
            "occupancy_rate",
            "Flag",
        ]
    ].rename(
        columns={
            "property_id": "Property ID",
            "property_name": "Property Name",
            "asset_type": "Asset Type",
            "month": "Month",
            "occupied_units": "Occupied Units",
            "total_units": "Total Units",
            "occupancy_rate": "Occupancy Rate",
        }
    )

    write_dataframe(ws, out, start_row=1)
    style_table_header(ws)
    style_data_cells(ws, 2, ws.max_row, ws.max_column)
    ws.freeze_panes = "A2"

    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=4).number_format = "mmm-yy"
        ws.cell(row=r, column=7).number_format = "0.0%"

        flag_cell = ws.cell(row=r, column=8)
        if flag_cell.value == "On Target":
            flag_cell.fill = PatternFill(fill_type="solid", fgColor=GREEN_FILL)
        elif flag_cell.value == "Watch":
            flag_cell.fill = PatternFill(fill_type="solid", fgColor=YELLOW_FILL)
        else:
            flag_cell.fill = PatternFill(fill_type="solid", fgColor=RED_FILL)

    auto_fit_columns(ws)


def build_delinquency_tab(wb: Workbook, properties_df: pd.DataFrame, delinquency_df: pd.DataFrame):
    ws = wb.create_sheet("Delinquency")
    ws.sheet_properties.tabColor = "ED7D31"

    dlq = delinquency_df.copy()
    dlq["month"] = pd.to_datetime(dlq["month"])
    dlq["delinquency_rate"] = normalize_pct_series(dlq["delinquency_rate"])

    dlq = dlq.merge(
        properties_df[["property_id", "property_name", "asset_type"]],
        on="property_id",
        how="left",
    )

    def risk_tier(rate):
        if rate < 0.05:
            return "Acceptable"
        if rate < 0.08:
            return "Monitor"
        if rate < 0.12:
            return "At Risk"
        return "Critical"

    dlq["Risk Tier"] = dlq["delinquency_rate"].apply(risk_tier)

    out = dlq[
        [
            "property_id",
            "property_name",
            "asset_type",
            "month",
            "rent_billed",
            "rent_collected",
            "past_due_amount",
            "delinquency_rate",
            "Risk Tier",
        ]
    ].rename(
        columns={
            "property_id": "Property ID",
            "property_name": "Property Name",
            "asset_type": "Asset Type",
            "month": "Month",
            "rent_billed": "Rent Billed",
            "rent_collected": "Rent Collected",
            "past_due_amount": "Past Due Amount",
            "delinquency_rate": "Delinquency Rate",
        }
    )

    write_dataframe(ws, out, start_row=1)
    style_table_header(ws)
    style_data_cells(ws, 2, ws.max_row, ws.max_column)
    ws.freeze_panes = "A2"

    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=4).number_format = "mmm-yy"
        for c in (5, 6, 7):
            ws.cell(row=r, column=c).number_format = "$#,##0.00"
        ws.cell(row=r, column=8).number_format = "0.0%"

        tier_cell = ws.cell(row=r, column=9)
        if tier_cell.value == "Acceptable":
            tier_cell.fill = PatternFill(fill_type="solid", fgColor=GREEN_FILL)
        elif tier_cell.value == "Monitor":
            tier_cell.fill = PatternFill(fill_type="solid", fgColor=YELLOW_FILL)
        elif tier_cell.value == "At Risk":
            tier_cell.fill = PatternFill(fill_type="solid", fgColor=ORANGE_FILL)
        else:
            tier_cell.fill = PatternFill(fill_type="solid", fgColor=RED_FILL)

    auto_fit_columns(ws)


def build_financials_tab(wb: Workbook, properties_df: pd.DataFrame, financials_df: pd.DataFrame):
    ws = wb.create_sheet("Financials")
    ws.sheet_properties.tabColor = "70AD47"

    fin = financials_df.copy()
    fin["month"] = pd.to_datetime(fin["month"])
    fin["noi_margin"] = normalize_pct_series(fin["noi_margin"])
    fin["expense_ratio"] = normalize_pct_series(fin["expense_ratio"])

    fin = fin.merge(
        properties_df[["property_id", "property_name", "asset_type"]],
        on="property_id",
        how="left",
    )

    out = fin[
        [
            "property_id",
            "property_name",
            "asset_type",
            "month",
            "total_revenue",
            "operating_expenses",
            "noi",
            "noi_margin",
            "expense_ratio",
        ]
    ].rename(
        columns={
            "property_id": "Property ID",
            "property_name": "Property Name",
            "asset_type": "Asset Type",
            "month": "Month",
            "total_revenue": "Total Revenue",
            "operating_expenses": "Operating Expenses",
            "noi": "NOI",
            "noi_margin": "NOI Margin",
            "expense_ratio": "Expense Ratio",
        }
    )

    write_dataframe(ws, out, start_row=1)
    style_table_header(ws)
    style_data_cells(ws, 2, ws.max_row, ws.max_column)
    ws.freeze_panes = "A2"

    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=4).number_format = "mmm-yy"
        for c in (5, 6, 7):
            ws.cell(row=r, column=c).number_format = "$#,##0.00"
        ws.cell(row=r, column=8).number_format = "0.0%"
        ws.cell(row=r, column=9).number_format = "0.0%"

        margin = ws.cell(row=r, column=8).value
        margin_cell = ws.cell(row=r, column=8)
        if margin >= 0.45:
            margin_cell.fill = PatternFill(fill_type="solid", fgColor=GREEN_FILL)
        elif margin >= 0.35:
            margin_cell.fill = PatternFill(fill_type="solid", fgColor=YELLOW_FILL)
        else:
            margin_cell.fill = PatternFill(fill_type="solid", fgColor=RED_FILL)

    auto_fit_columns(ws)


def build_market_comps_tab(wb: Workbook, properties_df: pd.DataFrame, market_df: pd.DataFrame):
    ws = wb.create_sheet("Market Comps")
    ws.sheet_properties.tabColor = "8064A2"

    mc = market_df.copy()
    mc["rent_upside_pct"] = normalize_pct_series(mc["rent_upside_pct"])
    mc["competitor_occupancy"] = normalize_pct_series(mc["competitor_occupancy"])
    mc["market_cap_rate"] = normalize_pct_series(mc["market_cap_rate"])

    mc = mc.merge(
        properties_df[["property_id", "property_name", "asset_type", "city", "state"]],
        on="property_id",
        how="left",
    )

    out = mc[
        [
            "property_name",
            "asset_type",
            "city",
            "state",
            "market_avg_rent",
            "property_avg_rent",
            "rent_upside_pct",
            "competitor_occupancy",
            "market_cap_rate",
            "opportunity_score",
        ]
    ].rename(
        columns={
            "property_name": "Property Name",
            "asset_type": "Asset Type",
            "city": "City",
            "state": "State",
            "market_avg_rent": "Market Avg Rent",
            "property_avg_rent": "Property Avg Rent",
            "rent_upside_pct": "Rent Upside %",
            "competitor_occupancy": "Competitor Occupancy",
            "market_cap_rate": "Market Cap Rate",
            "opportunity_score": "Opportunity Score",
        }
    )

    write_dataframe(ws, out, start_row=1)
    style_table_header(ws)
    style_data_cells(ws, 2, ws.max_row, ws.max_column)
    ws.freeze_panes = "A2"

    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=5).number_format = "$#,##0.00"
        ws.cell(row=r, column=6).number_format = "$#,##0.00"
        ws.cell(row=r, column=7).number_format = "0.0%"
        ws.cell(row=r, column=8).number_format = "0.0%"
        ws.cell(row=r, column=9).number_format = "0.0%"

        score_cell = ws.cell(row=r, column=10)
        score = score_cell.value
        if score >= 70:
            score_cell.fill = PatternFill(fill_type="solid", fgColor=DARK_GREEN_FILL)
            score_cell.font = Font(name="Calibri", size=10, color="FFFFFF", bold=True)
        elif score >= 50:
            score_cell.fill = PatternFill(fill_type="solid", fgColor=MID_GREEN_FILL)
        elif score >= 30:
            score_cell.fill = PatternFill(fill_type="solid", fgColor=YELLOW_FILL)
        else:
            score_cell.fill = PatternFill(fill_type="solid", fgColor=GRAY_FILL)

    auto_fit_columns(ws)


def build_capex_tab(wb: Workbook, properties_df: pd.DataFrame, capex_df: pd.DataFrame):
    ws = wb.create_sheet("CapEx Projects")
    ws.sheet_properties.tabColor = DARK_ORANGE

    capex = capex_df.copy()
    capex["start_date"] = pd.to_datetime(capex["start_date"])
    capex["end_date"] = pd.to_datetime(capex["end_date"])

    capex = capex.merge(
        properties_df[["property_id", "property_name", "asset_type"]],
        on="property_id",
        how="left",
    )

    # KPI summary cards
    total_budgeted = float(capex["budgeted_cost"].sum())
    total_actual = float(capex["actual_cost"].sum())
    total_variance = float(capex["variance"].sum())
    projects_over_budget = int((capex["variance"] > 0).sum())

    kpi_defs = [
        ("Total Budgeted", total_budgeted, "$#,##0.00"),
        ("Total Actual", total_actual, "$#,##0.00"),
        ("Total Variance", total_variance, "$#,##0.00"),
        ("Projects Over Budget", projects_over_budget, "0"),
    ]

    starts = [1, 4, 7, 10]
    for i, (label, value, num_fmt) in enumerate(kpi_defs):
        start_col = starts[i]
        end_col = start_col + 2
        ws.merge_cells(start_row=1, start_column=start_col, end_row=2, end_column=end_col)
        ws.merge_cells(start_row=3, start_column=start_col, end_row=5, end_column=end_col)

        label_cell = ws.cell(row=1, column=start_col, value=label)
        value_cell = ws.cell(row=3, column=start_col, value=value)

        for r in range(1, 6):
            for c in range(start_col, end_col + 1):
                cell = ws.cell(row=r, column=c)
                cell.fill = PatternFill(fill_type="solid", fgColor=DARK_NAVY)
                cell.border = CARD_BORDER

        label_cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        label_cell.alignment = Alignment(horizontal="center", vertical="center")

        value_color = "FFFFFF"
        if label == "Total Variance":
            value_color = "FF4D4D" if total_variance > 0 else "92D050"
        value_cell.font = Font(name="Calibri", size=16, bold=True, color=value_color)
        value_cell.alignment = Alignment(horizontal="center", vertical="center")
        value_cell.number_format = num_fmt

    detail = capex[
        [
            "property_name",
            "asset_type",
            "project_name",
            "project_type",
            "start_date",
            "end_date",
            "budgeted_cost",
            "actual_cost",
            "variance",
            "status",
        ]
    ].rename(
        columns={
            "property_name": "Property Name",
            "asset_type": "Asset Type",
            "project_name": "Project Name",
            "project_type": "Project Type",
            "start_date": "Start Date",
            "end_date": "End Date",
            "budgeted_cost": "Budgeted Cost",
            "actual_cost": "Actual Cost",
            "variance": "Variance",
            "status": "Status",
        }
    )

    data_start_row = 8
    write_dataframe(ws, detail, start_row=data_start_row)
    style_custom_header(ws, data_start_row, 1, detail.shape[1])
    style_data_cells(ws, data_start_row + 1, ws.max_row, ws.max_column)

    for r in range(data_start_row + 1, ws.max_row + 1):
        ws.cell(row=r, column=5).number_format = "mmm-yy"
        ws.cell(row=r, column=6).number_format = "mmm-yy"
        ws.cell(row=r, column=7).number_format = "$#,##0.00"
        ws.cell(row=r, column=8).number_format = "$#,##0.00"
        variance_cell = ws.cell(row=r, column=9)
        variance_cell.number_format = "$#,##0.00"
        variance_cell.font = Font(
            name="Calibri",
            size=10,
            color="C00000" if (variance_cell.value or 0) > 0 else "008000",
            bold=False,
        )

    ws.freeze_panes = "A9"
    auto_fit_columns(ws)


def build_dashboard_tab(
    wb: Workbook,
    properties_df: pd.DataFrame,
    occupancy_df: pd.DataFrame,
    delinquency_df: pd.DataFrame,
    financials_df: pd.DataFrame,
):
    ws = wb.create_sheet("Dashboard")
    ws.sheet_properties.tabColor = RED_TAB

    occ = occupancy_df.copy()
    dlq = delinquency_df.copy()
    fin = financials_df.copy()

    occ["month"] = pd.to_datetime(occ["month"])
    dlq["month"] = pd.to_datetime(dlq["month"])
    fin["month"] = pd.to_datetime(fin["month"])
    occ["occupancy_rate"] = normalize_pct_series(occ["occupancy_rate"])
    dlq["delinquency_rate"] = normalize_pct_series(dlq["delinquency_rate"])
    fin["noi_margin"] = normalize_pct_series(fin["noi_margin"])
    fin["expense_ratio"] = normalize_pct_series(fin["expense_ratio"])

    latest_month = occ["month"].max()
    occ_latest = occ[occ["month"] == latest_month].copy()
    dlq_latest = dlq[dlq["month"] == latest_month].copy()
    fin_latest = fin[fin["month"] == latest_month].copy()

    occ_latest = occ_latest.merge(
        properties_df[["property_id", "property_name", "asset_type"]],
        on="property_id",
        how="left",
    )
    dlq_latest = dlq_latest.merge(
        properties_df[["property_id", "property_name", "asset_type"]],
        on="property_id",
        how="left",
    )
    fin_latest = fin_latest.merge(
        properties_df[["property_id", "property_name", "asset_type"]],
        on="property_id",
        how="left",
    )

    ws.merge_cells("A1:L1")
    ws["A1"] = "Executive Portfolio Dashboard"
    ws["A1"].font = Font(name="Calibri", size=18, bold=True, color=DARK_NAVY)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells("A2:L2")
    ws["A2"] = "Most Recent Month Performance Summary"
    ws["A2"].font = Font(name="Calibri", size=11, color="666666")
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center")

    # Section A - KPI cards
    total_properties = int(len(properties_df))
    total_units = int(properties_df["total_units"].sum())
    portfolio_avg_occ = occ_latest["occupied_units"].sum() / occ_latest["total_units"].sum()
    total_monthly_noi = float(fin_latest["noi"].sum())
    avg_delinquency = dlq_latest["past_due_amount"].sum() / dlq_latest["rent_billed"].sum()
    total_aum = float(properties_df["purchase_price"].sum())

    kpis = [
        ("Total Properties", total_properties, "0"),
        ("Total Units", total_units, "#,##0"),
        ("Portfolio Avg Occupancy", portfolio_avg_occ, "0.0%"),
        ("Total Monthly NOI", total_monthly_noi, "$#,##0.00"),
        ("Avg Delinquency Rate", avg_delinquency, "0.0%"),
        ("Total AUM", total_aum, "$#,##0.00"),
    ]

    card_spans = [(1, 2), (3, 4), (5, 6), (7, 8), (9, 10), (11, 12)]
    for (label, value, number_format), (start_col, end_col) in zip(kpis, card_spans):
        ws.merge_cells(start_row=3, start_column=start_col, end_row=4, end_column=end_col)
        ws.merge_cells(start_row=5, start_column=start_col, end_row=8, end_column=end_col)

        label_cell = ws.cell(row=3, column=start_col, value=label)
        value_cell = ws.cell(row=5, column=start_col, value=value)

        for r in range(3, 9):
            for c in range(start_col, end_col + 1):
                cell = ws.cell(row=r, column=c)
                cell.fill = PatternFill(fill_type="solid", fgColor=DARK_NAVY)
                cell.border = CARD_BORDER

        label_cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        label_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        value_cell.font = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
        value_cell.alignment = Alignment(horizontal="center", vertical="center")
        value_cell.number_format = number_format

    # Section B - Asset type comparison
    ws.merge_cells("A10:C10")
    ws["A10"] = "Asset Type Comparison"
    ws["A10"].font = Font(name="Calibri", size=12, bold=True, color=DARK_NAVY)
    ws["A10"].alignment = Alignment(horizontal="left", vertical="center")

    metrics_header_row = 11
    comparison_rows = [
        "Avg Occupancy",
        "Avg Delinquency",
        "Total NOI",
        "Avg NOI Margin",
        "Avg Expense Ratio",
    ]

    mhc_occ = occ_latest[occ_latest["asset_type"] == "MHC"]
    stg_occ = occ_latest[occ_latest["asset_type"] == "Self-Storage"]
    mhc_dlq = dlq_latest[dlq_latest["asset_type"] == "MHC"]
    stg_dlq = dlq_latest[dlq_latest["asset_type"] == "Self-Storage"]
    mhc_fin = fin_latest[fin_latest["asset_type"] == "MHC"]
    stg_fin = fin_latest[fin_latest["asset_type"] == "Self-Storage"]

    comparison_values = [
        [
            mhc_occ["occupied_units"].sum() / mhc_occ["total_units"].sum(),
            stg_occ["occupied_units"].sum() / stg_occ["total_units"].sum(),
        ],
        [
            mhc_dlq["past_due_amount"].sum() / mhc_dlq["rent_billed"].sum(),
            stg_dlq["past_due_amount"].sum() / stg_dlq["rent_billed"].sum(),
        ],
        [
            mhc_fin["noi"].sum(),
            stg_fin["noi"].sum(),
        ],
        [
            mhc_fin["noi_margin"].mean(),
            stg_fin["noi_margin"].mean(),
        ],
        [
            mhc_fin["expense_ratio"].mean(),
            stg_fin["expense_ratio"].mean(),
        ],
    ]

    ws.cell(row=metrics_header_row, column=1, value="Metric")
    ws.cell(row=metrics_header_row, column=2, value="MHC")
    ws.cell(row=metrics_header_row, column=3, value="Self-Storage")
    style_custom_header(ws, metrics_header_row, 1, 3)

    for idx, metric in enumerate(comparison_rows, start=1):
        row = metrics_header_row + idx
        ws.cell(row=row, column=1, value=metric)
        ws.cell(row=row, column=2, value=comparison_values[idx - 1][0])
        ws.cell(row=row, column=3, value=comparison_values[idx - 1][1])

    style_data_cells(ws, metrics_header_row + 1, metrics_header_row + len(comparison_rows), 3)
    for r in range(metrics_header_row + 1, metrics_header_row + len(comparison_rows) + 1):
        metric_name = ws.cell(row=r, column=1).value
        if metric_name == "Total NOI":
            ws.cell(row=r, column=2).number_format = "$#,##0.00"
            ws.cell(row=r, column=3).number_format = "$#,##0.00"
        else:
            ws.cell(row=r, column=2).number_format = "0.0%"
            ws.cell(row=r, column=3).number_format = "0.0%"

    # Section C - Exception report
    ws.merge_cells("A20:G20")
    ws["A20"] = "Properties Requiring Attention — Most Recent Month"
    ws["A20"].font = Font(name="Calibri", size=12, bold=True, color=DARK_NAVY)
    ws["A20"].alignment = Alignment(horizontal="left", vertical="center")

    combined = (
        occ_latest[
            [
                "property_id",
                "property_name",
                "asset_type",
                "occupancy_rate",
            ]
        ]
        .merge(
            dlq_latest[
                ["property_id", "delinquency_rate"]
            ],
            on="property_id",
            how="inner",
        )
        .merge(
            fin_latest[
                ["property_id", "noi", "noi_margin"]
            ],
            on="property_id",
            how="inner",
        )
    )

    exceptions = combined[
        (combined["occupancy_rate"] < 0.90) | (combined["delinquency_rate"] > 0.08)
    ].copy()

    def risk_label(row):
        occ_risk = row["occupancy_rate"] < 0.90
        dlq_risk = row["delinquency_rate"] > 0.08
        if occ_risk and dlq_risk:
            return "HIGH RISK"
        if dlq_risk:
            return "DELINQUENCY RISK"
        if occ_risk:
            return "OCCUPANCY RISK"
        return "WATCH"

    priority = {"HIGH RISK": 1, "DELINQUENCY RISK": 2, "OCCUPANCY RISK": 3, "WATCH": 4}
    exceptions["Risk Label"] = exceptions.apply(risk_label, axis=1)
    exceptions["risk_order"] = exceptions["Risk Label"].map(priority)
    exceptions.sort_values(["risk_order", "delinquency_rate", "occupancy_rate"], ascending=[True, False, True], inplace=True)

    exception_out = exceptions[
        [
            "property_name",
            "asset_type",
            "occupancy_rate",
            "delinquency_rate",
            "noi",
            "noi_margin",
            "Risk Label",
        ]
    ].rename(
        columns={
            "property_name": "Property Name",
            "asset_type": "Asset Type",
            "occupancy_rate": "Occupancy Rate",
            "delinquency_rate": "Delinquency Rate",
            "noi": "NOI",
            "noi_margin": "NOI Margin",
        }
    )

    exception_header_row = 21
    write_dataframe(ws, exception_out, start_row=exception_header_row)
    style_custom_header(ws, exception_header_row, 1, 7)
    style_data_cells(ws, exception_header_row + 1, ws.max_row, 7)

    for r in range(exception_header_row + 1, exception_header_row + len(exception_out) + 1):
        ws.cell(row=r, column=3).number_format = "0.0%"
        ws.cell(row=r, column=4).number_format = "0.0%"
        ws.cell(row=r, column=5).number_format = "$#,##0.00"
        ws.cell(row=r, column=6).number_format = "0.0%"

        label_cell = ws.cell(row=r, column=7)
        label = label_cell.value
        if label == "HIGH RISK":
            label_cell.fill = PatternFill(fill_type="solid", fgColor=RED_FILL)
            label_cell.font = Font(name="Calibri", size=10, bold=True, color=DARK_RED_TEXT)
        elif label == "DELINQUENCY RISK":
            label_cell.fill = PatternFill(fill_type="solid", fgColor=ORANGE_FILL)
            label_cell.font = Font(name="Calibri", size=10, bold=True)
        elif label == "OCCUPANCY RISK":
            label_cell.fill = PatternFill(fill_type="solid", fgColor=YELLOW_FILL)
            label_cell.font = Font(name="Calibri", size=10, bold=True)
        else:
            label_cell.fill = PatternFill(fill_type="solid", fgColor=LIGHT_GRAY)
            label_cell.font = Font(name="Calibri", size=10)

    note_row = exception_header_row + len(exception_out) + 2
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=9)
    ws.cell(
        row=note_row,
        column=1,
        value="This exception report combines occupancy, collections risk, and NOI to help asset managers prioritize follow-up.",
    )
    ws.cell(row=note_row, column=1).font = Font(name="Calibri", size=10, italic=True, color="666666")

    # Section D - Top/Bottom 5 NOI Margin
    top_bottom_start = note_row + 2
    ws.merge_cells(start_row=top_bottom_start, start_column=1, end_row=top_bottom_start, end_column=3)
    ws.cell(row=top_bottom_start, column=1, value="Top 5 Properties by NOI Margin")
    ws.cell(row=top_bottom_start, column=1).font = Font(name="Calibri", size=12, bold=True, color=DARK_NAVY)

    ws.merge_cells(start_row=top_bottom_start, start_column=6, end_row=top_bottom_start, end_column=8)
    ws.cell(row=top_bottom_start, column=6, value="Bottom 5 Properties by NOI Margin")
    ws.cell(row=top_bottom_start, column=6).font = Font(name="Calibri", size=12, bold=True, color=DARK_NAVY)

    noi_rank = fin_latest[["property_id", "property_name", "asset_type", "noi_margin"]].copy()
    top5 = noi_rank.sort_values("noi_margin", ascending=False).head(5)
    bottom5 = noi_rank.sort_values("noi_margin", ascending=True).head(5)

    top_out = top5[["property_name", "asset_type", "noi_margin"]].rename(
        columns={"property_name": "Property Name", "asset_type": "Asset Type", "noi_margin": "NOI Margin"}
    )
    bottom_out = bottom5[["property_name", "asset_type", "noi_margin"]].rename(
        columns={"property_name": "Property Name", "asset_type": "Asset Type", "noi_margin": "NOI Margin"}
    )

    top_header = top_bottom_start + 1
    write_dataframe(ws, top_out, start_row=top_header, start_col=1)
    style_custom_header(ws, top_header, 1, 3)
    style_data_cells(ws, top_header + 1, top_header + len(top_out), 3)
    for r in range(top_header + 1, top_header + len(top_out) + 1):
        ws.cell(row=r, column=3).number_format = "0.0%"

    write_dataframe(ws, bottom_out, start_row=top_header, start_col=6)
    style_custom_header(ws, top_header, 6, 8)
    style_data_cells(ws, top_header + 1, top_header + len(bottom_out), 8)
    for r in range(top_header + 1, top_header + len(bottom_out) + 1):
        ws.cell(row=r, column=8).number_format = "0.0%"

    ws.freeze_panes = "A3"
    auto_fit_columns(ws)


def main():
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)

    properties_df = pd.read_csv(DATA_DIR / "properties.csv")
    occupancy_df = pd.read_csv(DATA_DIR / "monthly_occupancy.csv")
    delinquency_df = pd.read_csv(DATA_DIR / "delinquency.csv")
    financials_df = pd.read_csv(DATA_DIR / "property_financials.csv")
    market_df = pd.read_csv(DATA_DIR / "market_comps.csv")
    capex_df = pd.read_csv(DATA_DIR / "capex_projects.csv")

    wb = Workbook()
    wb.remove(wb.active)

    build_cover_tab(wb)
    build_properties_tab(wb, properties_df)
    build_occupancy_tab(wb, properties_df, occupancy_df)
    build_delinquency_tab(wb, properties_df, delinquency_df)
    build_financials_tab(wb, properties_df, financials_df)
    build_market_comps_tab(wb, properties_df, market_df)
    build_capex_tab(wb, properties_df, capex_df)
    build_dashboard_tab(wb, properties_df, occupancy_df, delinquency_df, financials_df)

    wb.save(OUTPUT_PATH)

    print("Updated workbook with Tabs 1–8: excel/portfolio_dashboard.xlsx")
    print(f"Sheets created: {', '.join(wb.sheetnames)}")


if __name__ == "__main__":
    main()
