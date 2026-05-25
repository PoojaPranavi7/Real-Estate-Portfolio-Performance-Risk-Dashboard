from pathlib import Path

import pandas as pd
from openpyxl import Workbook
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

THIN_BORDER = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)


def normalize_pct_value(value):
    if pd.isna(value):
        return value
    return value / 100 if value > 1.5 else value


def normalize_pct_series(series: pd.Series) -> pd.Series:
    return series.apply(normalize_pct_value)


def auto_fit_columns(ws, min_width=12, max_width=35):
    for col_cells in ws.columns:
        col_letter = col_cells[0].column_letter
        max_len = 0
        for cell in col_cells:
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


def write_dataframe(ws, df: pd.DataFrame):
    ws.append(list(df.columns))
    for row in df.itertuples(index=False):
        ws.append(list(row))


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

    write_dataframe(ws, df)
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

    write_dataframe(ws, out)
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

    write_dataframe(ws, out)
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

    write_dataframe(ws, out)
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

    write_dataframe(ws, out)
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


def main():
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)

    properties_df = pd.read_csv(DATA_DIR / "properties.csv")
    occupancy_df = pd.read_csv(DATA_DIR / "monthly_occupancy.csv")
    delinquency_df = pd.read_csv(DATA_DIR / "delinquency.csv")
    financials_df = pd.read_csv(DATA_DIR / "property_financials.csv")
    market_df = pd.read_csv(DATA_DIR / "market_comps.csv")
    pd.read_csv(DATA_DIR / "capex_projects.csv")  # read as required input source

    wb = Workbook()
    wb.remove(wb.active)

    build_cover_tab(wb)
    build_properties_tab(wb, properties_df)
    build_occupancy_tab(wb, properties_df, occupancy_df)
    build_delinquency_tab(wb, properties_df, delinquency_df)
    build_financials_tab(wb, properties_df, financials_df)
    build_market_comps_tab(wb, properties_df, market_df)

    wb.save(OUTPUT_PATH)

    print(f"Workbook created: {OUTPUT_PATH}")
    print(f"Sheets created: {', '.join(wb.sheetnames)}")


if __name__ == "__main__":
    main()
