import random
from pathlib import Path

import numpy as np
import pandas as pd


random.seed(42)
np.random.seed(42)


REPO_ROOT = Path(__file__).resolve().parent
DATA_DIR = REPO_ROOT / "data"


def clamp(value: float, low: float, high: float) -> float:
    return max(low, min(high, value))


def build_properties() -> pd.DataFrame:
    properties = [
        ("MHC-01", "Sunrise Meadows", "MHC", "Las Vegas", "NV", 120, 2018, 9_400_000),
        ("MHC-02", "Desert Palms", "MHC", "Henderson", "NV", 95, 2019, 7_100_000),
        ("MHC-03", "Canyon Ridge", "MHC", "Tucson", "AZ", 110, 2020, 6_900_000),
        ("MHC-04", "Saguaro Estates", "MHC", "Phoenix", "AZ", 140, 2017, 12_800_000),
        ("MHC-05", "Pinon Pines", "MHC", "Albuquerque", "NM", 88, 2021, 5_950_000),
        ("MHC-06", "Red Rock Commons", "MHC", "St. George", "UT", 102, 2018, 7_850_000),
        ("MHC-07", "Mesa Vista", "MHC", "Mesa", "AZ", 115, 2019, 8_600_000),
        ("STG-01", "ClearBox Storage", "Self-Storage", "Las Vegas", "NV", 280, 2016, 10_500_000),
        ("STG-02", "SunState Storage", "Self-Storage", "Scottsdale", "AZ", 320, 2018, 13_900_000),
        ("STG-03", "Desert Vault", "Self-Storage", "Tucson", "AZ", 240, 2017, 9_850_000),
        ("STG-04", "Basin Self Storage", "Self-Storage", "Reno", "NV", 195, 2022, 7_400_000),
        ("STG-05", "Summit Storage", "Self-Storage", "Salt Lake City", "UT", 260, 2019, 11_250_000),
    ]
    columns = [
        "property_id",
        "property_name",
        "asset_type",
        "city",
        "state",
        "total_units",
        "year_acquired",
        "purchase_price",
    ]
    return pd.DataFrame(properties, columns=columns)


def build_monthly_occupancy(properties_df: pd.DataFrame) -> pd.DataFrame:
    months = pd.date_range("2024-01-01", "2024-12-01", freq="MS")

    occupancy_rate_targets = {
        "MHC-01": (0.90, 0.92),
        "MHC-02": (0.88, 0.91),
        "MHC-03": (0.78, 0.83),  # Problem property
        "MHC-04": (0.95, 0.98),  # Star property
        "MHC-05": (0.81, 0.89),  # Value-add improving
        "MHC-06": (0.91, 0.93),  # Stable
        "MHC-07": (0.93, 0.95),
        "STG-01": (0.91, 0.94),  # Strong
        "STG-02": (0.89, 0.92),  # Good occupancy
        "STG-03": (0.89, 0.91),
        "STG-04": (0.76, 0.87),  # Recovering
        "STG-05": (0.88, 0.91),  # Steady
    }

    records = []
    for _, prop in properties_df.iterrows():
        pid = prop["property_id"]
        units = int(prop["total_units"])
        start_rate, end_rate = occupancy_rate_targets[pid]
        trend = np.linspace(start_rate, end_rate, len(months))

        for idx, month in enumerate(months):
            base_occupied = int(round(units * trend[idx]))
            occupied_units = int(clamp(base_occupied + random.randint(-2, 2), 0, units))
            occupancy_rate = occupied_units / units if units else 0.0
            records.append(
                {
                    "property_id": pid,
                    "month": month.strftime("%Y-%m-%d"),
                    "occupied_units": occupied_units,
                    "total_units": units,
                    "occupancy_rate": round(occupancy_rate, 4),
                }
            )

    return pd.DataFrame(records)


def build_delinquency(
    properties_df: pd.DataFrame, occupancy_df: pd.DataFrame
) -> tuple[pd.DataFrame, dict[str, float]]:
    avg_rents = {}
    for _, prop in properties_df.iterrows():
        if prop["asset_type"] == "MHC":
            avg_rents[prop["property_id"]] = round(random.uniform(650, 850), 2)
        else:
            avg_rents[prop["property_id"]] = round(random.uniform(90, 160), 2)

    delinquency_rate_targets = {
        "MHC-01": (0.05, 0.06),
        "MHC-02": (0.06, 0.07),
        "MHC-03": (0.12, 0.16),  # High delinquency
        "MHC-04": (0.02, 0.03),  # Very low delinquency
        "MHC-05": (0.09, 0.07),  # Improving as stabilization progresses
        "MHC-06": (0.05, 0.06),
        "MHC-07": (0.05, 0.11),  # Rising delinquency
        "STG-01": (0.03, 0.04),
        "STG-02": (0.07, 0.13),  # Rising delinquency puzzle
        "STG-03": (0.05, 0.07),
        "STG-04": (0.09, 0.07),
        "STG-05": (0.06, 0.07),
    }

    monthly_order = sorted(occupancy_df["month"].unique())
    monthly_index = {month: idx for idx, month in enumerate(monthly_order)}
    month_count = len(monthly_order)

    records = []
    for _, row in occupancy_df.iterrows():
        pid = row["property_id"]
        month = row["month"]
        occupied_units = int(row["occupied_units"])
        avg_rent = avg_rents[pid]
        rent_billed = round(occupied_units * avg_rent, 2)

        start_rate, end_rate = delinquency_rate_targets[pid]
        idx = monthly_index[month]
        if month_count > 1:
            base_rate = start_rate + ((end_rate - start_rate) * idx / (month_count - 1))
        else:
            base_rate = start_rate
        noisy_rate = clamp(base_rate + random.uniform(-0.006, 0.006), 0.01, 0.25)

        past_due_amount = round(rent_billed * noisy_rate, 2)
        rent_collected = round(rent_billed - past_due_amount, 2)
        delinquency_rate = past_due_amount / rent_billed if rent_billed else 0.0

        records.append(
            {
                "property_id": pid,
                "month": month,
                "rent_billed": rent_billed,
                "rent_collected": rent_collected,
                "past_due_amount": past_due_amount,
                "delinquency_rate": round(delinquency_rate, 4),
            }
        )

    return pd.DataFrame(records), avg_rents


def build_financials(delinquency_df: pd.DataFrame) -> pd.DataFrame:
    expense_ratio_targets = {
        "MHC-01": (0.43, 0.48),
        "MHC-02": (0.44, 0.50),
        "MHC-03": (0.49, 0.52),
        "MHC-04": (0.38, 0.42),  # Strong NOI
        "MHC-05": (0.50, 0.52),
        "MHC-06": (0.44, 0.47),
        "MHC-07": (0.46, 0.50),
        "STG-01": (0.47, 0.53),
        "STG-02": (0.51, 0.58),
        "STG-03": (0.62, 0.68),  # Expense-heavy
        "STG-04": (0.56, 0.61),
        "STG-05": (0.50, 0.56),
    }

    monthly_order = sorted(delinquency_df["month"].unique())
    monthly_index = {month: idx for idx, month in enumerate(monthly_order)}
    month_count = len(monthly_order)

    records = []
    for _, row in delinquency_df.iterrows():
        pid = row["property_id"]
        month = row["month"]
        rent_collected = float(row["rent_collected"])
        other_income_pct = random.uniform(0.05, 0.08)
        other_income = rent_collected * other_income_pct
        total_revenue = round(rent_collected + other_income, 2)

        start_ratio, end_ratio = expense_ratio_targets[pid]
        idx = monthly_index[month]
        if month_count > 1:
            base_ratio = start_ratio + ((end_ratio - start_ratio) * idx / (month_count - 1))
        else:
            base_ratio = start_ratio
        noisy_ratio = clamp(base_ratio + random.uniform(-0.01, 0.01), 0.35, 0.75)

        operating_expenses = round(total_revenue * noisy_ratio, 2)
        noi = round(total_revenue - operating_expenses, 2)
        noi_margin = noi / total_revenue if total_revenue else 0.0
        expense_ratio = operating_expenses / total_revenue if total_revenue else 0.0

        records.append(
            {
                "property_id": pid,
                "month": month,
                "total_revenue": total_revenue,
                "operating_expenses": operating_expenses,
                "noi": noi,
                "noi_margin": round(noi_margin, 4),
                "expense_ratio": round(expense_ratio, 4),
            }
        )

    return pd.DataFrame(records)


def build_market_comps(
    properties_df: pd.DataFrame,
    occupancy_df: pd.DataFrame,
    delinquency_df: pd.DataFrame,
    financials_df: pd.DataFrame,
    avg_rents: dict[str, float],
) -> pd.DataFrame:
    upside_targets = {
        "MHC-01": random.uniform(1, 5),
        "MHC-02": random.uniform(2, 6),
        "MHC-03": random.uniform(15, 20),  # Highest upside
        "MHC-04": random.uniform(3, 6),
        "MHC-05": random.uniform(13, 18),  # Highest upside
        "MHC-06": random.uniform(2, 5),
        "MHC-07": random.uniform(4, 8),
        "STG-01": random.uniform(1, 4),
        "STG-02": random.uniform(6, 10),
        "STG-03": random.uniform(5, 9),
        "STG-04": random.uniform(10, 14),
        "STG-05": random.uniform(3, 7),
    }

    cap_rate_ranges = {
        "MHC": (5.1, 7.0),
        "Self-Storage": (4.8, 6.6),
    }

    avg_occ = occupancy_df.groupby("property_id", as_index=True)["occupancy_rate"].mean()
    avg_delinq = delinquency_df.groupby("property_id", as_index=True)["delinquency_rate"].mean()
    avg_expense = financials_df.groupby("property_id", as_index=True)["expense_ratio"].mean()

    records = []
    for _, prop in properties_df.iterrows():
        pid = prop["property_id"]
        asset_type = prop["asset_type"]
        property_avg_rent = avg_rents[pid]
        rent_upside_pct = round(float(upside_targets[pid]), 2)
        market_avg_rent = round(property_avg_rent * (1 + rent_upside_pct / 100), 2)

        competitor_occupancy = round(
            clamp(float(avg_occ[pid]) + random.uniform(0.01, 0.05), 0.80, 0.98), 4
        )
        market_cap_rate = round(random.uniform(*cap_rate_ranges[asset_type]), 3)

        occupancy_upside = max(0.0, (0.93 - float(avg_occ[pid])) * 100)
        delinquency_risk = max(0.0, (float(avg_delinq[pid]) * 100) - 6)
        expense_risk = max(0.0, (float(avg_expense[pid]) * 100) - 52)
        base_score = (
            42
            + (rent_upside_pct * 2.8)
            + (occupancy_upside * 1.8)
            - (delinquency_risk * 1.4)
            - (expense_risk * 1.1)
        )
        opportunity_score = int(round(clamp(base_score, 0, 100)))

        records.append(
            {
                "property_id": pid,
                "market_avg_rent": market_avg_rent,
                "property_avg_rent": round(property_avg_rent, 2),
                "rent_upside_pct": rent_upside_pct,
                "competitor_occupancy": competitor_occupancy,
                "market_cap_rate": market_cap_rate,
                "opportunity_score": opportunity_score,
            }
        )

    market_df = pd.DataFrame(records)
    market_df.sort_values(by="opportunity_score", ascending=False, inplace=True)
    return market_df


def build_capex_projects(properties_df: pd.DataFrame) -> pd.DataFrame:
    project_templates = {
        "MHC-01": ("Entry Road Regrade", "Road Repair", "2024-02-15", "2024-05-30", 180000, "Completed"),
        "MHC-02": ("Perimeter Fence Refresh", "Fencing", "2024-03-01", "2024-06-20", 95000, "Completed"),
        "MHC-03": ("Drainage and Pad Stabilization", "Paving", "2024-01-20", "2024-09-15", 420000, "Active"),
        "MHC-04": ("Clubhouse Modernization", "Clubhouse Renovation", "2024-04-10", "2024-08-30", 260000, "Completed"),
        "MHC-05": ("Community-Wide Unit Upgrade Program", "Unit Upgrade", "2024-01-10", "2024-12-20", 780000, "Active"),
        "MHC-06": ("Domestic Water Loop Upgrade", "Water System", "2024-05-01", "2024-10-31", 310000, "Active"),
        "MHC-07": ("Monument Sign Replacement", "Signage", "2024-06-05", "2024-07-25", 42000, "Completed"),
        "STG-01": ("Exterior Lighting and CCTV", "HVAC", "2024-02-12", "2024-04-30", 165000, "Completed"),
        "STG-02": ("Drive Aisle Resurfacing", "Paving", "2024-03-20", "2024-08-10", 240000, "Completed"),
        "STG-03": ("HVAC Retrofit for Climate Units", "HVAC", "2024-01-15", "2024-11-15", 520000, "Active"),
        "STG-04": ("Security Gate and Fence Overhaul", "Fencing", "2024-04-01", "2024-09-30", 210000, "Completed"),
        "STG-05": ("Leasing Office Interior Refresh", "Clubhouse Renovation", "2024-07-01", "2024-10-30", 125000, "Completed"),
    }

    records = []
    for _, prop in properties_df.iterrows():
        pid = prop["property_id"]
        project_name, project_type, start_date, end_date, budgeted_cost, status = project_templates[pid]

        if pid in {"MHC-03", "MHC-05", "STG-03", "STG-02"}:
            variance_pct = random.uniform(0.04, 0.14)  # over budget
        elif pid in {"MHC-04", "STG-04", "STG-05"}:
            variance_pct = random.uniform(-0.10, -0.03)  # under budget
        else:
            variance_pct = random.uniform(-0.03, 0.05)

        actual_cost = round(budgeted_cost * (1 + variance_pct), 2)
        variance = round(actual_cost - budgeted_cost, 2)
        records.append(
            {
                "property_id": pid,
                "project_name": project_name,
                "project_type": project_type,
                "start_date": start_date,
                "end_date": end_date,
                "budgeted_cost": float(budgeted_cost),
                "actual_cost": actual_cost,
                "status": status,
                "variance": variance,
            }
        )

    return pd.DataFrame(records)


def main() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    properties_df = build_properties()
    occupancy_df = build_monthly_occupancy(properties_df)
    delinquency_df, avg_rents = build_delinquency(properties_df, occupancy_df)
    financials_df = build_financials(delinquency_df)
    market_comps_df = build_market_comps(
        properties_df, occupancy_df, delinquency_df, financials_df, avg_rents
    )
    capex_df = build_capex_projects(properties_df)

    properties_df.to_csv(DATA_DIR / "properties.csv", index=False)
    occupancy_df.to_csv(DATA_DIR / "monthly_occupancy.csv", index=False)
    delinquency_df.to_csv(DATA_DIR / "delinquency.csv", index=False)
    financials_df.to_csv(DATA_DIR / "property_financials.csv", index=False)
    market_comps_df.to_csv(DATA_DIR / "market_comps.csv", index=False)
    capex_df.to_csv(DATA_DIR / "capex_projects.csv", index=False)

    total_rows = sum(
        [
            len(properties_df),
            len(occupancy_df),
            len(delinquency_df),
            len(financials_df),
            len(market_comps_df),
            len(capex_df),
        ]
    )
    print(f"Generated {total_rows} rows across 6 CSV files.")


if __name__ == "__main__":
    main()
