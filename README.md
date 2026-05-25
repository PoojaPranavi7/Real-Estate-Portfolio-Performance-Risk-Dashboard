# Real Estate Portfolio Performance & Risk Dashboard

## Project Overview
This project is an internal analytics tool designed for a private equity real estate firm managing manufactured housing and self-storage assets. It provides a structured view of operating performance, risk, and value creation opportunities across the portfolio. The analysis combines monthly operational and financial reporting so teams can evaluate occupancy, delinquency, NOI, expenses, CapEx execution, and acquisition upside in one place. The current model covers a 12-property portfolio with consistent reporting outputs for asset management, acquisitions, and finance.

## Business Problem
This dashboard answers three key questions:
1. Which properties are performing well?
2. Which properties are at operational or financial risk?
3. Where is the value-add opportunity?

## Portfolio Snapshot
| Property ID | Property Name | Type | City | State | Units |
|---|---|---|---|---|---|
| MHC-01 | Sunrise Meadows | MHC | Las Vegas | NV | 120 |
| MHC-02 | Desert Palms | MHC | Henderson | NV | 95 |
| MHC-03 | Canyon Ridge | MHC | Tucson | AZ | 110 |
| MHC-04 | Saguaro Estates | MHC | Phoenix | AZ | 140 |
| MHC-05 | Pinon Pines | MHC | Albuquerque | NM | 88 |
| MHC-06 | Red Rock Commons | MHC | St. George | UT | 102 |
| MHC-07 | Mesa Vista | MHC | Mesa | AZ | 115 |
| STG-01 | ClearBox Storage | Self-Storage | Las Vegas | NV | 280 |
| STG-02 | SunState Storage | Self-Storage | Scottsdale | AZ | 320 |
| STG-03 | Desert Vault | Self-Storage | Tucson | AZ | 240 |
| STG-04 | Basin Self Storage | Self-Storage | Reno | NV | 195 |
| STG-05 | Summit Storage | Self-Storage | Salt Lake City | UT | 260 |

## Data Model
The project uses a star schema design to support clean, repeatable analytics across property operations and financial performance.

Dimension tables:
- dim_property
- dim_date
- dim_asset_type
- dim_market

Fact tables:
- fact_occupancy_monthly
- fact_delinquency_monthly
- fact_property_financials
- fact_capex_projects

The core analytical grain is one row per property per month. This grain was chosen because asset management teams review occupancy, collections, revenue, expenses, and NOI on a recurring monthly basis, while still allowing rollups by asset type, market, state, and portfolio.

## Key Metrics Tracked
| Metric | Formula | Business Purpose |
|---|---|---|
| Occupancy Rate | Occupied Units / Total Units | Measures demand and lease-up performance by property. |
| Delinquency Rate | Past Due Amount / Rent Billed | Identifies collections risk and revenue leakage. |
| Collection Rate | Rent Collected / Rent Billed | Tracks payment quality and billing conversion. |
| NOI | Total Revenue - Operating Expenses | Measures operating profitability before financing. |
| NOI Margin | NOI / Total Revenue | Compares profit efficiency across properties. |
| Expense Ratio | Operating Expenses / Total Revenue | Flags cost pressure and operational inefficiency. |
| Revenue per Unit | Total Revenue / Occupied Units | Evaluates revenue productivity at occupied-unit level. |
| CapEx Variance | Actual Cost - Budgeted Cost | Monitors capital project budget discipline. |

## Key Findings
- STG-02 SunState Storage shows healthy occupancy but rising delinquency, indicating a collections issue rather than a demand issue.
- MHC-03 Canyon Ridge is the highest-risk asset because it combines low occupancy with high delinquency.
- MHC-04 Saguaro Estates is the strongest benchmark property, with high occupancy, low delinquency, and strong NOI performance.
- MHC-05 Pinon Pines is the primary value-add opportunity, with improving occupancy and rent upside that should be tracked against the pro forma.
- STG-03 Desert Vault has NOI pressure driven by a high expense ratio, despite adequate occupancy.
- MHC-07 Mesa Vista has strong occupancy but rising delinquency, making it an early-warning collections risk.

## Tools Used
- SQL
- Microsoft Excel using openpyxl
- Python using pandas and numpy
- Power BI optional

## How to Run
1. Clone the repo
2. Install dependencies:  
   `pip install pandas openpyxl numpy`
3. Generate the synthetic data:  
   `python generate_data.py`
4. Build the Excel workbook:  
   `python build_excel.py`
5. Open:  
   `excel/portfolio_dashboard.xlsx`

## 60-Second Pitch
For this interview, I prepared a portfolio analytics project that simulates how a private equity real estate team would monitor performance and risk across manufactured housing and self-storage assets. I structured the data model at a monthly analytical grain, one row per property per month, so the analysis aligns with how teams actually review operating results and investor reporting cycles. The dashboard applies four diagnostic lenses—occupancy, delinquency, NOI, and CapEx—to quickly surface what is performing, what is drifting, and where operational action is needed. The output is designed to be practical for acquisitions, asset management, and finance so decisions can be made faster with a shared, data-driven view of portfolio health.
