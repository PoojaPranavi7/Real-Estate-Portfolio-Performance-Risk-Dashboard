# Real Estate Portfolio Performance & Risk Dashboard

## Project Overview
This project is an internal analytics solution built for a private equity real estate firm managing manufactured housing and self-storage assets. It is designed to help business teams monitor portfolio health, detect emerging risks, and prioritize value-add action. The analysis combines operational and financial performance across occupancy, delinquency, NOI, expense management, and capital execution. The model covers a 12-property portfolio and provides a practical framework for acquisition opportunity assessment and monthly performance review.

## Business Problem
This dashboard answers these three key questions:
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
The project uses a star schema design that separates descriptive dimensions from measurable operational and financial facts.

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

The core analytical grain is one row per property per month. This grain was chosen because asset management teams review occupancy, collections, revenue, expenses, and NOI on a recurring monthly basis, while still allowing rollups by asset type, market, state, and full portfolio.

## Key Metrics Tracked
| Metric | Formula | Business Purpose |
|---|---|---|
| Occupancy Rate | Occupied Units / Total Units | Measures property demand and lease-up effectiveness. |
| Delinquency Rate | Past Due Amount / Rent Billed | Identifies collections risk and revenue leakage. |
| Collection Rate | Rent Collected / Rent Billed | Tracks billing conversion and payment quality. |
| NOI | Total Revenue - Operating Expenses | Measures core operating profitability. |
| NOI Margin | NOI / Total Revenue | Compares profit efficiency across assets. |
| Expense Ratio | Operating Expenses / Total Revenue | Highlights cost pressure and operating efficiency gaps. |
| Revenue per Unit | Total Revenue / Occupied Units | Evaluates revenue productivity per occupied unit. |
| CapEx Variance | Actual Cost - Budgeted Cost | Monitors project budget performance and execution discipline. |

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
For this interview, I prepared a portfolio analytics project that demonstrates private equity real estate portfolio analytics across manufactured housing and self-storage assets. I structured the model at one row per property per month so performance can be reviewed in the same cadence used by operating and finance teams. The framework uses four diagnostic lenses—occupancy, delinquency, NOI, and CapEx—to separate demand issues from collections and cost issues, while surfacing value-add opportunity. This work matters because acquisitions, asset management, and finance all need a shared, decision-ready view of property performance to prioritize interventions and capital allocation.
