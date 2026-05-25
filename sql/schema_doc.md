# Real Estate Portfolio Star Schema Design

This model is designed for a portfolio analytics workflow where asset management, acquisitions, and finance all need a consistent monthly view of property performance.

## 1) Star Schema Design

The warehouse uses a star schema with shared dimensions and subject-area fact tables:

- **Dimensions**
  - `dim_property`: property master attributes (name, asset type, location, unit count, acquisition profile).
  - `dim_date`: calendar attributes used for monthly reporting and trend analysis.
  - `dim_asset_type`: standardized asset type definitions (MHC vs Self-Storage).
  - `dim_market`: market grouping by city/state/region for location-based analysis.
- **Facts**
  - `fact_occupancy_monthly`: occupancy performance and occupancy status flags.
  - `fact_delinquency_monthly`: billing, collections, past due, and delinquency risk tiers.
  - `fact_property_financials`: revenue, expenses, NOI, NOI margin, and expense ratio.
  - `fact_capex_projects`: project-level capital spending and variance tracking.
  - `market_comps`: one-row-per-property market rent and opportunity scoring for acquisition prioritization.

This layout keeps operational metrics separated by business process (occupancy, collections, finance, capital projects) while making cross-functional analysis straightforward through shared keys (`property_id`, `date_id`).

## 2) Grain Decision: One Row per Property per Month

The core grain is **one row per property per month** for occupancy, delinquency, and financial fact tables.

This grain was chosen because monthly is the cadence used for:

- Asset manager performance reviews
- Investor reporting
- Month-over-month trend analysis
- Exception monitoring (low occupancy, high delinquency, weak NOI margin)

A monthly property-level grain is detailed enough to diagnose underperformance and stable enough to avoid noise from daily operational fluctuations.

## 3) Roles of Dimension vs Fact Tables

- **Dimension tables** provide descriptive context used to slice and group metrics (for example by market, state, asset type, or quarter).
- **Fact tables** store measurable events and outcomes at defined grain (occupancy rates, rent billed/collected, NOI, project cost variance).

In practice:

- `dim_property` + `fact_occupancy_monthly` answers: *Which assets are below occupancy threshold this month?*
- `dim_property` + `fact_delinquency_monthly` answers: *Where is collections risk accelerating?*
- `dim_property` + `fact_property_financials` answers: *Which properties have weak NOI margins and high expense ratios?*
- `dim_property` + `market_comps` answers: *Which assets have strongest value-add upside?*

## 4) Why `NULLIF` Is Used in Delinquency Calculations

Delinquency rate is typically calculated as:

`past_due_amount / rent_billed`

If `rent_billed` is `0`, a direct division causes a runtime divide-by-zero error. Using:

`past_due_amount / NULLIF(rent_billed, 0)`

returns `NULL` instead of error in zero-billed edge cases. This keeps analysis queries robust and prevents reporting pipelines from failing due to rare data conditions.

## 5) How This Model Supports Each Team

- **Acquisitions**
  - Uses `market_comps` and latest occupancy to rank value-add opportunities.
  - Compares in-place rent vs market rent and identifies occupancy gaps by asset.
- **Asset Management**
  - Uses occupancy and delinquency exception reports to identify underperforming properties.
  - Tracks operational risk using threshold flags (`On Target`, `Watch`, `Below Threshold`, `At Risk`, `Critical`).
- **Finance**
  - Monitors NOI, NOI margin, and expense ratio trends monthly.
  - Ties financial outcomes to operational health and CapEx execution.

Overall, the schema is intentionally practical: it allows teams to move from raw monthly metrics to action-oriented decisions quickly.
