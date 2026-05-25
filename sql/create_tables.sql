-- Real Estate Portfolio Performance & Risk Dashboard
-- SQLite-compatible star schema DDL
-- Business purpose: model monthly property performance for occupancy, collections risk, and NOI review.

PRAGMA foreign_keys = ON;

-- =========================================================
-- Dimension Tables
-- =========================================================

-- Property master data used by all fact tables.
CREATE TABLE IF NOT EXISTS dim_property (
    property_id TEXT PRIMARY KEY,
    property_name TEXT,
    asset_type TEXT,
    city TEXT,
    state TEXT,
    total_units INTEGER,
    year_acquired INTEGER,
    purchase_price REAL
);

-- Calendar attributes for monthly reporting and trend analysis.
CREATE TABLE IF NOT EXISTS dim_date (
    date_id INTEGER PRIMARY KEY,
    full_date DATE,
    year INTEGER,
    quarter INTEGER,
    month_num INTEGER,
    month_name TEXT,
    is_year_end TEXT
);

-- Asset type lookup to support grouped analysis by MHC vs Self-Storage.
CREATE TABLE IF NOT EXISTS dim_asset_type (
    asset_type_id INTEGER PRIMARY KEY,
    asset_type_name TEXT,
    asset_type_description TEXT
);

-- Market lookup for city/state/region-level rollups.
CREATE TABLE IF NOT EXISTS dim_market (
    market_id INTEGER PRIMARY KEY,
    city TEXT,
    state TEXT,
    region TEXT
);

-- =========================================================
-- Fact Tables
-- =========================================================

-- Grain: one row per property per month.
-- occupancy_flag logic:
--   'On Target'       if occupancy_rate >= 93
--   'Watch'           if occupancy_rate >= 88 and occupancy_rate < 93
--   'Below Threshold' if occupancy_rate < 88
CREATE TABLE IF NOT EXISTS fact_occupancy_monthly (
    fact_id INTEGER PRIMARY KEY,
    property_id TEXT,
    date_id INTEGER,
    occupied_units INTEGER,
    total_units INTEGER,
    occupancy_rate REAL,
    occupancy_flag TEXT,
    FOREIGN KEY (property_id) REFERENCES dim_property(property_id),
    FOREIGN KEY (date_id) REFERENCES dim_date(date_id)
);

-- Grain: one row per property per month.
-- risk_tier logic:
--   'Acceptable' if delinquency_rate < 5
--   'Monitor'    if delinquency_rate >= 5 and delinquency_rate < 8
--   'At Risk'    if delinquency_rate >= 8 and delinquency_rate < 12
--   'Critical'   if delinquency_rate >= 12
CREATE TABLE IF NOT EXISTS fact_delinquency_monthly (
    fact_id INTEGER PRIMARY KEY,
    property_id TEXT,
    date_id INTEGER,
    rent_billed REAL,
    rent_collected REAL,
    past_due_amount REAL,
    delinquency_rate REAL,
    risk_tier TEXT,
    FOREIGN KEY (property_id) REFERENCES dim_property(property_id),
    FOREIGN KEY (date_id) REFERENCES dim_date(date_id)
);

-- Grain: one row per property per month.
CREATE TABLE IF NOT EXISTS fact_property_financials (
    fact_id INTEGER PRIMARY KEY,
    property_id TEXT,
    date_id INTEGER,
    total_revenue REAL,
    operating_expenses REAL,
    noi REAL,
    noi_margin REAL,
    expense_ratio REAL,
    FOREIGN KEY (property_id) REFERENCES dim_property(property_id),
    FOREIGN KEY (date_id) REFERENCES dim_date(date_id)
);

-- Grain: one row per property per capital project.
CREATE TABLE IF NOT EXISTS fact_capex_projects (
    fact_id INTEGER PRIMARY KEY,
    property_id TEXT,
    project_name TEXT,
    project_type TEXT,
    budgeted_cost REAL,
    actual_cost REAL,
    variance REAL,
    status TEXT,
    FOREIGN KEY (property_id) REFERENCES dim_property(property_id)
);

-- Grain: one row per property.
-- Included explicitly so acquisition opportunity ranking queries can run.
CREATE TABLE IF NOT EXISTS market_comps (
    property_id TEXT PRIMARY KEY,
    market_avg_rent REAL,
    property_avg_rent REAL,
    rent_upside_pct REAL,
    competitor_occupancy REAL,
    market_cap_rate REAL,
    opportunity_score INTEGER,
    FOREIGN KEY (property_id) REFERENCES dim_property(property_id)
);
