-- =========================================================
-- QUERY 1 — Monthly Occupancy Rate by Property
-- This query helps asset managers quickly see which properties are below threshold.
-- =========================================================
SELECT
    p.property_name,
    p.asset_type,
    d.full_date AS month,
    o.occupied_units,
    o.total_units,
    o.occupancy_rate,
    o.occupancy_flag
FROM fact_occupancy_monthly o
JOIN dim_property p
    ON p.property_id = o.property_id
JOIN dim_date d
    ON d.date_id = o.date_id
ORDER BY d.full_date ASC, o.occupancy_rate ASC;


-- =========================================================
-- QUERY 2 — Delinquency Rate by Property (worst to best)
-- Business purpose: rank collection risk and identify where past due balances are increasing.
-- Uses NULLIF to safely avoid divide-by-zero when rent_billed is zero.
-- =========================================================
SELECT
    p.property_name,
    d.full_date AS month,
    dl.rent_billed,
    dl.rent_collected,
    dl.past_due_amount,
    ROUND(dl.past_due_amount / NULLIF(dl.rent_billed, 0) * 100, 2) AS delinquency_rate,
    dl.risk_tier
FROM fact_delinquency_monthly dl
JOIN dim_property p
    ON p.property_id = dl.property_id
JOIN dim_date d
    ON d.date_id = dl.date_id
ORDER BY delinquency_rate DESC;


-- =========================================================
-- QUERY 3 — NOI and NOI Margin by Property
-- Sorting NOI margin ascending brings the weakest NOI margin properties to the top for review.
-- =========================================================
SELECT
    p.property_name,
    d.full_date AS month,
    f.total_revenue,
    f.operating_expenses,
    ROUND(f.total_revenue - f.operating_expenses, 2) AS noi_calculated,
    f.noi_margin,
    f.expense_ratio
FROM fact_property_financials f
JOIN dim_property p
    ON p.property_id = f.property_id
JOIN dim_date d
    ON d.date_id = f.date_id
ORDER BY f.noi_margin ASC;


-- =========================================================
-- QUERY 4 — Underperforming Asset Exception Report
-- This is the core asset manager exception report because it combines occupancy,
-- collections risk, and NOI into one review.
-- =========================================================
SELECT
    p.property_name,
    p.asset_type,
    d.full_date AS month,
    o.occupancy_rate,
    dl.delinquency_rate,
    f.noi,
    CASE
        WHEN o.occupancy_rate < 90 AND dl.delinquency_rate > 8 THEN 'High Risk'
        WHEN o.occupancy_rate < 90 THEN 'Occupancy Risk'
        WHEN dl.delinquency_rate > 8 THEN 'Delinquency Risk'
        ELSE 'Watch'
    END AS risk_label
FROM fact_occupancy_monthly o
JOIN fact_delinquency_monthly dl
    ON dl.property_id = o.property_id
   AND dl.date_id = o.date_id
JOIN fact_property_financials f
    ON f.property_id = o.property_id
   AND f.date_id = o.date_id
JOIN dim_property p
    ON p.property_id = o.property_id
JOIN dim_date d
    ON d.date_id = o.date_id
WHERE o.occupancy_rate < 90
   OR dl.delinquency_rate > 8
ORDER BY dl.delinquency_rate DESC, o.occupancy_rate ASC;


-- =========================================================
-- QUERY 5 — BONUS: Acquisition Opportunity Ranking
-- Business purpose: rank top value-add candidates using rent upside, market occupancy gap,
-- and precomputed opportunity score from market comps.
-- Uses latest month occupancy from fact_occupancy_monthly.
-- =========================================================
WITH latest_month AS (
    SELECT MAX(date_id) AS max_date_id
    FROM fact_occupancy_monthly
),
latest_occupancy AS (
    SELECT
        o.property_id,
        o.occupancy_rate AS latest_occupancy_rate
    FROM fact_occupancy_monthly o
    JOIN latest_month lm
        ON lm.max_date_id = o.date_id
)
SELECT
    p.property_name,
    p.asset_type,
    mc.rent_upside_pct,
    ROUND(mc.competitor_occupancy - lo.latest_occupancy_rate, 4) AS occupancy_gap,
    mc.opportunity_score
FROM dim_property p
JOIN market_comps mc
    ON mc.property_id = p.property_id
JOIN latest_occupancy lo
    ON lo.property_id = p.property_id
ORDER BY mc.opportunity_score DESC
LIMIT 5;
