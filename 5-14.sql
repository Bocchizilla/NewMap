WITH cohort_orders AS (
  SELECT
    u.user_id,
    DATE_TRUNC('month', u.created_at::timestamp) AS cohort_month,
    COUNT(o.order_id) FILTER (
      WHERE o.status = 'delivered'
        AND o.created_at::timestamp >= DATE_TRUNC('month', u.created_at::timestamp)
        AND o.created_at::timestamp <  DATE_TRUNC('month', u.created_at::timestamp) + INTERVAL '1 month'
    ) AS orders_in_first_month
  FROM users u
  LEFT JOIN orders o ON o.user_id = u.user_id
  GROUP BY
    u.user_id,
    DATE_TRUNC('month', u.created_at::timestamp)
)
SELECT
  cohort_month::date,
  COUNT(*) AS users_total,
  COUNT(*) FILTER (WHERE orders_in_first_month >= 2) AS users_2plus,
  ROUND(
    100.0 * (COUNT(*) FILTER (WHERE orders_in_first_month >= 2))::numeric
    / NULLIF(COUNT(*)::numeric, 0),
    2
  ) AS pct_2plus
FROM cohort_orders
WHERE cohort_month >= DATE '2024-01-01'
  AND cohort_month <  DATE '2025-08-01'
GROUP BY cohort_month
ORDER BY cohort_month;
