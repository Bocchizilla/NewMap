WITH base AS (
  SELECT
    oi.order_id,
    COUNT(DISTINCT p.category) AS cat_cnt,
    SUM(oi.quantity) AS qty_sum
  FROM order_items oi
  JOIN products p ON p.product_id = oi.product_id
  GROUP BY oi.order_id
),
orders_july AS (
  SELECT order_id
  FROM orders
  WHERE status = 'delivered'
    AND created_at::date >= DATE '2025-07-01'
    AND created_at::date < DATE '2025-08-01'
)
SELECT
  COUNT(*) AS orders_cnt,
  COUNT(CASE WHEN b.cat_cnt = 1 AND b.qty_sum >= 2 THEN 1 END) AS ideal_cnt,
  ROUND(
    100.0
    * COUNT(CASE WHEN b.cat_cnt = 1 AND b.qty_sum >= 2 THEN 1 END)
    / NULLIF(COUNT(*), 0),
    2
  ) AS ideal_pct
FROM orders_july oj
JOIN base b ON b.order_id = oj.order_id;
