WITH revenue_by_product AS (
  SELECT
    p.category,
    oi.product_id,
    SUM(oi.quantity * oi.price) AS revenue
  FROM orders o
  JOIN order_items oi ON oi.order_id = o.order_id
  JOIN products p ON p.product_id = oi.product_id
  WHERE o.status = 'delivered'
    AND o.created_at::date >= '2025-04-01'
    AND o.created_at::date <  '2025-07-01'
  GROUP BY p.category, oi.product_id
)
SELECT
  category,
  product_id,
  revenue,
  ROUND(
    (100 * revenue / SUM(revenue) OVER (PARTITION BY category))::numeric,
    2
  ) AS pct_of_category
FROM (
  SELECT
    *,
    ROW_NUMBER() OVER (PARTITION BY category ORDER BY revenue DESC) AS rn
  FROM revenue_by_product
) t
WHERE rn <= 3
ORDER BY category, rn;
