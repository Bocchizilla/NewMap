WITH sushi_buyers AS (
    SELECT DISTINCT o.user_id
    FROM orders o
    JOIN order_items oi
        ON o.order_id = oi.order_id
    JOIN products p
        ON oi.product_id = p.product_id
    WHERE o.status = 'delivered'
      AND o.created_at::date >= '2025-04-01'
      AND o.created_at::date < '2025-07-01'
      AND p.category = 'Sushi'
),
pizza_buyers AS (
    SELECT DISTINCT o.user_id
    FROM orders o
    JOIN order_items oi
        ON o.order_id = oi.order_id
    JOIN products p
        ON oi.product_id = p.product_id
    WHERE o.status = 'delivered'
      AND o.created_at::date >= '2025-04-01'
      AND o.created_at::date < '2025-07-01'
      AND p.category = 'Pizza'
)
SELECT
    COUNT(*) AS both_count,
    COUNT(*)::float
        / NULLIF(
            (
                SELECT COUNT(DISTINCT user_id)
                FROM orders
                WHERE status = 'delivered'
                  AND created_at::date >= '2025-04-01'
                  AND created_at::date < '2025-07-01'
            ),
            0
        ) AS both_share
FROM (
    SELECT s.user_id
    FROM sushi_buyers s
    JOIN pizza_buyers p
        ON s.user_id = p.user_id
) t;
