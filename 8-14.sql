SELECT
    product_id,
    revenue,
    revenue_cum_share,
    CASE
        WHEN revenue_cum_share <= 0.80 THEN 'A'
        WHEN revenue_cum_share <= 0.95 THEN 'B'
        ELSE 'C'
    END AS abc_class
FROM (
    SELECT
        product_id,
        SUM(revenue) AS revenue,
        SUM(SUM(revenue)) OVER (ORDER BY SUM(revenue) DESC) 
            / SUM(SUM(revenue)) OVER () AS revenue_cum_share
    FROM (
        SELECT
            oi.product_id,
            oi.price * oi.quantity AS revenue
        FROM order_items oi
        JOIN orders o ON oi.order_id = o.order_id
        WHERE o.created_at::date >= '2025-01-01'
          AND o.created_at::date < '2025-07-01'
          AND o.status = 'delivered'
    ) AS t
    GROUP BY product_id
) AS s
ORDER BY revenue DESC;
