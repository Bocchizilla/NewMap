SELECT
    DATE_TRUNC('month', u.created_at::timestamp) AS reg_month,
    COUNT(DISTINCT u.user_id) AS users_total,
    COUNT(DISTINCT CASE
        WHEN o_cnt.orders_count >= 2 THEN u.user_id
    END)::float / COUNT(DISTINCT u.user_id) AS share_users_2plus_orders
FROM users u
LEFT JOIN (
    SELECT
        o.user_id,
        COUNT(o.order_id) AS orders_count
    FROM orders o
    JOIN users u2
        ON o.user_id = u2.user_id
    WHERE o.status = 'delivered'
      AND DATE_TRUNC('month', o.created_at::timestamp)
          = DATE_TRUNC('month', u2.created_at::timestamp)
    GROUP BY o.user_id
) o_cnt
ON u.user_id = o_cnt.user_id
GROUP BY DATE_TRUNC('month', u.created_at::timestamp)
ORDER BY reg_month;
