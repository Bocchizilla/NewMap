SELECT
    gmvs.order_date,
    gmvs.daily_gmv,
    ROUND(
        AVG(gmvs.daily_gmv) OVER (
            ORDER BY gmvs.order_date
            ROWS BETWEEN 6 PRECEDING AND CURRENT ROW
        )::numeric,
        2
    ) AS moving_avg_7d
FROM (
    SELECT
        o.created_at::date AS order_date,
        SUM(o.total_amount) AS daily_gmv
    FROM orders o
    WHERE o.status = 'delivered'
      AND o.created_at::date >= '2025-05-01'
      AND o.created_at::date < '2025-08-01'
    GROUP BY o.created_at::date
) gmvs
ORDER BY gmvs.order_date;
