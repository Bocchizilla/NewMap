SELECT
  COUNT(*) AS orders_cnt,
  ROUND(
    SUM(
      CASE
        WHEN u.is_vip = true THEN o.total_amount * 0.98
        ELSE o.total_amount
      END
    )::numeric,
    2
  ) AS gmv_adj,
  ROUND(
    AVG(
      CASE
        WHEN u.is_vip = true THEN o.total_amount * 0.98
        ELSE o.total_amount
      END
    )::numeric,
    2
  ) AS avg_check
FROM orders o
JOIN users u ON u.user_id = o.user_id
WHERE o.status = 'delivered'
  AND u.city = 'Moscow'
  AND o.created_at >= '2025-07-01'
  AND o.created_at <  '2025-08-01';
