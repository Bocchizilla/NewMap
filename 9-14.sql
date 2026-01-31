SELECT
    COUNT(DISTINCT june.user_id) AS reactivated_users,
    COUNT(DISTINCT june.user_id)::float
        / COUNT(DISTINCT all_june.user_id) AS share_reactivated_users
FROM orders june
LEFT JOIN orders may
    ON june.user_id = may.user_id
   AND may.created_at::date >= '2025-05-01'
   AND may.created_at::date < '2025-06-01'
JOIN orders all_june
    ON june.user_id = all_june.user_id
WHERE june.created_at::date >= '2025-06-01'
  AND june.created_at::date < '2025-07-01'
  AND may.order_id IS NULL;
