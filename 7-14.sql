SELECT
    COUNT(*) AS orders_total,
    COUNT(CASE
        WHEN EXTRACT(EPOCH FROM (delivered_at::timestamp - picked_at::timestamp)) / 60 > 60
        THEN 1
    END)::float / COUNT(*) AS share_deliveries_over_60_min
FROM deliveries
WHERE delivered_at IS NOT NULL
  AND delivered_at::timestamp >= '2025-07-01'
  AND delivered_at::timestamp < '2025-08-01';
