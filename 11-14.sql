SELECT
    AVG(CASE WHEN ab.variant = 'A' THEN o.total_amount END) AS avg_a,
    AVG(CASE WHEN ab.variant = 'B' THEN o.total_amount END) AS avg_b,
    (
        AVG(CASE WHEN ab.variant = 'B' THEN o.total_amount END)
        - AVG(CASE WHEN ab.variant = 'A' THEN o.total_amount END)
    ) * 100.0
    / NULLIF(
        AVG(CASE WHEN ab.variant = 'A' THEN o.total_amount END),
        0
    ) AS uplift_percent
FROM orders o
JOIN ab_tests ab
    ON o.user_id = ab.user_id
WHERE o.status = 'delivered'
  AND o.created_at::date >= '2025-07-01'
  AND o.created_at::date < '2025-08-01';
