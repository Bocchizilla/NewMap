SELECT
    COUNT(*) AS users_with_many_promos,
    COUNT(*)::float
        / NULLIF((
            SELECT COUNT(DISTINCT user_id)
            FROM orders
            WHERE created_at::date >= CURRENT_DATE - 30
          ), 0) AS share_users_with_many_promos
FROM (
    SELECT
        user_id
    FROM orders
    WHERE created_at::date >= CURRENT_DATE - 30
      AND promo_code IS NOT NULL
    GROUP BY user_id
    HAVING COUNT(DISTINCT promo_code) > 1
) t;
