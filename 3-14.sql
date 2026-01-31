SELECT
  c.channel,
  COUNT(*) AS users_registered,
  COUNT(o.user_id) AS users_bought_7d,
  ROUND(
    100.0 * COUNT(o.user_id)::numeric / NULLIF(COUNT(*)::numeric, 0),
    2
  ) AS conv_pct
FROM (
  SELECT
    u.user_id,
    u.created_at,
    COALESCE(fs.channel, 'unknown') AS channel
  FROM users u
  LEFT JOIN (
    SELECT user_id, channel
    FROM (
      SELECT
        user_id,
        channel,
        ROW_NUMBER() OVER (PARTITION BY user_id ORDER BY started_at) AS rn
      FROM sessions
    ) s
    WHERE rn = 1
  ) fs ON fs.user_id = u.user_id
  WHERE u.created_at::timestamp >= TIMESTAMP '2025-06-01'
    AND u.created_at::timestamp <  TIMESTAMP '2025-07-01'
) c
LEFT JOIN orders o
  ON o.user_id = c.user_id
 AND o.status = 'delivered'
 AND o.created_at::timestamp <= c.created_at::timestamp + INTERVAL '7 days'
GROUP BY c.channel
ORDER BY conv_pct DESC;
