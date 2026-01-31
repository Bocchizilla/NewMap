WITH payments_ranked AS (
    SELECT
        payment_id,
        order_id,
        amount,
        method,
        paid_at,
        ROW_NUMBER() OVER (
            PARTITION BY order_id, amount, method
            ORDER BY paid_at DESC
        ) AS rn
    FROM payments
)
SELECT
    COUNT(*)        AS duplicate_payments_cnt,
    SUM(amount)     AS duplicate_extra_amount
FROM payments_ranked
WHERE rn > 1;
