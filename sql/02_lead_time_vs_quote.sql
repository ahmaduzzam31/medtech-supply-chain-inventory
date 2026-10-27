-- Quoted vs actual lead time by category: where do planners under-estimate lead times?
SELECT
    i.category,
    ROUND(AVG(i.quoted_lead_time_days), 0)                                  AS quoted_days,
    ROUND(AVG(DATE_DIFF('day', po.order_date, po.received_date)), 0)        AS actual_days,
    ROUND(AVG(DATE_DIFF('day', po.order_date, po.received_date)) - AVG(i.quoted_lead_time_days), 1) AS gap_days,
    ROUND(STDDEV(DATE_DIFF('day', po.order_date, po.received_date)), 1)     AS actual_std_days
FROM purchase_orders po
JOIN items i USING (sku)
WHERE po.status = 'Received'
GROUP BY i.category
ORDER BY gap_days DESC;
