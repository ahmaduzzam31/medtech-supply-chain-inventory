-- Supplier OTIF scorecard: on-time, in-full and quality per supplier (received POs only).
SELECT
    s.supplier_name,
    s.country,
    COUNT(*)                                                                    AS pos_received,
    ROUND(100.0 * AVG(CASE WHEN po.received_date <= po.promised_date THEN 1 ELSE 0 END), 1) AS on_time_pct,
    ROUND(100.0 * AVG(CASE WHEN po.qty_received >= po.qty_ordered THEN 1 ELSE 0 END), 1)    AS in_full_pct,
    ROUND(100.0 * AVG(CASE WHEN po.received_date <= po.promised_date
                            AND po.qty_received >= po.qty_ordered THEN 1 ELSE 0 END), 1) AS otif_pct,
    ROUND(100.0 * SUM(po.qty_rejected) / SUM(po.qty_received), 2)               AS reject_rate_pct,
    ROUND(AVG(DATE_DIFF('day', po.order_date, po.received_date)), 0)            AS avg_lead_time_days
FROM purchase_orders po
JOIN suppliers s USING (supplier_id)
WHERE po.status = 'Received'
GROUP BY s.supplier_name, s.country
ORDER BY otif_pct;
