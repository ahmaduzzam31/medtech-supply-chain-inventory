-- Supply risk: spend that depends on a single approved supplier, with that supplier's OTIF.
WITH perf AS (
    SELECT supplier_id,
           AVG(CASE WHEN received_date <= promised_date AND qty_received >= qty_ordered THEN 1.0 ELSE 0 END) AS otif
    FROM purchase_orders WHERE status = 'Received' GROUP BY supplier_id
),
spend AS (
    SELECT sku, SUM(qty_ordered * unit_price_eur) AS spend_eur
    FROM purchase_orders WHERE order_date >= DATE '2025-10-01' GROUP BY sku
)
SELECT
    s.supplier_name,
    i.category,
    COUNT(*)                         AS single_source_skus,
    ROUND(SUM(sp.spend_eur), 0)      AS fy_spend_eur,
    ROUND(100 * ANY_VALUE(p.otif), 1) AS supplier_otif_pct
FROM items i
JOIN suppliers s USING (supplier_id)
JOIN perf p      USING (supplier_id)
LEFT JOIN spend sp USING (sku)
WHERE i.single_source
GROUP BY s.supplier_name, i.category
ORDER BY supplier_otif_pct;
