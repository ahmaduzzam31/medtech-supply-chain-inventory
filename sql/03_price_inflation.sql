-- Purchase price index by category and quarter (standard cost = 100).
SELECT
    i.category,
    STRFTIME(DATE_TRUNC('quarter', po.order_date), '%Y-Q') || QUARTER(po.order_date) AS quarter,
    ROUND(100.0 * SUM(po.qty_ordered * po.unit_price_eur) / SUM(po.qty_ordered * i.unit_cost_eur), 1) AS price_index
FROM purchase_orders po
JOIN items i USING (sku)
GROUP BY i.category, DATE_TRUNC('quarter', po.order_date), quarter
ORDER BY i.category, DATE_TRUNC('quarter', po.order_date);
