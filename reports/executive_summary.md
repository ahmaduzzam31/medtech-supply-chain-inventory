# Executive Summary: Component Inventory & Supplier Review

**To:** Operations Director, Head of Supply Chain, Finance Controller
**Company:** Corrib Medical Devices (fictional; simulated data)
**Scope:** 116 components, 12 suppliers, 2,024 purchase orders; stock snapshot 30-Sep-2026

## Situation
Operations reports *"too much stock, but we still run out of chips"*, and purchasing is **€3.2M (9.4%) over budget**.

## Key findings
1. **The money is in the wrong parts.** Inventory is €5.03M, of which **€1.06M (21%) is excess**, mostly cheap passive
   components and sensors. Carrying it costs about **€234k a year**.
2. **Production is at risk.** **9 parts are below safety stock, and all are chips or batteries.** Several A-class ICs have 1 to 4
   weeks of cover against 3 to 5 month lead times.
3. **Two suppliers drive most of the risk.** Formosa Microtech (MCUs) has **59% OTIF**. Shenzhen Brightcell has **64% OTIF**,
   the highest reject rate (1.8%), and is the **single source for all six batteries** (€3.0M a year).
4. **The overspend is price, not volume.** €2.5M of the €3.2M variance is price, of which **€1.9M is an 18% increase on
   microcontrollers** since July 2025. Volume and timing explain only €0.7M.
5. **25% of parts carry 80% of spend.** 29 A-class components deserve weekly control. Half of the SKUs (C-class) are 5% of spend.

## Recommendations
| # | Action | Impact | Owner |
|---|---|---|---|
| 1 | Expedite the 9 parts below safety stock and agree allocation with suppliers | Protects all three product lines | Planning / Purchasing |
| 2 | Move C-class parts to min/max using the model's reorder points and let excess run down | About €1M cash released, €234k a year carrying cost | Planning |
| 3 | Qualify a second battery supplier, including design verification and regulatory change control | Removes a single point of failure on €3.0M of spend | Engineering, Quality, Purchasing |
| 4 | Supplier improvement plans for Formosa and Shenzhen Brightcell, with monthly OTIF reviews | OTIF 59% / 64% vs 85% minimum | Supplier Quality |
| 5 | Long-term agreements on A-class ICs; reset standard costs for next year's budget | Addresses the €1.9M chip price variance | Purchasing / Finance |
| 6 | Plan with actual lead times (from PO history), not supplier quotes | More reliable reorder points | Planning |

*All data is simulated for a fictional company to demonstrate the method.*
