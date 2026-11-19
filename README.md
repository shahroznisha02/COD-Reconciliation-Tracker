# COD Reconciliation Tracker (Excel)

**Caught 7 cash mismatches worth ₹6,897 across 36 COD orders, automatically, using Excel.**

![What the tracker caught](dashboard.png)

A self-project I built in November 2026 to understand **Cash on Delivery (COD) reconciliation** in quick-commerce logistics. Riders collect cash from customers. This tracker checks, order by order, whether the cash in the rider report matches what our system expected, and turns the result into a daily MIS report.

> All orders, riders and amounts are **sample data** made up for practice.

## What it does
- Matches **36 orders** from internal records with the **rider/partner collection report**
- Flags every discrepancy: **short collected, excess collected, missing in rider report, cash shown on RTO/cancelled orders**, and rider entries **not in our system**
- Builds a **daily MIS summary**: totals, match rate, status-wise, rider-wise and date-wise reports
- Highlights problem rows in red so they can be followed up with operations and finance teams

## Sheets
| Sheet | Purpose |
|---|---|
| How to Use | Explanation of the process and every key formula |
| MIS_Summary | Daily MIS report for the manager |
| Orders | Internal records (order, date, rider, COD amount, delivery status) |
| Rider_Collections | Partner/rider report of cash collected |
| Reconciliation | Expected vs collected, difference and recon status per order |

## Key Excel formulas
- **Expected COD:** `=IF(Status="Delivered", Amount, 0)`
- **Collected:** `=IFERROR(VLOOKUP(OrderID, Rider_Collections!A:D, 4, FALSE), 0)`
- **Difference:** `=Expected - Collected` (positive = short, negative = excess)
- **Recon status:** nested `IF` + `COUNTIF`
- **MIS summary:** `COUNTIF`, `COUNTIFS`, `SUMIFS`
- Conditional formatting to highlight discrepancies

## Result on the sample data
- 26 of 31 COD orders matched (83.9% match rate)
- 7 issues caught: 2 short collections, 1 excess, 2 missing from the rider report, 1 cash on an RTO order, 1 rider entry not in the system

## Skills used
MS Excel, VLOOKUP, IF, COUNTIF, SUMIFS, conditional formatting, MIS reporting, reconciliation

**Author:** Shahroz Nisha · [LinkedIn](https://www.linkedin.com/in/shahroz-nisha-093957432) · [Portfolio](https://shahroznisha02.github.io)
