# Part 1 — Data Cleaning Capstone

## Problem summary
A retail company exported order-level sales data from multiple internal systems. The raw export had inconsistent text formatting, mixed date formats, duplicate records, missing values, invalid discounts, sales/profit calculation mismatches, and order-status inconsistencies. The goal is a clean, validated, analysis-ready dataset plus quality and pivot summary reports for business review.

## Dataset description
- **Source:** `data/raw_orders.xlsx`, sheet `raw_orders`
- **Size:** 932 rows × 21 columns
- **Grain:** one row per order line (`order_id`)
- **Key fields:** order/ship dates, customer + segment, region/state/city, category/sub-category, ship_mode, quantity, unit_price, discount, sales, cost, profit, payment_status, order_status

## Tools used
- Microsoft Excel (`.xlsx` workbook with live formulas: `TRIM`, `PROPER`, `SUBSTITUTE`, `IF`, `MONTH`, `YEAR`, `IFERROR`, `SUMIFS`, `COUNTIFS`)

## Repository structure
```
part1_data_cleaning/
├── data/
│   ├── raw_orders.xlsx         # untouched
│   └── cleaned_orders.xlsx     # cleaned + formulas + Flagged_Duplicates sheet
├── outputs/
│   ├── data_quality_report.xlsx
│   ├── pivot_summary.xlsx
│   └── cleaning_log.md
├── screenshots/
│   ├── raw_data_preview.png
│   ├── cleaned_data_preview.png
│   ├── pivot_summary_1.png
│   ├── pivot_summary_2.png
|   ├── pivot_summary_3.png
|   ├── pivot_summary_4.png
|   └── pivot_summary_5.png
└── README.md
```

## Cleaning steps performed
1. Trimmed leading/trailing and collapsed internal whitespace in all text fields.
2. Normalised case to Title Case for `customer_name`, `segment`, `region`, `state`, `city`, `category`, `sub_category`, `ship_mode`, `payment_status`, `order_status`.
3. Parsed `order_date` and `ship_date` from four input formats (`dd MMM yyyy`, `MM/dd/yyyy`, `dd-MM-yyyy`, `yyyy-MM-dd`) to ISO `yyyy-mm-dd`.
4. Detected and removed 20 exact duplicate rows; moved 12 conflicting `order_id` duplicate rows to a `Flagged_Duplicates` sheet.
5. Filled missing `region`, `ship_mode` with `"Unknown"`; missing `discount` with `0`. All filled rows flagged.
6. Coerced string discounts (`"70%"`, `"85%"`) to numeric.
7. Computed `cleaned_discount`, `calculated_sales`, `calculated_profit`, `profit_margin`, `shipping_delay_days`, `order_month`, `order_year`, and `data_quality_flag` as Excel formulas in the cleaned sheet.
8. Compared raw `sales` against `quantity * unit_price * (1 - cleaned_discount)`; flagged rows whose absolute difference exceeded 1.
9. Flagged negative discounts, discounts above 0.50, and rows where `ship_date < order_date` as invalid.

## Business rules applied
| Rule | Treatment |
|---|---|
| Missing region / ship_mode | Fill `Unknown`, flag `warning` |
| Missing discount | Fill `0`, flag `warning` |
| Negative discount | Keep, flag `invalid` |
| Discount > 0.50 | Keep, flag `invalid` |
| ship_date < order_date | Keep, flag `invalid` |
| sales mismatch (diff > 1) | Use recomputed value, flag `invalid` |
| Cancelled / Failed payments | Excluded from completed-sales pivots |
| Refunded orders | Excluded from completed-sales; reported separately |

## Summary of data quality issues found

| Bucket | Count |
|---|---:|
| Raw rows | 932 |
| Exact duplicates removed | 20 |
| Conflicting order_id duplicates flagged | 12 |
| Final cleaned rows | **900** |
| Clean | 782 |
| Warning | 47 |
| Invalid | 71 |
| Missing region (filled) | 25 |
| Missing ship_mode (filled) | 20 |
| Missing discount (set to 0) | 18 |
| Negative discount | 15 |
| Discount > 0.50 | 15 |
| ship_date < order_date | 20 |
| Sales calculation mismatch | 47 |

Full breakdown in `outputs/data_quality_report.xlsx`.

## Summary of final pivot reports
`outputs/pivot_summary.xlsx` contains:
1. **Sales Profit by Region** — completed sales & profit per region (sorted by sales desc).
2. **Sales Profit by Category** — completed sales & profit by category + sub-category (sorted).
3. **Orders by Ship_Mode** — completed order count by ship mode (sorted).
4. **Margin by Segment** — profit margin by customer segment.
5. **NonCompleted by Region** — refunded / cancelled / failed orders per region.
6. **Monthly Sales Trend** — completed sales month-by-month.

"Completed" = `order_status = Completed` AND `payment_status = Paid`. Refunded, cancelled, and failed flows are excluded from these totals and summarised separately. There are 601 completed-paid rows out of 900.

## Key business insights
- **Failed/cancelled/refunded together represent ~31% of all rows** (282 of 900). That's a meaningful signal for "revenue-at-risk": cleaning revealed many cancellations that were silently mixed into the raw sales total.
- **Negative discounts (15 rows) and discounts above 0.50 (15 rows)** look like data-entry or system-export errors — they distort both raw sales and the computed margin. They've been flagged, not silently corrected.
- **47 rows have raw `sales` that don't agree with `quantity * unit_price * (1 - discount)`** by more than ₹1 — a strong indicator that the upstream source is not internally consistent. Downstream analytics should use the recomputed `calculated_sales` rather than raw `sales`.
- Date formats coming from upstream systems are not standardised — at least 4 distinct formats in a single column. The cleaned file enforces one ISO format.

## Assumptions and limitations
- **Discount cap = 0.50** is my choice (the brief said "above allowed range" without a number).
- For conflicting `order_id` duplicates, the first occurrence is kept; the brief doesn't define which row is canonical.
- `product_name` is only trimmed (not Title-Cased) because it contains embedded model codes.

## Screenshots included
| File | Shows |
|---|---|
| `screenshots/raw_data_preview.png` | First ~25 rows of the raw dataset before cleaning |
| `screenshots/cleaned_data_preview.png` | Cleaned dataset including the calculated columns and `data_quality_flag` |
| `screenshots/pivot_summary_1.png` | `Sales Profit by Region` sheet |
| `screenshots/pivot_summary_2.png` | `Sales Profit by Category` sheet |
| `screenshots/pivot_summary_3.png` | `Orders by Ship_Mode` sheet |
| `screesnhots/pivot_summary_4.png` | `Margin by Segment` sheet |
| `screenshots/pivot_summary_5.png` | `Monthly Sales Trend` sheet |
