# Cleaning Log

## 1. Source dataset
- File: `data/raw_orders.xlsx`, sheet `raw_orders`
- Rows: **932**, columns: **21**
- Untouched in place; cleaned output written to `data/cleaned_orders.xlsx`

## 2. Issues found in the raw data

| Issue | Count | Examples |
|---|---:|---|
| Mixed-case + whitespace in text fields | many | `  North `, `OFFICE SUPPLIES`, `Small  Business`, `STANDARD CLASS`, `completed`, `Paid ` |
| Mixed date formats | all rows mix-and-match | `21 Jul 2024`, `07/27/2024`, `28-11-2024`, `2024-05-24` |
| Missing `region` | 26 | filled as `Unknown` |
| Missing `ship_mode` | 22 | filled as `Unknown` |
| Missing `discount` | 18 | filled as 0 |
| Discount stored as percent string | 8 | `'70%'`, `'85%'` — parsed to 0.70 / 0.85 |
| Negative discount | 15 | flagged invalid |
| Discount above the 0.50 cap | 15 | flagged invalid |
| ship_date earlier than order_date | 20 | flagged invalid |
| Exact duplicate rows | 20 | dropped (first kept) |
| Conflicting `order_id` duplicates | 12 ids / 12 extra rows | first kept; extras to `Flagged_Duplicates` |
| Raw `sales` ≠ `quantity * unit_price * (1 - discount)` | 47 | flagged invalid; cleaned columns use recomputed values |

## 3. Cleaning actions performed

### Text fields
Applied to `customer_name`, `segment`, `region`, `state`, `city`, `category`, `sub_category`, `ship_mode`, `payment_status`, `order_status`:

- `TRIM` to remove leading/trailing whitespace **and** collapse internal multi-spaces
- `PROPER`-style Title Case for consistent display
- `SUBSTITUTE` (CHAR(160) → space) as a safety net for non-breaking spaces

These transformations are pre-applied in the cleaned sheet for clarity. The equivalent Excel formula is documented in the workbook's `Cleaning_Notes` sheet, e.g. `=PROPER(TRIM(Raw_Data!G2))`.

### Dates
- Input formats encountered: `dd MMM yyyy`, `MM/dd/yyyy`, `dd-MM-yyyy`, `yyyy-MM-dd`.
- Dates were re-written as **real Excel date values** (`yyyy-mm-dd` formatted).
- `shipping_delay_days = ship_date - order_date` computed via Excel formula in the cleaned sheet.

### Duplicates
- Exact duplicate rows (all 21 columns identical) → dropped, count reported.
- Conflicting `order_id` duplicates (same id, different field values) → **first occurrence retained** in the cleaned sheet; remaining copies moved to a `Flagged_Duplicates` sheet inside `cleaned_orders.xlsx` with a `reason` column.

### Numeric / business-rule fields
- `cleaned_discount = IF(ISBLANK(discount), 0, discount)` (Excel formula in cleaned sheet)
- `calculated_sales = quantity * unit_price * (1 - cleaned_discount)`
- `calculated_profit = calculated_sales - cost`
- `profit_margin = IFERROR(calculated_profit / calculated_sales, 0)`
- `order_month = MONTH(order_date)`, `order_year = YEAR(order_date)`
- `data_quality_flag ∈ {clean, warning, invalid}` set during build based on the rules below; conditional formatting highlights flag values in the cleaned sheet.

## 4. Business rules applied

| Rule | Action | Severity |
|---|---|---|
| Missing `region` | fill `Unknown` | warning |
| Missing `ship_mode` | fill `Unknown` | warning |
| Missing `discount` (other sales fields valid) | fill `0` | warning |
| Negative `discount` | keep, flag | **invalid** |
| `discount > 0.50` (chosen as "above allowed range") | keep, flag | **invalid** |
| `ship_date < order_date` | keep, flag | **invalid** |
| `|raw_sales - recomputed_sales| > 1` | use recomputed, flag | **invalid** |
| `order_status = Cancelled` | exclude from completed-sales pivots | — |
| `payment_status = Failed` | exclude from completed-sales pivots | — |
| `payment_status = Refunded` | excluded from completed-sales pivots, summarised separately | — |

A row's `data_quality_flag` is `invalid` if any hard-fail rule triggers; otherwise `warning` if any soft rule triggers; otherwise `clean`.

## 5. Records removed / flagged

- Rows removed (exact duplicates): **20**
- Rows moved to Flagged_Duplicates sheet: **12**
- Final cleaned rows: **900**
- Of those: **782 clean**, **47 warning**, **71 invalid**

## 6. Assumptions

- The "discount above allowed range" threshold is not specified in the brief. **Chosen as 0.50** — gives a non-empty invalid-discount category in the report (15 rows).
- For conflicting `order_id` duplicates, "first occurrence" means the first row position in the source file (no requirement specified in the brief favours one over the other).
- `product_name` is left case-as-is after `TRIM`, since product names embed model codes that should not be Title-Cased.
- Cancelled orders that paid (`Paid` + `Cancelled`) are still excluded from completed-sales pivots — the brief treats `Cancelled` as the disqualifying state regardless of payment.
- Refunded orders are counted in the refund summary regardless of their `order_status`.
