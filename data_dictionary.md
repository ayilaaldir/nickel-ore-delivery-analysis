# Data dictionary

**SYNTHETIC EDUCATIONAL DATA. Author: Ridla A. Gani.** Fixed seed 20260901. Ground truth contains 300 original deliveries, five deliveries per day for 60 days, 2026-09-01 through 2026-10-30. Raw contains 310 immutable rows. Dates in CSV are ISO dates; workbook dates are numeric Excel serials formatted dd-mmm-yyyy. Text IDs retain leading zeros.

| Field | Type / unit | Meaning and rule |
|---|---|---|
| source_row_id | unique text | R000001 etc.; stable raw-row identity; correction key; never reused |
| delivery_id | text | Business delivery identifier; trim and upper-case for analysis; conflicts block admission |
| delivery_date | date | Delivery day; missing dates remain blank; correction requires reason |
| supplier_id | text | Normalized exact lookup to four fictional active suppliers |
| stockpile_id | text | Normalized exact lookup to three fictional stockpiles |
| supplier_reported_wmt | decimal WMT | Supplier statement; must be positive; variance denominator |
| weighbridge_net_wmt | decimal WMT | Measured wet mass; must be positive; accepted metrics use this basis |
| moisture_pct | decimal fraction | Water mass fraction on wet basis; 0.32 displays 32.00%; mathematical bounds [0,1] |
| ni_grade_pct_dry | decimal fraction | Dry-basis Ni fraction; 0.015 displays 1.50%; mathematical bounds [0,1] |
| record_note | text | Synthetic source context; preserved, including duplicate payload notes |
| supplier_name / stockpile_name | text | Fictional reference labels; production VLOOKUP uses exact match |
| active_status | text | Active status required for supplier acceptance |

## Review and corrections

DATA REVIEW A:AK is calculated apart from the stable source_row_id key. AB:AJ shows the original source payload. The separate native ReviewDecisions Table in AQ6:BA46 contains source_row_id, decision, reason, corrected_date, corrected_supplier_id, corrected_stockpile_id, corrected_reported_wmt, corrected_weighbridge_wmt, corrected_moisture_fraction, corrected_grade_fraction and reviewer_notes. Original delivery_id cannot be manually renamed to bypass a conflict.

Blank overrides mean retain the source value. Corrections are not zero-filled. Numeric zero remains an explicit override and is subject to mathematical gates. A missing original is shown as an empty string by formulas, with an explanatory validation status. Do not infer zero from an empty cell. Blank override fields cannot deliberately clear a valid original; use Hold to block it pending review instead.

Inclusion precedence is mutually exclusive: (1) exact later retransmission = Excluded duplicate; (2) invalid correction control = Unresolved; (3) documented Reject = Rejected; (4) mathematical invalidity, unresolved conflicting ID, Hold, or unapproved illustrative review flag = Unresolved; (5) otherwise Accepted. Only a valid documented Reject removes a competitor. Exact-match signatures use all nine original payload fields (including record_note), with numbers formatted to six decimals, matching the fixed dataset precision. The smallest zero-padded source ID is the retained candidate, independent of sort order.

## Accepted calculations

| Metric | Formula | Unit / denominator |
|---|---|---|
| WMT | weighbridge_net_wmt | wet metric tonnes, accepted records |
| DMT | WMT * (1 - moisture_fraction) | dry metric tonnes |
| theoretical_contained_Ni_t | DMT * grade_dry_fraction | nickel tonnes, theoretical content |
| weighted_grade_dry | SUM(theoretical Ni) / SUM(DMT) | dry-basis fraction; unavailable when DMT = 0 |
| weighted_moisture | 1 - SUM(DMT) / SUM(WMT) | wet-basis water fraction; unavailable when WMT = 0 |
| signed_variance_WMT | weighbridge - supplier reported | signed tonnes |
| variance_fraction | signed variance / supplier reported | requires positive numeric denominator and valid measured weight |
| weight_exception | ABS(variance_fraction) > threshold | review flag, preserves signed variance separately |
| accepted_flag | IF(status = Accepted, 1, 0) | count denominator |
| selected_scope_flag | accepted AND date bounds AND supplier AND stockpile | dashboard/scenario population; 0 for invalid scope |
| accepted_ID_count | COUNTIFS(accepted delivery_id) | must be exactly 1 for every accepted record |

Theoretical content does not estimate recovery, payable nickel, revenue or profit. No simple average of moisture or grade is used. WMT/DMT/Ni retain full internal precision; displayed rounding never feeds calculation. Both fractions may be zero or one mathematically, while business-review ceilings remain separate.

## Scenarios

Volume input is a relative fraction: 10% = 0.10. Moisture and grade inputs are numeric percentage points: 2 pp = 2 / 100 = 0.02; 0.10 pp = 0.10 / 100 = 0.001.

scenario_WMT = baseline_WMT * (1 + volume_change)

scenario_moisture = baseline_moisture + moisture_pp_change / 100

scenario_grade = baseline_grade + grade_pp_change / 100

scenario_DMT = scenario_WMT * (1 - scenario_moisture)

scenario_Ni = scenario_DMT * scenario_grade

Apply to each scoped accepted delivery, then aggregate. Inputs must be numeric, volume_change >= -1 and every resulting moisture/grade fraction in [0,1]. There is no silent clipping. If any scoped record is invalid, the entire scenario summary is Unavailable with BLOCKED status. An empty accepted scope reports NO ACCEPTED RECORDS IN SCOPE. One-way moisture sensitivity uses DMT_new = DMT_base - WMT_base * pp/100 and Ni_new = Ni_base - pp/100 * SUM(WMT * grade), with delivery bounds checked. Grade sensitivity uses Ni_new = Ni_base + pp/100 * DMT_base and unchanged WMT/DMT. These equal delivery-level recomputation for uniform additive changes.

## References, exports and evaluation separation

REFERENCES B20:B22 stores illustrative review thresholds: absolute variance 0.03, moisture ceiling 0.40 and dry-grade ceiling 0.03. Units are fractions, displayed as percentages. These are not industry standards. Threshold cells affect review flags and inclusion immediately on recalculation; reasons and Approve remain required when flagged.

exports/accepted_deliveries.csv and exports/summary.csv include a synthetic_notice on every row and are initial snapshot exports. summary.csv grain is dimension + group: total, supplier, stockpile or date. Do not add different dimension rows together; each dimension represents the same accepted population. Ratios use that row's own tonnage denominators.

evaluation/clean_ground_truth.csv is evaluation-only uncorrupted reference data. It is not the accepted dataset: legitimate weight discrepancies can remain in truth, while unresolved raw records must be excluded from production analysis. evaluation/issue_manifest.csv records 64 injected issue entries with expected initial handling. Production workbook formulas and builder input files never read either evaluation file. The verifier reads them only to audit injection coverage and handling.
