# 2-3 minute portfolio walkthrough

**Ridla A. Gani | SYNTHETIC EDUCATIONAL PORTFOLIO**

**0:00-0:20 - Context.** Open START HERE. Say: "I built this independent Excel portfolio using fictional records. It demonstrates delivery reconciliation and assumption analysis, not actual mine experience. The workbook uses no macros or external links."

**0:20-0:45 - One source issue and its review.** In RAW DELIVERIES filter source_row_id to R000041: moisture is blank and the fictional source note records a 32% retest. DATA REVIEW original_moisture_fraction is blank, while effective moisture is 32%. Go to AQ6: the keyed R000041 entry has Approve, a reason and corrected_moisture_fraction 32%. Explain that production does not read evaluation ground truth. Show R000042, whose absent moisture remains Unresolved and whose clean DMT/Ni stay blank.

**0:45-1:10 - Inclusion control.** In an editable copy, set R000041's keyed decision to Hold. Recalculate: the correction requires Approve and the accepted count drops from 262 to 261. Restore Approve. Show the reason requirement and R000307's documented Reject resolving R000081's competing submission. Original rows remain present. Do not edit the calculated review fields.

**1:10-1:35 - Filtered operations.** On DASHBOARD select supplier S02 and stockpile P01, leaving both dates unchanged. Show the selected accepted count/WMT and daily trend. Point to the separately labelled global counts: these do not follow selectors. Restore All/All. The all-scope verified baseline is 8,542.047 WMT, 5,939.540 DMT and 92.678 t theoretical Ni.

**1:35-2:05 - Numerical scenario.** On SCENARIOS explain B6 = 10% relative volume, E6 = 2 moisture percentage points and H6 = 0.10 dry-grade percentage points. For the all-scope case, theoretical Ni becomes 105.359 t. Use the visible 100 WMT example: 30% becomes 32%, 1.50% becomes 1.60%, yielding 74.8 DMT and 1.1968 t Ni after +10% WMT. These are assumptions, not forecasts or recoverable metal. Enter 100 moisture pp to show BLOCKED; restore 2.

**2:05-2:30 - Reconciliation.** Show QA CHECKS: 310 = 262 accepted + 6 excluded duplicate + 41 unresolved + 1 rejected, and supplier WMT reconciliation = zero. SUMMARY contains SUMIFS formula tables, explicitly not PivotTables. Point to the exact Excel steps for genuine PivotTables and refresh.

**2:30-2:50 - Limits.** Say: "I verified the saved workbook and mutation cases with artifact-tool and independent Python. Excel automation was unavailable, so I do not claim Excel UI or native Pivot refresh testing. This fixed synthetic sample supports a transparent portfolio demonstration, not real savings, mine approval or deployment." Finish with the two-page case study and test report.

Before a live Excel demonstration, validate dropdowns, chart redraw, recalculation and protection behavior in Excel 365. Restore default decisions, dates, All/All and 10% / 2 pp / 0.10 pp after the demo. CSV and case-study snapshots do not update with live edits.
