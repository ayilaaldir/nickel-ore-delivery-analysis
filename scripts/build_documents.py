"""Write the handoff documents and a two-page management case study from executed results."""
import json
import platform
from collections import Counter
from pathlib import Path
from project_paths import local
from analysis_model import analyze, aggregate

def save(name,text):local(name).write_text(text.strip()+'\n',encoding='utf-8')

def main():
    r=json.loads(local('reports/verification_results.json').read_text(encoding='utf-8'))
    t=r['totals'];s=r['scenario'];counts=r['counts'];tests=r['tests'];fails=sum(x['result']=='FAIL' for x in tests)
    change=(s['contained_ni_t']/t['contained_ni_t']-1)*100
    save('README.md',f'''
# NickelOps — Ore Delivery Reconciliation and Scenario Analysis

**Ridla A. Gani | Independent portfolio project | SYNTHETIC EDUCATIONAL DATA**

I built NickelOps to explore an operational question: which delivery records can an analyst include, how do supplier and weighbridge weights differ, and how do volume, moisture and nickel-grade assumptions change the result? All suppliers, stockpiles and records are fictional. This is independent portfolio work, not actual mine operations, paid client work, an engineering-approved model or professional mining experience.

The main deliverable is a macro-free Excel 365 workbook with nine sheets, native Excel Tables, exact-match VLOOKUPs, SUMIFS summaries, IF validation gates, conditional formatting and two native charts. **The installed summaries are formula summaries, not PivotTables.** Manual PivotTable steps are below.

## Download and use

[Download NickelOps.xlsx](https://raw.githubusercontent.com/ayilaaldir/nickel-ore-delivery-analysis/main/NickelOps.xlsx) · [Two-page case study](case_study.pdf) · [Complete project release](https://github.com/ayilaaldir/nickel-ore-delivery-analysis/releases/latest)

Normal workbook use is offline. It needs Excel 365, without Python, macros, external workbook links or a cloud account.

1. Open **START HERE**, then **DASHBOARD**. The default scope is 1 September–30 October 2026, all suppliers and all stockpiles. Dashboard **B6/D6/F6/H6** control start date, end date, supplier and stockpile. Selected KPIs and the trend chart share that accepted-record scope; global data-quality counts are labelled separately.
2. Inspect **RAW DELIVERIES** and **DATA REVIEW**. The source snapshot is retained. DATA REVIEW shows original values in AB:AJ, effective values in C:J and inclusion status in B. Do not edit the calculated table.
3. Use the yellow **DATA REVIEW!AQ6:BA46** register to enter a stable `source_row_id`, decision and reason. Corrections require **Approve** and a reason. Approve/Reject require reasons; Hold blocks inclusion. Sort the whole ReviewDecisions Table, never one column. Blank overrides retain the source; zero is an explicit value subject to validation. Enter moisture/grade as fractions, such as **32%** and **1.5%**, not 32 and 1.5.
4. In **SCENARIOS**, B6 is a relative volume change (**10%**), E6 is moisture percentage points (**2**), and H6 is dry-grade percentage points (**0.10**). These add 0.02 and 0.001 to the respective delivery fractions. Invalid assumptions block the scenario with an explanatory status.
5. Check **QA CHECKS** and the [executed test report](test_report.md). Use an editable copy for review/scenario experiments. Worksheet filters change visible rows; dashboard controls determine the reporting scope.

![Dashboard showing synthetic accepted deliveries](previews/dashboard.png)

*Preview rendered with artifact-tool. It is not evidence of native Excel UI testing.*

## Dataset and reconciliation

Seed **20260901** generates 300 underlying deliveries over 60 days, four fictional suppliers and three fictional stockpiles. A separate raw snapshot has 310 rows with deliberate identifier noise, unknown suppliers, missing values, invalid fractions, negative weights, exact retransmissions, conflicting delivery IDs and material weight differences. An independent evaluation manifest contains 64 injected-issue entries. Production formulas do not read the manifest or clean ground truth.

Every raw row has a stable source ID. Exact retransmissions retain the smallest source ID; later copies stay visible as excluded duplicates. Conflicting delivery IDs remain blocked until competing submissions receive documented rejections. Mathematical validity always applies, even after approval. Four initial review examples are documented in [data_dictionary.md](data_dictionary.md); other missing measurements are not imputed.

The verified initial population is **310 = 262 accepted + 6 excluded duplicate + 41 unresolved + 1 rejected**. Accepted business IDs are unique.

| Accepted-record metric | Initial snapshot |
|---|---:|
| Weighbridge wet mass | {t['wmt']:,.3f} WMT |
| Dry mass | {t['dmt']:,.3f} DMT |
| Weighted dry-basis nickel grade | {t['weighted_grade']:.4%} |
| Aggregate moisture | {t['weighted_moisture']:.4%} |
| Theoretical contained nickel | {t['contained_ni_t']:,.6f} t |
| Accepted weight-variance exceptions | 1 |

Supplier, stockpile and daily summaries reconcile to the accepted population within **1e-7 tonnes**. The 41 unresolved records remain outside these totals. These are observations from a synthetic review state, not conclusions about supplier performance.

## Metrics and assumptions

The analysis uses weighbridge net WMT. **DMT = WMT × (1 − moisture)** and **theoretical Ni tonnes = DMT × dry-basis grade**. Aggregate grade is total Ni / total DMT; aggregate moisture is 1 − total DMT / total WMT. Percentages are stored as fractions, so 1.5% is 0.015. Percentage fields are never averaged without their appropriate weights. Missing/invalid measurements produce unavailable metrics, not false zeros.

Signed weight variance is weighbridge minus supplier-reported WMT. Percentage variance divides by positive reported WMT. Review flags use its absolute percentage while retaining the sign. The default 3% variance, 40% moisture and 3% dry-grade thresholds are configurable **illustrative assumptions, not industry standards**. They are separate from mathematical fraction bounds [0,1] and positive-weight rules.

The default combined case assumes **+10% relative WMT, +2 moisture pp and +0.10 dry-grade pp**, applied at each scoped accepted delivery before aggregation. It produces **{s['wmt']:,.3f} WMT, {s['dmt']:,.3f} DMT and {s['contained_ni_t']:,.6f} t theoretical Ni**. These are assumptions, not forecasts. One-way sensitivities isolate moisture and grade effects.

Theoretical contained nickel is **not recoverable nickel, payable nickel, revenue or profit**. The project makes no claim of real savings, production gains or deployment.

## Verification and reproduction

**{len(tests)} recorded checks passed; {fails} failed.** Artifact-tool exercised formula mutations and chart-source updates; Python independently recomputed results and openpyxl inspected the saved workbook and cached values. Checks cover row accounting, all manifest entries, missing/invalid inputs, keyed decisions, weighted calculations, summary reconciliation, scenario scaling/directions, percentage-point examples and invalid assumptions. Workbook previews and the two-page PDF were rendered and visually inspected. See [test_report.md](test_report.md) and `reports/` for expected/actual values, tolerances and engines.

**Native Excel UI and Excel-engine recalculation were not tested.** Excel COM activation failed with HRESULT 0x80070520. Native dropdowns, protection interaction, chart redraw and PivotTable creation/refresh remain manual checks. Installed formulas and charts alone do not prove those interactions.

To check the supplied workbook, run `python scripts/verify_calculations.py` from the project root with Python 3.11+ and openpyxl 3.1.5. It does not modify the XLSX; it refreshes CSV exports and verification reports.

A complete rebuild also needs the packages in `requirements.txt`, Node (tested 26.13.0) and an existing **@oai/artifact-tool 2.8.158** installation. Artifact-tool is not bundled or installed by `requirements.txt`; full rebuilding is unavailable without that runtime. Set `NICKELOPS_NODE_MODULES` to its installed node_modules directory when needed. The finished workbook does not depend on that runtime. PDF generation uses ReportLab; validation/rendering uses pypdf and pypdfium2.

Scripts enforce the original folder name. Clone into that name and use project-local temporary/cache directories:

```powershell
git clone https://github.com/ayilaaldir/nickel-ore-delivery-analysis.git NickelOps-Ore-Delivery-Reconciliation-and-Scenario-Analysis
cd NickelOps-Ore-Delivery-Reconciliation-and-Scenario-Analysis
$env:TEMP = Join-Path (Get-Location).ProviderPath 'tmp'
$env:TMP = $env:TEMP
$env:PYTHONPYCACHEPREFIX = Join-Path (Get-Location).ProviderPath 'tmp/cache'
python scripts/generate_data.py
python scripts/build_spec.py
node scripts/build_workbook.mjs
node scripts/test_qa_compatibility.mjs
python scripts/finalize_xlsx.py
node scripts/render_final_dashboard.mjs
python scripts/verify_calculations.py
python scripts/build_documents.py
python scripts/package_project.py
```

Use a project-local environment for any dependency installation. Generation rebuilds the synthetic snapshot and resets the four initial review decisions: save edited work first. Generated `data/workbook_input.json` is an intermediate recreated by generation. CSV exports, case study, PDF and test report are snapshots and do not update with Excel edits. Preserve reviewed inputs in `data/review_decisions.csv` and the corresponding generated workbook input before rebuilding without `generate_data.py`.

## Capacity, refresh and manual Excel checks

The workbook has fixed capacities: **310 source/review/clean/scenario rows, 40 keyed review entries and 60 daily-summary dates**. Appending beyond those ranges is not supported. Regenerate and update capacities/date ranges for a larger dataset. Accepted dates outside the daily-summary horizon trigger a reconciliation failure until the window is extended.

In Excel 365, set Formulas → Calculation Options → Automatic and press F9. Exercise a dashboard selector, keyed decision and scenario input; check the affected outputs/charts, then restore defaults. Native interaction remains unverified until this is done. Initial Python QA controls are delivery-time benchmarks and report REVIEW when the population changes.

To create genuine PivotTables manually:

1. Select **CLEAN DELIVERIES!A6:T316**. Insert → PivotTable → From Table/Range; confirm **CleanDeliveries** and choose Existing Worksheet **SUMMARY!K20**.
2. Filter `inclusion_status` to **Accepted**. Put `supplier_name` in Rows. Put `accepted_flag`, `weighbridge_WMT`, `DMT` and `theoretical_contained_Ni_t` in Values, with every field set to **SUM**. The accepted_flag sum is delivery count.
3. Repeat at **SUMMARY!K35** with `stockpile_name`. For a daily trend use `delivery_date`, ungrouping dates if needed.
4. After input changes, calculate first, then Data → Refresh All. Confirm the Accepted filter. Compare WMT/DMT/Ni grand totals with `SUM(SUMMARY!C7:C10)`, `SUM(SUMMARY!D7:D10)` and `SUM(SUMMARY!E7:E10)`.
5. Outside the Pivot, calculate grade as total Ni / total DMT and moisture as 1 − total DMT / total WMT. Do not use Average of the percentage fields.

No PivotTable is installed, and its creation/refresh has not been tested. These actions require local Excel software, without an additional online account.

## File guide and attribution

- `NickelOps.xlsx`: working workbook; `previews/`: rendered sheets, filtered views and PDF pages.
- `data/`: raw deliveries, fictional reference tables and initial keyed review decisions.
- `evaluation/`: clean ground truth and issue manifest, isolated from production formulas.
- `exports/`: accepted-delivery and summary CSV snapshots labelled synthetic.
- `scripts/`: data generation, workbook creation/finalization, verification, document generation and packaging source.
- `reports/`: saved-artifact, mutation and reproducibility evidence, plus file checksums.
- [Case study](case_study.md), [data dictionary](data_dictionary.md), [test report](test_report.md) and [2–3 minute demo](demo_script.md).

I developed this independent portfolio with AI assistance for code and documentation. I do not claim to have manually written every line or deployed it in mining operations. No repository-wide license has been assigned. Third-party dependencies are not bundled; their original licenses and terms apply.
''')
    save('data_dictionary.md','''
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
''')
    header='''# Executed test report

SYNTHETIC EDUCATIONAL PORTFOLIO. Author: Ridla A. Gani. Delivery snapshot: 20 November 2026 (Asia/Jakarta). This report records checks executed on the workbook and its inputs.

'''
    body=f"{len(tests)} recorded checks: {len(tests)-fails} PASS, {fails} FAIL. Formula mutation tests used @oai/artifact-tool; independent calculation and saved XLSX inspection used Python {r['python']} and openpyxl {r['openpyxl']}. Tonnage and fraction comparisons use absolute tolerance 1e-7 unless a row states otherwise; identifiers/status/counts use exact comparison. Hand examples use 1e-12.\n\n"
    body+='An earlier development build failed the independent WMT/DMT/Ni controls because empty text constants in the correction register were counted as populated overrides by the available engine. Those inputs were changed to physical blank cells and the complete artifact suite was rerun. A saved-file QA check also exposed COUNTIFS treating empty formula results as populated. The replacement uses ISNUMBER in SUMPRODUCT and was fault-tested: a blank counts 0 and an injected numeric zero counts 1. The table below records the final delivered-artifact results. The detailed evidence retains the earlier QA observation and its correction.\n\n'
    body+='| Check | Expected | Actual | Absolute tolerance | Result | Engine |\n|---|---|---|---:|---|---|\n'
    def brief(v):
        if isinstance(v,float):return f'{v:.12g}'
        text=json.dumps(v,ensure_ascii=True) if isinstance(v,(dict,list,bool)) else str(v)
        return text.replace('|','/').replace('\n',' ')[:250]
    for test in tests:body+='| '+' | '.join([test['name'],brief(test['expected']),brief(test['actual']),str(test['tolerance']),test['result'],test['engine']])+' |\n'
    body+='''
## Visual and package inspection

The nine workbook sheets, keyed correction-input panel and filtered dashboard were rendered to PNG with artifact-tool and inspected. The PDF was generated with reportlab, parsed with pypdf, rendered with pypdfium2 and visually inspected as a two-page management memo. Poppler was not available. Native XML inspection checked formulas/caches, Tables, two chart source bindings, dropdown validations, frozen headers, no VBA and no external workbook links. The review/register mutation tests restored every input before the final export. The saved XLSX was then independently reopened and checked, after OOXML view/print/validation metadata finalization.

The manifest audit checks all 64 issue entries against raw payloads, independent Python statuses and saved workbook statuses. Blank moisture/grade, invalid fractions and negative weights cannot produce accepted zero-valued metrics. Correction tests include a missing-date repair, a missing reason, duplicated correction keys, a Hold decision, an undocumented rejection and reordered keyed entries. Selector tests inspect the affected chart-source values and a filtered rendered dashboard. Native Excel chart redraw has not been verified.

## Testing limitations and manual checks

- Excel COM activation failed with HRESULT 0x80070520 (no usable logon session). No native Excel UI or Excel-engine test is claimed. Cached XLSX values are the tested artifact-tool results, not Excel-calculated evidence.
- Excel dropdown interaction, Stop validation prompts, filter interaction, raw-sheet protection and target-application rendering remain manual checks. Formula gates also validate pasted inputs, but native prompts were only structurally inspected.
- Genuine PivotTables were not installed. Creation/refresh is NOT EXECUTED. Exact creation and refresh instructions appear in README.md and SUMMARY. The installed summaries are SUMIFS formulas, not PivotTables.
- Native Excel refresh: set Automatic calculation, press F9, change a selector, change a keyed decision, change a scenario input, verify affected outputs/charts, restore defaults, then refresh manually created PivotTables. Confirm the accepted filter and reconcile grand totals.
- Fixed capacities and the 60-day daily-summary horizon are deliberate. Appending data or extending periods is not a tested workflow. Do not treat the cached initial QA controls as independent proof after editing inputs; those controls change to REVIEW when the initial snapshot changes.
'''
    save('test_report.md',header+body)
    case=f'''
# NickelOps - Ore Delivery Reconciliation and Scenario Analysis

**Ridla A. Gani | Independent portfolio project | SYNTHETIC EDUCATIONAL DATA**

## Operational question and context

I built NickelOps to examine which reported deliveries a head-office analyst can include, how supplier and weighbridge weights differ, and how dry tonnage and theoretical nickel content respond to stated assumptions. I used synthetic data so I could publish difficult records and their handling without exposing company, employee, vehicle, location or commercial information. All entities and observations are fictional. This project does not represent paid work, actual mine operations, an engineering-approved model or professional mining experience.

## Cleaning and reconciliation method

I generated 300 underlying deliveries across 60 days, four fictional suppliers and three stockpiles using seed 20260901. I kept a separate 310-row raw snapshot with deliberate identifier noise, unknown suppliers, missing measurements, invalid fractions, negative weights, six exact retransmissions and four conflicting submissions. Every raw row remains visible. A 64-entry independent evaluation manifest audits issue coverage; production calculations never read the manifest or ground truth.

I trimmed and standardized identifiers before exact reference lookups. I kept mathematical gates for valid dates, active references, positive weights and fractions in [0,1] separate from business review flags. Illustrative business flags are separate: absolute weight variance >3%, moisture >40% or dry grade >3% require review; these are not industry standards. Corrections are keyed to stable source_row_id values, require approval and a reason, and retain the original measurement. Exact copies are excluded by a deterministic smallest-source-ID rule. Conflicting IDs remain blocked until competing records are documented as rejected. Four initial review examples are disclosed; other ambiguous values are not imputed.

## Verified observations from the synthetic snapshot

- 310 raw rows = **262 accepted + 6 excluded duplicate + 41 unresolved + 1 rejected**. There is no unaccounted row and accepted delivery identifiers are unique.
- Accepted mass is **{t['wmt']:,.3f} WMT**, **{t['dmt']:,.3f} DMT**, aggregate moisture **{t['weighted_moisture']:.4%}**, weighted dry-basis nickel grade **{t['weighted_grade']:.4%}** and theoretical nickel content **{t['contained_ni_t']:,.6f} t**.
- The 41 unresolved rows comprise 3 unknown suppliers, 4 missing dates, 3 missing moisture values, 4 missing grades, 6 invalid fractions, 4 negative-weight records, 11 unapproved material weight differences and 6 still-conflicting submissions. These categories are mutually exclusive in this generated snapshot.
- Twelve original deliveries exceed the illustrative weight-variance threshold. One is explicitly approved in the demonstration and 11 remain blocked. The accepted population therefore includes **one** weight-variance exception. This is a review-state observation, not a supplier-performance conclusion.
- Supplier, stockpile and date summaries reconcile to the same accepted totals within **0.0000001 t**. WMT, DMT and Ni agree with an independent Python calculation and cached values read back from the delivered XLSX.

## Metric definitions

DMT = weighbridge WMT x (1 - moisture fraction). Theoretical Ni = DMT x dry-basis nickel fraction. Aggregate grade is total Ni / total DMT; aggregate moisture is 1 - total DMT / total WMT. Signed variance is weighbridge minus reported WMT, divided by positive reported WMT for its percentage. No unweighted averaging of percentages is used. Theoretical Ni is **not recoverable or payable metal, revenue or profit**.

## Assumption-driven scenario results

The following case applies to the same 262 accepted records, with all dates, suppliers and stockpiles selected. It assumes **+10% relative WMT**, **+2 percentage points moisture**, and **+0.10 percentage points dry-basis grade**. These are user assumptions, not observations or forecasts. Every accepted delivery is recomputed before aggregation.

| Metric | Baseline observation | Assumed combined case |
|---|---:|---:|
| WMT | {t['wmt']:,.3f} | {s['wmt']:,.3f} |
| DMT | {t['dmt']:,.3f} | {s['dmt']:,.3f} |
| Weighted dry grade | {t['weighted_grade']:.4%} | {s['weighted_grade']:.4%} |
| Theoretical Ni tonnes | {t['contained_ni_t']:,.6f} | {s['contained_ni_t']:,.6f} |

Combined theoretical content rises by **{s['contained_ni_t']-t['contained_ni_t']:,.6f} t ({change:.2f}%)**, conditional on all three changes. This is not a production or financial gain. In a one-way case, +2 pp moisture alone reduces DMT to **{r['moisture_only']['dmt']:,.3f}** and theoretical Ni to **{r['moisture_only']['contained_ni_t']:,.6f} t**. A +0.10 pp grade change alone raises theoretical Ni to **{r['grade_only']['contained_ni_t']:,.6f} t**, with unchanged WMT and DMT.

Hand check: 100 WMT at 30% moisture and 1.50% grade gives 70 DMT and 1.05 t Ni. The combined assumptions give 110 WMT, 32% moisture, 1.60% grade, 74.8 DMT and 1.1968 t Ni. Percentage-point inputs are divided by 100 exactly once. Out-of-range assumptions block the case instead of clipping or reporting partial results.

## Practical recommendations supported by the example

1. Reconcile the 41 unresolved rows before using them in an operational total. Prioritize documented resolution of the 11 material weight differences and six conflicting submissions, then obtain missing measurements/reference confirmations. Do not estimate their metal content from absent inputs.
2. Preserve immutable sources, stable row IDs, reasons and separate approval decisions. The demonstrated rejection-without-reason check prevents a competing submission from becoming accepted prematurely.
3. Compare dry tonnage alongside wet tonnage. The +2 pp moisture-only case removes {t['dmt']-r['moisture_only']['dmt']:,.3f} DMT from the same wet mass, illustrating why wet tonnes alone cannot support content comparisons.
4. Agree review thresholds and evidence requirements with qualified operational stakeholders before a real implementation. Keep observed reconciliations separate from hypothetical scenarios and retain weighted denominators in every report.

## Verification, limitations and next steps

{len(tests)} automated checks passed using artifact-tool, Python and openpyxl saved-file inspection. They cover population reconciliation, all injected issues, row metrics, zero/volume/moisture/grade cases, invalid inputs, review-key mutations, scope/chart-source updates and package errors. Workbook and PDF renders were visually inspected. Excel COM was unavailable; no Excel UI or Excel-engine recalculation is claimed. Native Tables/validations/charts are installed but their Excel UI behavior remains to be checked. The workbook contains formula summaries, not genuine PivotTables; exact manual Excel creation/refresh steps are supplied.

Next steps are to validate the workbook in Excel 365, create and reconcile genuine PivotTables, then define evidence ownership, uncertainty, sampling and engineering approval rules for any future real-data design. The fixed-capacity small synthetic sample does not establish mine performance, real savings, recovery, financial impact or operational deployment.
'''
    save('case_study.md',case)
    save('demo_script.md',f'''
# 2-3 minute portfolio walkthrough

**Ridla A. Gani | SYNTHETIC EDUCATIONAL PORTFOLIO**

**0:00-0:20 - Context.** Open START HERE. Say: "I built this independent Excel portfolio using fictional records. It demonstrates delivery reconciliation and assumption analysis, not actual mine experience. The workbook uses no macros or external links."

**0:20-0:45 - One source issue and its review.** In RAW DELIVERIES filter source_row_id to R000041: moisture is blank and the fictional source note records a 32% retest. DATA REVIEW original_moisture_fraction is blank, while effective moisture is 32%. Go to AQ6: the keyed R000041 entry has Approve, a reason and corrected_moisture_fraction 32%. Explain that production does not read evaluation ground truth. Show R000042, whose absent moisture remains Unresolved and whose clean DMT/Ni stay blank.

**0:45-1:10 - Inclusion control.** In an editable copy, set R000041's keyed decision to Hold. Recalculate: the correction requires Approve and the accepted count drops from 262 to 261. Restore Approve. Show the reason requirement and R000307's documented Reject resolving R000081's competing submission. Original rows remain present. Do not edit the calculated review fields.

**1:10-1:35 - Filtered operations.** On DASHBOARD select supplier S02 and stockpile P01, leaving both dates unchanged. Show the selected accepted count/WMT and daily trend. Point to the separately labelled global counts: these do not follow selectors. Restore All/All. The all-scope verified baseline is {t['wmt']:,.3f} WMT, {t['dmt']:,.3f} DMT and {t['contained_ni_t']:,.3f} t theoretical Ni.

**1:35-2:05 - Numerical scenario.** On SCENARIOS explain B6 = 10% relative volume, E6 = 2 moisture percentage points and H6 = 0.10 dry-grade percentage points. For the all-scope case, theoretical Ni becomes {s['contained_ni_t']:,.3f} t. Use the visible 100 WMT example: 30% becomes 32%, 1.50% becomes 1.60%, yielding 74.8 DMT and 1.1968 t Ni after +10% WMT. These are assumptions, not forecasts or recoverable metal. Enter 100 moisture pp to show BLOCKED; restore 2.

**2:05-2:30 - Reconciliation.** Show QA CHECKS: 310 = 262 accepted + 6 excluded duplicate + 41 unresolved + 1 rejected, and supplier WMT reconciliation = zero. SUMMARY contains SUMIFS formula tables, explicitly not PivotTables. Point to the exact Excel steps for genuine PivotTables and refresh.

**2:30-2:50 - Limits.** Say: "I verified the saved workbook and mutation cases with artifact-tool and independent Python. Excel automation was unavailable, so I do not claim Excel UI or native Pivot refresh testing. This fixed synthetic sample supports a transparent portfolio demonstration, not real savings, mine approval or deployment." Finish with the two-page case study and test report.

Before a live Excel demonstration, validate dropdowns, chart redraw, recalculation and protection behavior in Excel 365. Restore default decisions, dates, All/All and 10% / 2 pp / 0.10 pp after the demo. CSV and case-study snapshots do not update with live edits.
''')
    make_pdf(r)
    print('Wrote README, data dictionary, executed test report, case study, demo and two-page PDF.')

def make_pdf(r, output_name='case_study.pdf', preview_prefix='previews/case_study_page_'):
    from reportlab.platypus import SimpleDocTemplate,Paragraph,Spacer,Table,TableStyle,PageBreak,KeepTogether
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet,ParagraphStyle
    from reportlab.lib.enums import TA_LEFT
    from reportlab.lib.pagesizes import A4
    from pypdf import PdfReader
    import pypdfium2 as pdfium
    t=r['totals'];s=r['scenario'];navy=colors.HexColor('#243C55');blue=colors.HexColor('#356B9A');gray=colors.HexColor('#566579')
    styles=getSampleStyleSheet();styles.add(ParagraphStyle(name='BodyN',fontName='Helvetica',fontSize=9.5,leading=13,spaceAfter=7,textColor=colors.HexColor('#23354B')))
    styles.add(ParagraphStyle(name='LabelN',fontName='Helvetica-Bold',fontSize=11,leading=14,textColor=navy,spaceBefore=8,spaceAfter=6))
    styles.add(ParagraphStyle(name='TitleN',fontName='Helvetica-Bold',fontSize=22,leading=25,textColor=navy,spaceAfter=6))
    styles.add(ParagraphStyle(name='SubN',fontName='Helvetica',fontSize=9,leading=12,textColor=gray,spaceAfter=10))
    flow=[]
    def p(text,style='BodyN'):return Paragraph(text,styles[style])
    def para(text,style='BodyN'):flow.append(p(text,style))
    def tab(rows,widths):
        obj=Table([[p(str(c),'BodyN') for c in row] for row in rows],colWidths=widths,hAlign='LEFT')
        obj.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),navy),('TEXTCOLOR',(0,0),(-1,0),colors.white),('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),8),('RIGHTPADDING',(0,0),(-1,-1),8),('TOPPADDING',(0,0),(-1,-1),6),('BOTTOMPADDING',(0,0),(-1,-1),3),('LINEBELOW',(0,-1),(-1,-1),.5,colors.HexColor('#BDC9D4'))]))
        # Paragraph font colors are explicit, so use a dedicated white header style.
        headerstyle=ParagraphStyle(name='thead',fontName='Helvetica-Bold',fontSize=9,leading=12,textColor=colors.white)
        obj=Table([[Paragraph(str(c),headerstyle) if i==0 else p(str(c),'BodyN') for c in row] for i,row in enumerate(rows)],colWidths=widths,hAlign='LEFT')
        obj.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),navy),('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),8),('RIGHTPADDING',(0,0),(-1,-1),8),('TOPPADDING',(0,0),(-1,-1),6),('BOTTOMPADDING',(0,0),(-1,-1),3),('LINEBELOW',(0,-1),(-1,-1),.5,colors.HexColor('#BDC9D4'))]))
        flow.append(obj);flow.append(Spacer(1,8))
    para('NickelOps','TitleN');para('Ore Delivery Reconciliation and Scenario Analysis','LabelN')
    para('Ridla A. Gani | Independent portfolio project | 20 November 2026<br/><b>SYNTHETIC EDUCATIONAL DATA</b>','SubN')
    para('Management question','LabelN');para('Which deliveries can be included in a defensible reconciliation, and how do stated volume, moisture and grade assumptions change dry mass and theoretical nickel content? I used synthetic records to make the methods and deliberate defects publicly reviewable without company or confidential data.')
    para('Observed synthetic-data findings','LabelN')
    tab([['Measure','Verified initial snapshot'],['Row reconciliation','310 raw = 262 accepted + 6 excluded duplicate + 41 unresolved + 1 rejected'],['Accepted WMT / DMT',f"{t['wmt']:,.3f} WMT / {t['dmt']:,.3f} DMT"],['Weighted dry Ni grade',f"{t['weighted_grade']:.4%}"],['Aggregate moisture',f"{t['weighted_moisture']:.4%}"],['Theoretical contained Ni',f"{t['contained_ni_t']:,.6f} t"]],[177,310])
    para('Twelve original deliveries exceed the illustrative 3% absolute weight-variance threshold. One is approved in the demonstration; eleven remain blocked. The accepted population therefore contains one such exception. All supplier, stockpile and date summaries reconcile within 0.0000001 t; accepted business identifiers are unique.')
    para('Cleaning and review controls','LabelN');para('I generated 300 deliveries over 60 days with a fixed seed and retained 310 raw rows, including retransmissions and conflicting submissions. Identifiers are normalized and looked up against fictional references. Missing/invalid dates, weights, moisture and grade block inclusion. Six exact later copies are retained but excluded by smallest source ID. Conflicting IDs stay blocked until competing submissions are rejected with reasons.')
    para('The separate keyed review register preserves original and corrected measurements. Corrections require Approve plus a reason; no other missing measurements are imputed. Four initial decisions are educational examples. The 64-entry evaluation manifest and clean ground truth are isolated from production formulas. Business thresholds (3% variance, 40% moisture, 3% grade) are illustrative assumptions, not industry standards.')
    para('Recommendations from the observed review state','LabelN');para('Resolve the 41 open rows before adding them to operational totals. Prioritize the 11 unapproved material weight differences and six still-conflicting submissions, then obtain missing measurements and reference confirmations. Preserve row keys, source evidence and reasons. Use matching accepted populations and weighted denominators in every comparison.')
    flow.append(PageBreak())
    para('Assumptions and decision limits','TitleN');para('SYNTHETIC EDUCATIONAL DATA | Same 262 accepted deliveries; all-scope comparison','SubN')
    para('Definitions and units','LabelN');para('DMT = weighbridge WMT x (1 - moisture fraction). Theoretical Ni = DMT x dry-basis grade. Weighted grade = total Ni / total DMT; aggregate moisture = 1 - total DMT / total WMT. Signed variance is weighbridge minus reported weight; percentage variance uses positive reported weight. Theoretical content is <b>not recoverable or payable nickel, revenue or profit</b>.')
    para('Combined case: assumptions, not a forecast','LabelN');para('Apply +10% relative WMT, +2 moisture percentage points and +0.10 dry-grade percentage points to every scoped accepted delivery, then recompute and aggregate. Invalid fractions block the entire case; the model does not clip them.')
    tab([['Metric','Baseline observation','Assumed case'],['WMT',f"{t['wmt']:,.3f}",f"{s['wmt']:,.3f}"],['DMT',f"{t['dmt']:,.3f}",f"{s['dmt']:,.3f}"],['Weighted dry grade',f"{t['weighted_grade']:.4%}",f"{s['weighted_grade']:.4%}"],['Theoretical Ni (t)',f"{t['contained_ni_t']:,.6f}",f"{s['contained_ni_t']:,.6f}"]],[177,155,155])
    para(f"The combined case changes theoretical content by {s['contained_ni_t']-t['contained_ni_t']:,.6f} t ({(s['contained_ni_t']/t['contained_ni_t']-1)*100:.2f}%), conditional on all three assumptions. This is not a production or financial gain. In a one-way +2 pp moisture case, DMT falls to {r['moisture_only']['dmt']:,.3f} and theoretical Ni to {r['moisture_only']['contained_ni_t']:,.6f} t. Compare dry tonnes alongside wet tonnes because unchanged wet mass can contain less dry material.")
    para('Percentage-point hand check','LabelN');para('100 WMT at 30% moisture and 1.50% dry grade gives 70 DMT and 1.05 t Ni. The combined assumptions give 110 WMT, 32% moisture and 1.60% dry grade: <b>74.8 DMT and 1.1968 t Ni</b>. Enter moisture change as 2 and grade change as 0.10; divide each by 100 exactly once.')
    para('Verification and limitations','LabelN');para(f"{len(r['tests'])} automated checks passed using artifact-tool, independent Python and openpyxl inspection of cached values in the delivered XLSX. Tests cover row accounting, all injected issues, blank measurements, weighted metrics, scenario directions/scaling, invalid assumptions, keyed review changes and selector/chart-source updates. Workbook previews and this PDF were rendered and visually inspected.")
    para('Excel COM activation failed in the build environment. No Excel UI, Excel-engine recalculation, dropdown interaction or native PivotTable-refresh test is claimed. The workbook contains native Tables, formulas, validations and charts, plus formula summaries <b>rather than installed PivotTables</b>. Exact manual creation and refresh steps are provided. Fixed capacity is 310 source rows, 40 keyed decisions and a 60-day date summary.')
    para('Next steps','LabelN');para('Validate in local Excel 365, create/reconcile genuine PivotTables and agree evidence ownership, uncertainty rules and thresholds with qualified stakeholders before any real-data adaptation. This independent synthetic project does not establish mining experience, actual savings, production gains, engineering approval or operational deployment.')
    def footer(canvas,doc):
        canvas.setStrokeColor(colors.HexColor('#BDC9D4'));canvas.line(54,42,541,42);canvas.setFont('Helvetica',8);canvas.setFillColor(gray);canvas.drawString(54,29,'SYNTHETIC EDUCATIONAL PORTFOLIO | Ridla A. Gani');canvas.drawRightString(541,29,f'{doc.page}')
    SimpleDocTemplate(str(local(output_name)),pagesize=A4,rightMargin=54,leftMargin=54,topMargin=38,bottomMargin=54,title='NickelOps - Synthetic Educational Case Study',author='Ridla A. Gani').build(flow,onFirstPage=footer,onLaterPages=footer)
    pdf=PdfReader(str(local(output_name)))
    if len(pdf.pages)!=2:raise RuntimeError(f'Expected two pages, got {len(pdf.pages)}')
    for i,page in enumerate(pdf.pages):
        text=page.extract_text()
        if 'SYNTHETIC' not in text:raise RuntimeError('Missing synthetic label')
    rendered=pdfium.PdfDocument(str(local(output_name)))
    for i in range(len(rendered)):
        rendered[i].render(scale=1.5).to_pil().save(local(f'{preview_prefix}{i+1}.png'))
    rendered.close()

if __name__=='__main__':main()
