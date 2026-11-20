# NickelOps - Ore Delivery Reconciliation and Scenario Analysis

**Ridla A. Gani | Independent portfolio project | SYNTHETIC EDUCATIONAL DATA**

## Operational question and context

I built NickelOps to examine which reported deliveries a head-office analyst can include, how supplier and weighbridge weights differ, and how dry tonnage and theoretical nickel content respond to stated assumptions. I used synthetic data so I could publish difficult records and their handling without exposing company, employee, vehicle, location or commercial information. All entities and observations are fictional. This project does not represent paid work, actual mine operations, an engineering-approved model or professional mining experience.

## Cleaning and reconciliation method

I generated 300 underlying deliveries across 60 days, four fictional suppliers and three stockpiles using seed 20260901. I kept a separate 310-row raw snapshot with deliberate identifier noise, unknown suppliers, missing measurements, invalid fractions, negative weights, six exact retransmissions and four conflicting submissions. Every raw row remains visible. A 64-entry independent evaluation manifest audits issue coverage; production calculations never read the manifest or ground truth.

I trimmed and standardized identifiers before exact reference lookups. I kept mathematical gates for valid dates, active references, positive weights and fractions in [0,1] separate from business review flags. Illustrative business flags are separate: absolute weight variance >3%, moisture >40% or dry grade >3% require review; these are not industry standards. Corrections are keyed to stable source_row_id values, require approval and a reason, and retain the original measurement. Exact copies are excluded by a deterministic smallest-source-ID rule. Conflicting IDs remain blocked until competing records are documented as rejected. Four initial review examples are disclosed; other ambiguous values are not imputed.

## Verified observations from the synthetic snapshot

- 310 raw rows = **262 accepted + 6 excluded duplicate + 41 unresolved + 1 rejected**. There is no unaccounted row and accepted delivery identifiers are unique.
- Accepted mass is **8,542.047 WMT**, **5,939.540 DMT**, aggregate moisture **30.4670%**, weighted dry-basis nickel grade **1.5604%** and theoretical nickel content **92.678336 t**.
- The 41 unresolved rows comprise 3 unknown suppliers, 4 missing dates, 3 missing moisture values, 4 missing grades, 6 invalid fractions, 4 negative-weight records, 11 unapproved material weight differences and 6 still-conflicting submissions. These categories are mutually exclusive in this generated snapshot.
- Twelve original deliveries exceed the illustrative weight-variance threshold. One is explicitly approved in the demonstration and 11 remain blocked. The accepted population therefore includes **one** weight-variance exception. This is a review-state observation, not a supplier-performance conclusion.
- Supplier, stockpile and date summaries reconcile to the same accepted totals within **0.0000001 t**. WMT, DMT and Ni agree with an independent Python calculation and cached values read back from the delivered XLSX.

## Metric definitions

DMT = weighbridge WMT x (1 - moisture fraction). Theoretical Ni = DMT x dry-basis nickel fraction. Aggregate grade is total Ni / total DMT; aggregate moisture is 1 - total DMT / total WMT. Signed variance is weighbridge minus reported WMT, divided by positive reported WMT for its percentage. No unweighted averaging of percentages is used. Theoretical Ni is **not recoverable or payable metal, revenue or profit**.

## Assumption-driven scenario results

The following case applies to the same 262 accepted records, with all dates, suppliers and stockpiles selected. It assumes **+10% relative WMT**, **+2 percentage points moisture**, and **+0.10 percentage points dry-basis grade**. These are user assumptions, not observations or forecasts. Every accepted delivery is recomputed before aggregation.

| Metric | Baseline observation | Assumed combined case |
|---|---:|---:|
| WMT | 8,542.047 | 9,396.252 |
| DMT | 5,939.540 | 6,345.569 |
| Weighted dry grade | 1.5604% | 1.6604% |
| Theoretical Ni tonnes | 92.678336 | 105.359180 |

Combined theoretical content rises by **12.680844 t (13.68%)**, conditional on all three changes. This is not a production or financial gain. In a one-way case, +2 pp moisture alone reduces DMT to **5,768.699** and theoretical Ni to **90.012373 t**. A +0.10 pp grade change alone raises theoretical Ni to **98.617876 t**, with unchanged WMT and DMT.

Hand check: 100 WMT at 30% moisture and 1.50% grade gives 70 DMT and 1.05 t Ni. The combined assumptions give 110 WMT, 32% moisture, 1.60% grade, 74.8 DMT and 1.1968 t Ni. Percentage-point inputs are divided by 100 exactly once. Out-of-range assumptions block the case instead of clipping or reporting partial results.

## Practical recommendations supported by the example

1. Reconcile the 41 unresolved rows before using them in an operational total. Prioritize documented resolution of the 11 material weight differences and six conflicting submissions, then obtain missing measurements/reference confirmations. Do not estimate their metal content from absent inputs.
2. Preserve immutable sources, stable row IDs, reasons and separate approval decisions. The demonstrated rejection-without-reason check prevents a competing submission from becoming accepted prematurely.
3. Compare dry tonnage alongside wet tonnage. The +2 pp moisture-only case removes 170.841 DMT from the same wet mass, illustrating why wet tonnes alone cannot support content comparisons.
4. Agree review thresholds and evidence requirements with qualified operational stakeholders before a real implementation. Keep observed reconciliations separate from hypothetical scenarios and retain weighted denominators in every report.

## Verification, limitations and next steps

174 automated checks passed using artifact-tool, Python and openpyxl saved-file inspection. They cover population reconciliation, all injected issues, row metrics, zero/volume/moisture/grade cases, invalid inputs, review-key mutations, scope/chart-source updates and package errors. Workbook and PDF renders were visually inspected. Excel COM was unavailable; no Excel UI or Excel-engine recalculation is claimed. Native Tables/validations/charts are installed but their Excel UI behavior remains to be checked. The workbook contains formula summaries, not genuine PivotTables; exact manual Excel creation/refresh steps are supplied.

Next steps are to validate the workbook in Excel 365, create and reconcile genuine PivotTables, then define evidence ownership, uncertainty, sampling and engineering approval rules for any future real-data design. The fixed-capacity small synthetic sample does not establish mine performance, real savings, recovery, financial impact or operational deployment.
