# Executed test report

SYNTHETIC EDUCATIONAL PORTFOLIO. Author: Ridla A. Gani. Delivery snapshot: 20 November 2026 (Asia/Jakarta). This report records checks executed on the workbook and its inputs.

174 recorded checks: 174 PASS, 0 FAIL. Formula mutation tests used @oai/artifact-tool; independent calculation and saved XLSX inspection used Python 3.11.9 and openpyxl 3.1.5. Tonnage and fraction comparisons use absolute tolerance 1e-7 unless a row states otherwise; identifiers/status/counts use exact comparison. Hand examples use 1e-12.

An earlier development build failed the independent WMT/DMT/Ni controls because empty text constants in the correction register were counted as populated overrides by the available engine. Those inputs were changed to physical blank cells and the complete artifact suite was rerun. A saved-file QA check also exposed COUNTIFS treating empty formula results as populated. The replacement uses ISNUMBER in SUMPRODUCT and was fault-tested: a blank counts 0 and an injected numeric zero counts 1. The table below records the final delivered-artifact results. The detailed evidence retains the earlier QA observation and its correction.

| Check | Expected | Actual | Absolute tolerance | Result | Engine |
|---|---|---|---:|---|---|
| Raw rows partition mutually exclusively | 310 | 310 | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| Initial raw row count | 310 | 310 | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| Ground-truth delivery count | 300 | 300 | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| Every source_row_id is unique and stable | 310 | 310 | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| Accepted business identifiers unique | 262 | 262 | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| Saved workbook statuses agree for every source ID | [] | [] | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| Accepted identifiers, reference IDs and dates agree with Python | [] | [] | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| Saved workbook per-record WMT/DMT/Ni agrees with independent Python | [] | [] | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| All non-accepted measurements remain blank, not zero | [] | [] | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| Saved RAW DELIVERIES preserves every original field | [] | [] | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| Saved dashboard independent control: count | 262 | 262 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Saved dashboard independent control: wmt | 8542.047 | 8542.047 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Saved dashboard independent control: dmt | 5939.5402752 | 5939.5402752 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Saved dashboard independent control: weighted_grade | 0.0156036210904 | 0.0156036210904 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Saved dashboard independent control: contained_ni_t | 92.6783359053 | 92.6783359053 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Saved dashboard independent control: weighted_moisture | 0.304670148127 | 0.304670148127 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Saved dashboard independent control: weight_exceptions | 1 | 1 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Saved supplier summary reconciles wmt | 8542.047 | 8542.047 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Saved stockpile summary reconciles wmt | 8542.047 | 8542.047 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Saved date summary reconciles wmt | 8542.047 | 8542.047 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Saved supplier summary reconciles dmt | 5939.5402752 | 5939.5402752 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Saved stockpile summary reconciles dmt | 5939.5402752 | 5939.5402752 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Saved date summary reconciles dmt | 5939.5402752 | 5939.5402752 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Saved supplier summary reconciles contained_ni_t | 92.6783359053 | 92.6783359053 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Saved stockpile summary reconciles contained_ni_t | 92.6783359053 | 92.6783359053 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Saved date summary reconciles contained_ni_t | 92.6783359053 | 92.6783359053 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| supplier S01 independently agrees: count | 67 | 67 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| supplier S01 independently agrees: wmt | 2057.387 | 2057.387 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| supplier S01 independently agrees: dmt | 1426.14688749 | 1426.14688749 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| supplier S01 independently agrees: contained_ni_t | 22.3416675411 | 22.3416675411 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| supplier S01 independently agrees: weighted_grade | 0.0156657548651 | 0.0156657548651 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| supplier S01 independently agrees: weighted_moisture | 0.306816419328 | 0.306816419328 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| supplier S02 independently agrees: count | 64 | 64 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| supplier S02 independently agrees: wmt | 2105.891 | 2105.891 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| supplier S02 independently agrees: dmt | 1474.36725318 | 1474.36725318 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| supplier S02 independently agrees: contained_ni_t | 23.1644197929 | 23.1644197929 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| supplier S02 independently agrees: weighted_grade | 0.0157114312889 | 0.0157114312889 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| supplier S02 independently agrees: weighted_moisture | 0.299884346731 | 0.299884346731 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| supplier S03 independently agrees: count | 66 | 66 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| supplier S03 independently agrees: wmt | 2198.018 | 2198.018 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| supplier S03 independently agrees: dmt | 1530.90044468 | 1530.90044468 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| supplier S03 independently agrees: contained_ni_t | 24.2262197463 | 24.2262197463 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| supplier S03 independently agrees: weighted_grade | 0.0158248172378 | 0.0158248172378 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| supplier S03 independently agrees: weighted_moisture | 0.303508686153 | 0.303508686153 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| supplier S04 independently agrees: count | 65 | 65 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| supplier S04 independently agrees: wmt | 2180.751 | 2180.751 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| supplier S04 independently agrees: dmt | 1508.12568985 | 1508.12568985 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| supplier S04 independently agrees: contained_ni_t | 22.946028825 | 22.946028825 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| supplier S04 independently agrees: weighted_grade | 0.0152149313411 | 0.0152149313411 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| supplier S04 independently agrees: weighted_moisture | 0.308437464958 | 0.308437464958 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| stockpile P01 independently agrees: count | 80 | 80 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| stockpile P01 independently agrees: wmt | 2604.295 | 2604.295 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| stockpile P01 independently agrees: dmt | 1818.73109647 | 1818.73109647 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| stockpile P01 independently agrees: contained_ni_t | 27.5935553038 | 27.5935553038 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| stockpile P01 independently agrees: weighted_grade | 0.0151718719482 | 0.0151718719482 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| stockpile P01 independently agrees: weighted_moisture | 0.301641674053 | 0.301641674053 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| stockpile P02 independently agrees: count | 103 | 103 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| stockpile P02 independently agrees: wmt | 3330.975 | 3330.975 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| stockpile P02 independently agrees: dmt | 2311.36970569 | 2311.36970569 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| stockpile P02 independently agrees: contained_ni_t | 36.6717247225 | 36.6717247225 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| stockpile P02 independently agrees: weighted_grade | 0.0158657979432 | 0.0158657979432 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| stockpile P02 independently agrees: weighted_moisture | 0.306098152736 | 0.306098152736 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| stockpile P03 independently agrees: count | 79 | 79 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| stockpile P03 independently agrees: wmt | 2606.777 | 2606.777 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| stockpile P03 independently agrees: dmt | 1809.43947304 | 1809.43947304 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| stockpile P03 independently agrees: contained_ni_t | 28.4130558789 | 28.4130558789 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| stockpile P03 independently agrees: weighted_grade | 0.0157026837882 | 0.0157026837882 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| stockpile P03 independently agrees: weighted_moisture | 0.305871015035 | 0.305871015035 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Visible QA checks all pass | 13 | 13 | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| All injected manifest entries detected and handled as expected | [] | [] | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| Independent manifest entry count | 64 | 64 | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| Accepted normalized/corrected measurements agree with evaluation truth | [] | [] | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Fixed-seed generation reproduces all data/evaluation bytes | true | true | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| Python zero scenario wmt | 8542.047 | 8542.047 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Python +10% volume wmt | 9396.2517 | 9396.2517 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Python zero scenario dmt | 5939.5402752 | 5939.5402752 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Python +10% volume dmt | 6533.49430272 | 6533.49430272 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Python zero scenario contained_ni_t | 92.6783359053 | 92.6783359053 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Python +10% volume contained_ni_t | 101.946169496 | 101.946169496 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Python moisture raises WMT-independent losses | true | true | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| Python grade leaves WMT and DMT unchanged | true | true | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| Hand example scenario DMT: 110 x (1 - .32) | 74.8 | 74.8 | 1e-12 | PASS | Python 3 stdlib / openpyxl read-only |
| Hand example scenario theoretical Ni: 74.8 x .016 | 1.1968 | 1.1968 | 1e-12 | PASS | Python 3 stdlib / openpyxl read-only |
| Python invalid scenario blocked moisture | true | true | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| Python invalid scenario blocked grade | true | true | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| Python invalid scenario blocked volume | true | true | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| Saved compound scenario wmt | 9396.2517 | 9396.2517 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Saved compound scenario dmt | 6345.56926872 | 6345.56926872 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Saved compound scenario contained_ni_t | 105.35917986 | 105.35917986 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Moisture one-way -4 pp: dmt | 6281.2221552 | 6281.2221552 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Moisture one-way -4 pp: contained_ni_t | 98.0102611854 | 98.0102611854 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Grade one-way -0.4 pp: dmt | 5939.5402752 | 5939.5402752 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Grade one-way -0.4 pp: contained_ni_t | 68.9201748045 | 68.9201748045 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Moisture one-way -2 pp: dmt | 6110.3812152 | 6110.3812152 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Moisture one-way -2 pp: contained_ni_t | 95.3442985453 | 95.3442985453 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Grade one-way -0.2 pp: dmt | 5939.5402752 | 5939.5402752 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Grade one-way -0.2 pp: contained_ni_t | 80.7992553549 | 80.7992553549 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Moisture one-way +0 pp: dmt | 5939.5402752 | 5939.5402752 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Moisture one-way +0 pp: contained_ni_t | 92.6783359053 | 92.6783359053 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Grade one-way +0.0 pp: dmt | 5939.5402752 | 5939.5402752 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Grade one-way +0.0 pp: contained_ni_t | 92.6783359053 | 92.6783359053 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Moisture one-way +2 pp: dmt | 5768.6993352 | 5768.6993352 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Moisture one-way +2 pp: contained_ni_t | 90.0123732652 | 90.0123732652 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Grade one-way +0.2 pp: dmt | 5939.5402752 | 5939.5402752 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Grade one-way +0.2 pp: contained_ni_t | 104.557416456 | 104.557416456 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Moisture one-way +4 pp: dmt | 5597.8583952 | 5597.8583952 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Moisture one-way +4 pp: contained_ni_t | 87.3464106251 | 87.3464106251 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Grade one-way +0.4 pp: dmt | 5939.5402752 | 5939.5402752 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Grade one-way +0.4 pp: contained_ni_t | 116.436497006 | 116.436497006 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Mutated selected scope independently agrees: count | 13 | 13 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Mutated selected scope independently agrees: wmt | 425.768 | 425.768 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Mutated selected scope independently agrees: dmt | 296.702034535 | 296.702034535 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Mutated selected scope independently agrees: ni | 4.57946820033 | 4.57946820033 | 1e-07 | PASS | Python 3 stdlib / openpyxl read-only |
| Initial Python control wmt | 8542.047 | 8542.047 | 1e-07 | PASS | @oai/artifact-tool |
| Initial Python control dmt | 5939.5402752 | 5939.5402752 | 1e-07 | PASS | @oai/artifact-tool |
| Initial Python control ni | 92.6783359053 | 92.6783359053 | 1e-07 | PASS | @oai/artifact-tool |
| Zero-change scenario wmt | 8542.047 | 8542.047 | 1e-07 | PASS | @oai/artifact-tool |
| Zero-change scenario dmt | 5939.5402752 | 5939.5402752 | 1e-07 | PASS | @oai/artifact-tool |
| Zero-change scenario ni | 92.6783359053 | 92.6783359053 | 1e-07 | PASS | @oai/artifact-tool |
| Higher moisture reduces DMT | true | true | 1e-07 | PASS | @oai/artifact-tool |
| Higher moisture reduces Ni | true | true | 1e-07 | PASS | @oai/artifact-tool |
| Higher moisture leaves WMT | 8542.047 | 8542.047 | 1e-07 | PASS | @oai/artifact-tool |
| Higher grade leaves WMT | 8542.047 | 8542.047 | 1e-07 | PASS | @oai/artifact-tool |
| Higher grade leaves DMT | 5939.5402752 | 5939.5402752 | 1e-07 | PASS | @oai/artifact-tool |
| Grade pp uses division by 100 once | 98.6178761805 | 98.6178761805 | 1e-07 | PASS | @oai/artifact-tool |
| 10% volume scales wmt | 9396.2517 | 9396.2517 | 1e-07 | PASS | @oai/artifact-tool |
| 10% volume scales dmt | 6533.49430272 | 6533.49430272 | 1e-07 | PASS | @oai/artifact-tool |
| 10% volume scales ni | 101.946169496 | 101.946169496 | 1e-07 | PASS | @oai/artifact-tool |
| Minus 100% volume valid | 0 | 0 | 1e-07 | PASS | @oai/artifact-tool |
| Zero DMT grade unavailable | Unavailable | Unavailable | 1e-07 | PASS | @oai/artifact-tool |
| Volume below -100% blocked | true | true | 1e-07 | PASS | @oai/artifact-tool |
| Invalid scenario total unavailable | Unavailable | Unavailable | 1e-07 | PASS | @oai/artifact-tool |
| Out-of-range moisture blocked | true | true | 1e-07 | PASS | @oai/artifact-tool |
| Out-of-range grade blocked | true | true | 1e-07 | PASS | @oai/artifact-tool |
| Blank assumption blocked | true | true | 1e-07 | PASS | @oai/artifact-tool |
| Supplier/stockpile selector count | 13 | 13 | 0 | PASS | @oai/artifact-tool |
| Supplier/stockpile selector WMT | 425.768 | 425.768 | 1e-07 | PASS | @oai/artifact-tool |
| Filtered chart source reconciles | 425.768 | 425.768 | 1e-07 | PASS | @oai/artifact-tool |
| Filtered Ni chart source updates | 4.57946820033 | 4.57946820033 | 1e-07 | PASS | @oai/artifact-tool |
| Reversed dates blocked | true | true | 1e-07 | PASS | @oai/artifact-tool |
| Invalid scope KPI unavailable | Unavailable | Unavailable | 1e-07 | PASS | @oai/artifact-tool |
| Keyed date correction accepts intended record | Accepted | Accepted | 1e-07 | PASS | @oai/artifact-tool |
| Keyed correction increases count by one | 263 | 263 | 0 | PASS | @oai/artifact-tool |
| Neighboring row unchanged | Unresolved | Unresolved | 1e-07 | PASS | @oai/artifact-tool |
| Correction without reason blocked | Unresolved | Unresolved | 1e-07 | PASS | @oai/artifact-tool |
| Duplicate override keys block record | Unresolved | Unresolved | 1e-07 | PASS | @oai/artifact-tool |
| Undocumented rejection keeps competing record blocked | Unresolved | Unresolved | 1e-07 | PASS | @oai/artifact-tool |
| Reordered correction register retains accepted count | 262 | 262 | 0 | PASS | @oai/artifact-tool |
| Reordered correction register retains DMT | 5939.5402752 | 5939.5402752 | 1e-07 | PASS | @oai/artifact-tool |
| Keyed Hold decreases accepted population | 261 | 261 | 0 | PASS | @oai/artifact-tool |
| All inputs restored | 8542.047 | 8542.047 | 1e-07 | PASS | @oai/artifact-tool |
| QA ignores non-accepted empty formula results | 0 | 0 | 0 | PASS | @oai/artifact-tool |
| QA detects a false numeric zero in non-accepted DMT | 1 | 1 | 0 | PASS | @oai/artifact-tool |
| Saved XLSX import dashboard count | 262 | 262 | 1e-07 | PASS | @oai/artifact-tool XLSX import/recalculation |
| Saved XLSX import dashboard WMT | 8542.047 | 8542.047 | 1e-07 | PASS | @oai/artifact-tool XLSX import/recalculation |
| Saved XLSX import dashboard_date_filtered count | 112 | 112 | 1e-07 | PASS | @oai/artifact-tool XLSX import/recalculation |
| Saved XLSX import dashboard_date_filtered WMT | 3666.257 | 3666.257 | 1e-07 | PASS | @oai/artifact-tool XLSX import/recalculation |
| Saved XLSX import dashboard_filtered count | 13 | 13 | 1e-07 | PASS | @oai/artifact-tool XLSX import/recalculation |
| Saved XLSX import dashboard_filtered WMT | 425.768 | 425.768 | 1e-07 | PASS | @oai/artifact-tool XLSX import/recalculation |
| Runtime mutation tests all pass | 0 | 0 | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| Delivered workbook contains no unexpected cached formula errors | [] | [] | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| Production formulas never reference evaluation files | [] | [] | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| Required native formula present: VLOOKUP | true | true | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| Required native formula present: SUMIFS | true | true | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| Required native formula present: IF | true | true | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| No external workbook links or macros | [] | [] | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| No genuine PivotTables are falsely claimed | 0 | 0 | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| Two native formula-backed charts installed | 2 | 2 | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| Both native chart value axes start at zero | true | true | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| Native chart sources point to formula cells in DASHBOARD | true | true | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| Native Excel Tables installed | true | true | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| RAW DELIVERIES protection safeguard installed | true | true | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| Raw-sheet protection metadata has one valid element | 1 | 1 | 0 | PASS | Python 3 stdlib / openpyxl read-only |
| Named validation lists installed | true | true | 0 | PASS | Python 3 stdlib / openpyxl read-only |

## Visual and package inspection

The nine workbook sheets, keyed correction-input panel and filtered dashboard were rendered to PNG with artifact-tool and inspected. The PDF was generated with reportlab, parsed with pypdf, rendered with pypdfium2 and visually inspected as a two-page management memo. Poppler was not available. Native XML inspection checked formulas/caches, Tables, two chart source bindings, dropdown validations, frozen headers, no VBA and no external workbook links. The review/register mutation tests restored every input before the final export. The saved XLSX was then independently reopened and checked, after OOXML view/print/validation metadata finalization.

The manifest audit checks all 64 issue entries against raw payloads, independent Python statuses and saved workbook statuses. Blank moisture/grade, invalid fractions and negative weights cannot produce accepted zero-valued metrics. Correction tests include a missing-date repair, a missing reason, duplicated correction keys, a Hold decision, an undocumented rejection and reordered keyed entries. Selector tests inspect the affected chart-source values and a filtered rendered dashboard. Native Excel chart redraw has not been verified.

## Testing limitations and manual checks

- Excel COM activation failed with HRESULT 0x80070520 (no usable logon session). No native Excel UI or Excel-engine test is claimed. Cached XLSX values are the tested artifact-tool results, not Excel-calculated evidence.
- Excel dropdown interaction, Stop validation prompts, filter interaction, raw-sheet protection and target-application rendering remain manual checks. Formula gates also validate pasted inputs, but native prompts were only structurally inspected.
- Genuine PivotTables were not installed. Creation/refresh is NOT EXECUTED. Exact creation and refresh instructions appear in README.md and SUMMARY. The installed summaries are SUMIFS formulas, not PivotTables.
- Native Excel refresh: set Automatic calculation, press F9, change a selector, change a keyed decision, change a scenario input, verify affected outputs/charts, restore defaults, then refresh manually created PivotTables. Confirm the accepted filter and reconcile grand totals.
- Fixed capacities and the 60-day daily-summary horizon are deliberate. Appending data or extending periods is not a tested workflow. Do not treat the cached initial QA controls as independent proof after editing inputs; those controls change to REVIEW when the initial snapshot changes.
