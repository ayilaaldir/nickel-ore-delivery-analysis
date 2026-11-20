"""Create deterministic, wholly fictional educational delivery data. Python stdlib only."""
import csv
import json
import random
from datetime import date, timedelta
from project_paths import local

SEED = 20260901
FIELDS = ["source_row_id", "delivery_id", "delivery_date", "supplier_id", "stockpile_id",
          "supplier_reported_wmt", "weighbridge_net_wmt", "moisture_pct", "ni_grade_pct_dry", "record_note"]

def write_csv(path, fields, rows):
    with local(path).open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)

def main():
    rng = random.Random(SEED)
    suppliers = [dict(supplier_id=f"S{i:02}", supplier_name=name, active_status="Active") for i, name in enumerate(
        ["Luma Ore Cooperative", "Tera Mineral Supply", "Navo Earth Partners", "Koro Materials"], 1)]
    piles = [dict(stockpile_id=f"P{i:02}", stockpile_name=name) for i, name in enumerate(
        ["Cedar Blend", "Quartz Blend", "Amber Blend"], 1)]
    truth = []
    for i in range(1, 301):
        s = (i - 1) % 4 + 1
        pile = rng.randrange(1, 4)
        wmt = round(rng.uniform(24, 42) + (s - 2.5) * 0.6, 3)
        moisture = round(rng.uniform(0.25, 0.37) + (pile - 2) * 0.006, 6)
        grade = round(rng.uniform(0.012, 0.019) + (pile - 2) * 0.0004, 6)
        reported = round(wmt / (1 + rng.uniform(-0.014, 0.014)), 3)
        note = "Synthetic educational delivery."
        if 91 <= i <= 102:
            reported = round(wmt / (1 + (0.08 if i % 2 else -0.09)), 3)
            note = "Synthetic material weight difference; investigation required."
        row = dict(zip(FIELDS, [f"R{i:06}", f"DEL-{i:05}", (date(2026, 9, 1) + timedelta(days=(i-1)//5)).isoformat(),
                 f"S{s:02}", f"P{pile:02}", reported, wmt, moisture, grade, note]))
        truth.append(row)
    # The fictional retest used in the demonstration must agree with evaluation truth.
    truth[40]['moisture_pct'] = 0.32
    raw = [dict(x) for x in truth]
    manifest = []
    def issue(i, kind, handling, **changes):
        raw[i-1].update(changes)
        manifest.append(dict(source_row_id=raw[i-1]["source_row_id"], injected_issue_type=kind, expected_handling=handling))
    for i in range(1, 13):
        issue(i, "IDENTIFIER_NORMALIZATION", "Accepted after trim and upper-case normalization",
              supplier_id=" " + raw[i-1]["supplier_id"].lower() + " ", stockpile_id=raw[i-1]["stockpile_id"].lower(),
              delivery_id=" " + raw[i-1]["delivery_id"].lower() + " ")
    for i in range(21, 25):
        issue(i, "UNKNOWN_SUPPLIER", "Accepted with keyed correction" if i == 21 else "Unresolved", supplier_id="S99")
    raw[20]["record_note"] = "Synthetic supplier confirmation: S01. Demonstration correction evidence."
    for i in range(31, 35):
        issue(i, "MISSING_DATE", "Unresolved", delivery_date="")
    for i in range(41, 45):
        issue(i, "MISSING_MOISTURE", "Accepted with keyed correction" if i == 41 else "Unresolved", moisture_pct="")
    raw[40]["record_note"] = "Synthetic retest confirms moisture fraction 0.320000. Demonstration evidence."
    for i in range(51, 55):
        issue(i, "MISSING_GRADE", "Unresolved", ni_grade_pct_dry="")
    for i, value in zip(range(61, 64), [1.08, -0.04, 1.2]):
        issue(i, "INVALID_MOISTURE_FRACTION", "Unresolved", moisture_pct=value)
    for i, value in zip(range(64, 67), [-0.02, 1.1, 1.02]):
        issue(i, "INVALID_GRADE_FRACTION", "Unresolved", ni_grade_pct_dry=value)
    for i in (71, 72):
        issue(i, "NEGATIVE_REPORTED_WEIGHT", "Unresolved", supplier_reported_wmt=-abs(raw[i-1]["supplier_reported_wmt"]))
    for i in (73, 74):
        issue(i, "NEGATIVE_WEIGHBRIDGE_WEIGHT", "Unresolved", weighbridge_net_wmt=-abs(raw[i-1]["weighbridge_net_wmt"]))
    for i in range(91, 103):
        issue(i, "MATERIAL_WEIGHT_DIFFERENCE", "Accepted after documented approval" if i == 91 else "Unresolved")
    for i in range(151, 157):
        copy = dict(raw[i-1]); copy["source_row_id"] = f"R{len(raw)+1:06}"
        raw.append(copy)
        manifest.append(dict(source_row_id=copy["source_row_id"], injected_issue_type="EXACT_RETRANSMISSION", expected_handling="Excluded duplicate"))
    for i in range(81, 85):
        copy = dict(raw[i-1]); copy["source_row_id"] = f"R{len(raw)+1:06}"
        copy["weighbridge_net_wmt"] = round(copy["weighbridge_net_wmt"] + 0.4, 3)
        copy["record_note"] = "Synthetic conflicting submission; same delivery identifier."
        raw.append(copy)
        manifest.append(dict(source_row_id=copy["source_row_id"], injected_issue_type="CONFLICTING_DUPLICATE", expected_handling="Rejected with reason" if i == 81 else "Unresolved"))
        manifest.append(dict(source_row_id=raw[i-1]["source_row_id"], injected_issue_type="CONFLICTING_DUPLICATE", expected_handling="Accepted after competing submission rejected" if i == 81 else "Unresolved"))
    correction_fields = ["source_row_id", "decision", "reason", "corrected_date", "corrected_supplier_id", "corrected_stockpile_id",
                         "corrected_reported_wmt", "corrected_weighbridge_wmt", "corrected_moisture_fraction", "corrected_grade_fraction", "reviewer_notes"]
    decisions = []
    for key, action, reason, changes in [
        ("R000021", "Approve", "Synthetic supplier confirmation in raw record_note identifies S01.", {"corrected_supplier_id": "S01"}),
        ("R000041", "Approve", "Synthetic retest in raw record_note confirms moisture 0.320000.", {"corrected_moisture_fraction": 0.32}),
        ("R000307", "Reject", "Demonstration adjudication: retain R000081 and reject conflicting submission.", {}),
        ("R000091", "Approve", "Demonstration weight difference reviewed; retain weighbridge basis; no imputation.", {})]:
        d = dict.fromkeys(correction_fields, ""); d.update(source_row_id=key, decision=action, reason=reason, **changes)
        decisions.append(d)
    write_csv("data/raw_deliveries.csv", FIELDS, raw)
    write_csv("data/supplier_reference.csv", list(suppliers[0])+['synthetic_notice'], [dict(x,synthetic_notice='SYNTHETIC EDUCATIONAL DATA') for x in suppliers])
    write_csv("data/stockpile_reference.csv", list(piles[0])+['synthetic_notice'], [dict(x,synthetic_notice='SYNTHETIC EDUCATIONAL DATA') for x in piles])
    write_csv("data/review_decisions.csv", correction_fields, decisions)
    write_csv("evaluation/clean_ground_truth.csv", FIELDS, truth)
    write_csv("evaluation/issue_manifest.csv", list(manifest[0])+['synthetic_notice'], [dict(x,synthetic_notice='SYNTHETIC EVALUATION ONLY') for x in manifest])
    payload = dict(seed=SEED, synthetic=True, raw=raw, suppliers=suppliers, stockpiles=piles, decisions=decisions,
                   thresholds=dict(weight_variance=0.03, moisture_review=0.40, grade_review=0.03))
    local("data/workbook_input.json").write_text(json.dumps(payload, indent=2), encoding="utf-8")
    print(json.dumps(dict(seed=SEED, ground_truth=len(truth), raw=len(raw), issue_manifest_entries=len(manifest))))

if __name__ == "__main__":
    main()
