"""Independent Python reference model. Never imports the workbook builder or ground truth."""
import csv
from collections import Counter
from datetime import date
from project_paths import local

def csv_rows(path):
    with local(path).open(newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))

def number(value):
    try:
        return float(value) if value != "" and value is not None else None
    except (ValueError, TypeError):
        return None

def analyze(raw=None, decisions=None):
    raw = raw if raw is not None else csv_rows("data/raw_deliveries.csv")
    decisions = decisions if decisions is not None else csv_rows("data/review_decisions.csv")
    supplier_ids = {x["supplier_id"] for x in csv_rows("data/supplier_reference.csv")}
    pile_ids = {x["stockpile_id"] for x in csv_rows("data/stockpile_reference.csv")}
    dcounts = Counter(x["source_row_id"] for x in decisions if x["source_row_id"])
    keyed = {x["source_row_id"]: x for x in decisions if x["source_row_id"]}
    seen = {}; result = []
    for original in sorted(raw, key=lambda x: x["source_row_id"]):
        r = dict(original); key = r["source_row_id"]; d = keyed.get(key, {})
        signature = tuple(str(original[k]) for k in original if k != "source_row_id")
        retransmission = signature in seen; seen.setdefault(signature, key)
        decision = d.get("decision") or "Auto"
        reason = d.get("reason", "").strip()
        overrides = {"delivery_date": "corrected_date", "supplier_id": "corrected_supplier_id", "stockpile_id": "corrected_stockpile_id",
                     "supplier_reported_wmt": "corrected_reported_wmt", "weighbridge_net_wmt": "corrected_weighbridge_wmt",
                     "moisture_pct": "corrected_moisture_fraction", "ni_grade_pct_dry": "corrected_grade_fraction"}
        changed = False
        for field, override in overrides.items():
            if d.get(override, "") != "":
                r[field] = d[override]; changed = True
        for field in ("delivery_id", "supplier_id", "stockpile_id"):
            r[field] = str(r[field]).strip().upper()
        issues = []
        if not r["delivery_id"]: issues.append("Missing delivery ID")
        try: date.fromisoformat(str(r["delivery_date"]))
        except ValueError: issues.append("Missing/invalid date")
        if r["supplier_id"] not in supplier_ids: issues.append("Unknown/inactive supplier")
        if r["stockpile_id"] not in pile_ids: issues.append("Unknown stockpile")
        for field, label, fraction in [("supplier_reported_wmt", "Reported WMT", False), ("weighbridge_net_wmt", "Weighbridge WMT", False),
                                       ("moisture_pct", "Moisture", True), ("ni_grade_pct_dry", "Grade", True)]:
            value = number(r[field]); r[field] = value
            if value is None or (not fraction and value <= 0) or (fraction and not 0 <= value <= 1): issues.append(f"Missing/invalid {label}")
        control_ok = dcounts[key] <= 1 and decision in ("Auto", "Approve", "Reject", "Hold")
        control_ok &= not ((changed or decision in ("Approve", "Reject")) and not reason)
        control_ok &= not (changed and decision != "Approve")
        weight_ok = r["supplier_reported_wmt"] is not None and r["supplier_reported_wmt"] > 0 and r["weighbridge_net_wmt"] is not None and r["weighbridge_net_wmt"] > 0
        variance = r["weighbridge_net_wmt"] - r["supplier_reported_wmt"] if weight_ok else None
        pct = variance / r["supplier_reported_wmt"] if weight_ok else None
        review = not issues and (abs(pct) > 0.03 or r["moisture_pct"] > 0.4 or r["ni_grade_pct_dry"] > 0.03)
        r.update(decision=decision, reason=reason, control_ok=bool(control_ok), retransmission=retransmission, issues=issues,
                 needs_review=review, signed_variance_wmt=variance, variance_fraction=pct)
        result.append(r)
    live_counts = Counter(x["delivery_id"] for x in result if not x["retransmission"] and not (x["decision"] == "Reject" and x["control_ok"]))
    for r in result:
        if r["retransmission"]: status = "Excluded duplicate"
        elif not r["control_ok"]: status = "Unresolved"
        elif r["decision"] == "Reject": status = "Rejected"
        elif r["issues"] or live_counts[r["delivery_id"]] > 1 or r["decision"] == "Hold" or (r["needs_review"] and r["decision"] != "Approve"): status = "Unresolved"
        else: status = "Accepted"
        r["status"] = status
        r["dmt"] = r["weighbridge_net_wmt"] * (1-r["moisture_pct"]) if status == "Accepted" else None
        r["contained_ni_t"] = r["dmt"] * r["ni_grade_pct_dry"] if status == "Accepted" else None
    return result

def aggregate(rows):
    rows = list(rows)
    wmt = sum(x["weighbridge_net_wmt"] for x in rows)
    dmt = sum(x["dmt"] for x in rows)
    ni = sum(x["contained_ni_t"] for x in rows)
    return dict(count=len(rows), wmt=wmt, dmt=dmt, contained_ni_t=ni, weighted_grade=ni/dmt if dmt else None,
                weighted_moisture=1-dmt/wmt if wmt else None, weight_exceptions=sum(abs(x["variance_fraction"]) > .03 for x in rows))

def scenario(rows, volume=0.0, moisture_pp=0.0, grade_pp=0.0):
    if volume < -1: raise ValueError("Volume change must be at least -100%")
    out = []
    for r in rows:
        w = r["weighbridge_net_wmt"]*(1+volume); m=r["moisture_pct"]+moisture_pp/100; g=r["ni_grade_pct_dry"]+grade_pp/100
        if not (0 <= m <= 1 and 0 <= g <= 1): raise ValueError("Scenario fraction outside [0,1]")
        out.append(dict(r, weighbridge_net_wmt=w, moisture_pct=m, ni_grade_pct_dry=g, dmt=w*(1-m), contained_ni_t=w*(1-m)*g))
    return aggregate(out)
