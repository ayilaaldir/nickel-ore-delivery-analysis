"""Boundary checks shared by the reproducible local project scripts."""
from pathlib import Path
import os

ROOT = Path(__file__).resolve().parent.parent
EXPECTED = "NickelOps-Ore-Delivery-Reconciliation-and-Scenario-Analysis"
if ROOT.name != EXPECTED or Path.cwd().resolve() != ROOT:
    raise RuntimeError("Run from the existing authorized project root.")

def local(relative):
    candidate = ROOT / relative
    resolved = candidate.resolve()
    if not resolved.is_relative_to(ROOT) or candidate.is_symlink():
        raise RuntimeError(f"Path outside authorized root: {relative}")
    return resolved

for name in ("data", "evaluation", "exports", "reports", "previews", "tmp", "tmp/cache", "tmp/pdfs"):
    local(name).mkdir(parents=True, exist_ok=True)
os.environ["TMP"] = os.environ["TEMP"] = str(local("tmp"))
os.environ["PYTHONPYCACHEPREFIX"] = str(local("tmp/cache"))
