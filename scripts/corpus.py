"""Run `convert` over a folder of files and summarise what breaks.

    .venv/bin/python scripts/corpus.py DIR [--fuzz N] [--jobs J] [--timeout S] [--json OUT]

Each file runs in its own process with a time limit, so one hang or crash
cannot stop the run. Formula checking here is the replay against the
results Excel saved in each file.
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
import tempfile
import time
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SUFFIXES = {".xlsx", ".xlsm", ".docx"}


def one(path: str, fuzz: int) -> dict:
    """Convert one file in this process and return a summary (worker mode)."""
    import logging
    import traceback
    import warnings
    sys.path.insert(0, str(ROOT))
    from provenance.canon import ConversionError
    from provenance.cli import convert
    logging.disable(logging.CRITICAL)
    warnings.simplefilter("ignore")
    start = time.time()
    with tempfile.TemporaryDirectory() as tmp:
        try:
            cert = convert(Path(path), Path(tmp), fuzz_cases=fuzz, engine="formulas" if fuzz else "none")
        except ConversionError as exc:
            return {"status": "rejected", "reason": str(exc)[:200], "seconds": time.time() - start}
        except Exception as exc:
            tb = traceback.extract_tb(exc.__traceback__)[-1]
            return {"status": "crash", "reason": f"{type(exc).__name__}: {str(exc)[:160]} at {Path(tb.filename).name}:{tb.lineno}",
                    "seconds": time.time() - start}
    layers = cert["layers"]
    formulas = layers.get("formulas", {})
    return {
        "status": "ok",
        "verdict": cert["verdict"]["level"],
        "structure": layers["structure"]["status"],
        "structure_problems": layers["structure"]["problems"][:5],
        "formula_total": formulas.get("total", 0),
        "formula_counts": formulas.get("counts", {}),
        "replay": formulas.get("replay", {}).get("counts", {}),
        "stored_by": formulas.get("replay", {}).get("stored_by"),
        "drift": formulas.get("replay", {}).get("end_to_end_drift", {}).get("count", 0),
        "mismatches": formulas.get("replay", {}).get("mismatches", [])[:20],
        "fuzz": {k: v for k, v in formulas.get("fuzz", {}).items() if k in ("status", "counts", "reason")},
        "unverified": list(formulas.get("unverified", {}).values()),
        "seconds": time.time() - start,
    }


def run(path: Path, fuzz: int, timeout: int) -> dict:
    cmd = [sys.executable, __file__, "--one", str(path), "--fuzz", str(fuzz)]
    try:
        out = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
    except subprocess.TimeoutExpired:
        return {"status": "timeout", "reason": f"over {timeout}s"}
    lines = [ln for ln in out.stdout.splitlines() if ln.startswith("{")]
    if not lines:
        return {"status": "crash", "reason": "worker died: " + out.stderr.strip()[-300:]}
    return json.loads(lines[-1])


def reason_key(reason: str) -> str:
    reason = re.sub(r"depends on \S+, which is unverified: ", "", reason)
    reason = re.sub(r"'[^']*'", "'…'", reason)
    return re.sub(r"\(\w+Error: .*", "(error)", reason)[:110]


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("dir", type=Path, nargs="?")
    parser.add_argument("--one")
    parser.add_argument("--fuzz", type=int, default=0)
    parser.add_argument("--jobs", type=int, default=6)
    parser.add_argument("--timeout", type=int, default=180)
    parser.add_argument("--json", type=Path)
    args = parser.parse_args()
    if args.one:
        print(json.dumps(one(args.one, args.fuzz), default=str))
        return 0

    files = sorted(p for p in args.dir.rglob("*") if p.suffix.lower() in SUFFIXES and p.is_file() and "/.git/" not in str(p))
    with ThreadPoolExecutor(args.jobs) as pool:
        results = dict(zip(files, pool.map(lambda p: run(p, args.fuzz, args.timeout), files)))
    if args.json:
        args.json.write_text(json.dumps({str(k.relative_to(args.dir)): v for k, v in results.items()}, indent=1, default=str))

    status = Counter(r["status"] for r in results.values())
    ok = [r for r in results.values() if r["status"] == "ok"]
    print(f"{len(files)} files: " + ", ".join(f"{n} {s}" for s, n in status.most_common()))
    print(f"structure round trip: {sum(r['structure'] == 'proven' for r in ok)} proven, {sum(r['structure'] != 'proven' for r in ok)} failed")
    print("verdicts: " + ", ".join(f"{n} {v}" for v, n in Counter(r["verdict"] for r in ok).most_common()))
    replay = Counter()
    for r in ok:
        replay.update(r["replay"])
    print(f"formulas: {sum(r['formula_total'] for r in ok)} total; replay vs Excel's stored results: {dict(replay)}")
    if any(r["fuzz"] for r in ok):
        fz = Counter()
        for r in ok:
            fz.update(r["fuzz"].get("counts", {}))
        print(f"fuzz: {dict(fz)}")

    for kind in ("crash", "timeout"):
        for path, r in results.items():
            if r["status"] == kind:
                print(f"  {kind.upper()} {path.relative_to(args.dir)}: {r['reason']}")
    for path, r in results.items():
        if r["status"] == "ok" and r["structure"] != "proven":
            print(f"  ROUND TRIP FAILED {path.relative_to(args.dir)}: {r['structure_problems'][:2]}")
    rejected = Counter(r["reason"][:80] for r in results.values() if r["status"] == "rejected")
    for reason, n in rejected.most_common(6):
        print(f"  rejected x{n}: {reason}")
    excel = [r for r in ok if "Excel" in (r.get("stored_by") or "")]
    by_excel = Counter()
    for r in excel:
        by_excel.update(r["replay"])
    print(f"  of which in files saved by Microsoft Excel ({len(excel)} files): {dict(by_excel)}")
    print(f"end-to-end drift (correct per formula, rounding differs downstream): {sum(r.get('drift', 0) for r in ok)} cells")
    print("replay mismatches (formula: stored -> ours):")
    for path, r in results.items():
        for m in r.get("mismatches", []) if r["status"] == "ok" else []:
            print(f"  [{(r.get('stored_by') or '?').split('/')[0][:16]}] {path.relative_to(args.dir)} {m['cell']}  "
                  f"={m['formula'][:70]}  stored={m['excel']!r} ours={m['compiled']!r}")
    reasons = Counter(reason_key(u) for r in ok for u in r["unverified"])
    print("top reasons formulas were not verified:")
    for reason, n in reasons.most_common(25):
        print(f"  {n:6}  {reason}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
