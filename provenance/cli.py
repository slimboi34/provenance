"""Command line: provenance convert | rebuild | check."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

from . import certificate, diff
from .canon import ConversionError, build_package, dumps, read_package, sha256
from .compiler import compile_workbook, load_module
from .model import read_workbook, to_markdown
from .verify import fingerprint, round_trip

EXIT = {"verified": 0, "failed": 1, "partial": 2}


def convert(path: Path, out_dir: Path, fuzz_cases: int = 25, seed: int = 0, engine: str = "auto") -> dict:
    """Convert one file, write every artifact into out_dir, return the certificate."""
    source = path.read_bytes()
    pkg = read_package(source)
    package_json = dumps(pkg)
    structure = round_trip(source, pkg)

    out_dir.mkdir(parents=True, exist_ok=True)
    stem = path.stem
    (out_dir / f"{stem}.package.json").write_bytes(package_json)

    notes: list[str] = []
    model = compile_report = replay = fuzz = logic = None
    try:
        model = read_workbook(pkg)
    except ConversionError as exc:
        notes.append(f"Not a workbook ({exc}); only the structure layer applies.")
    if model is not None:
        (out_dir / f"{stem}.workbook.json").write_bytes(json.dumps(model, indent=1, ensure_ascii=False).encode())
        (out_dir / f"{stem}.md").write_text(to_markdown(model, path.name, sha256(source)), encoding="utf-8")
        logic, compile_report = compile_workbook(model, sha256(source))
        (out_dir / f"{re.sub(r'[^0-9A-Za-z_]', '_', stem)}_logic.py").write_text(logic, encoding="utf-8")
        namespace = load_module(logic)
        replay = diff.replay(model, namespace)
        if model["external_link_parts"]:
            notes.append("Workbook links to other workbooks; formulas using those links are not verified.")
        for sheet in model["sheets"]:
            notes += [f"{sheet['name']}: {n}" for n in sheet["notes"]]
        if compile_report["compiled"] and fuzz_cases > 0:
            chosen, engine_note = diff.find_engine(engine)
            if engine_note:
                notes.append(engine_note)
            if chosen is None:
                notes.append("No independent engine available; fuzz testing skipped.")
            elif not diff.variable_inputs(model):
                notes.append("No inputs to vary; fuzz testing skipped.")
            else:
                fuzz = diff.fuzz(pkg, model, namespace, chosen, fuzz_cases, seed)
                if fuzz["status"] == "engine_failed":
                    notes.append(f"Fuzz testing did not run: {fuzz['reason']}")

    cert = certificate.build(path.name, source, package_json, structure, model, compile_report, replay, fuzz, logic, notes)
    (out_dir / f"{stem}.certificate.json").write_bytes(json.dumps(cert, indent=1, ensure_ascii=False).encode())
    return cert


def check(cert_path: Path, source_path: Path | None) -> list[str]:
    """Re-verify a certificate against the artifacts beside it (and the source, if given)."""
    cert = json.loads(cert_path.read_text(encoding="utf-8"))
    problems = certificate.check_seal(cert)
    stem = cert_path.name.removesuffix(".certificate.json")
    package_path = cert_path.with_name(f"{stem}.package.json")
    if not package_path.exists():
        return problems + [f"{package_path.name} not found next to the certificate"]
    package_json = package_path.read_bytes()
    if sha256(package_json) != cert["artifacts"]["package_json_sha256"]:
        problems.append("package JSON does not match the certificate")
    rebuilt = fingerprint(build_package(json.loads(package_json)))
    recorded = [(p["name"], p["method"], p["source_sha256"]) for p in cert["layers"]["structure"]["parts"]]
    if [(e["name"], e["method"], e["sha256"]) for e in rebuilt] != recorded:
        problems.append("file rebuilt from the package JSON does not match the certified source content")
    if source_path is not None:
        source = source_path.read_bytes()
        if sha256(source) != cert["source"]["sha256"]:
            problems.append("source file is not the one that was certified")
        elif [(e["name"], e["method"], e["sha256"]) for e in fingerprint(source)] != recorded:
            problems.append("source content does not match the certificate")
    return problems


def _print_summary(cert: dict, out_dir: Path) -> None:
    v = cert["verdict"]
    print(f"{v['level'].upper()}: {cert['source']['name']}")
    print(f"  {v['summary']}")
    formulas = cert["layers"].get("formulas", {})
    fz = formulas.get("fuzz")
    if fz and fz.get("status") == "ran":
        c, cal = fz["counts"], fz["calibration"]
        print(f"  Fuzz: {fz['cases']} cases against {fz['engine']}: {c['exact']} exact, {c['close']} float-close, {c['mismatch']} mismatches"
              f" (engine agrees with Excel on {cal['engine_agrees_with_excel']}/{cal['cells_with_stored_value']} stored results).")
    reasons: dict[str, int] = {}
    for reason in formulas.get("unverified", {}).values():
        reason = re.sub(r"depends on \S+, which is unverified: ", "", reason)
        reasons[reason] = reasons.get(reason, 0) + 1
    for reason, n in sorted(reasons.items(), key=lambda kv: -kv[1])[:8]:
        print(f"  not verified ({n} cells): {reason}")
    for note in cert["notes"][:8]:
        print(f"  note: {note}")
    print(f"  Artifacts in {out_dir}/")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="provenance", description="Verified conversion of Excel workbooks to JSON and Python.")
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("convert", help="convert a file and certify the conversion")
    p.add_argument("file", type=Path)
    p.add_argument("-o", "--out", type=Path, help="output directory (default: <file>.provenance)")
    p.add_argument("--fuzz", type=int, default=25, metavar="N", help="random input cases to test (0 = off, default 25)")
    p.add_argument("--seed", type=int, default=0)
    p.add_argument("--engine", choices=["auto", "libreoffice", "formulas", "none"], default="auto")
    p.add_argument("--json", action="store_true", help="print the certificate instead of a summary")

    p = sub.add_parser("rebuild", help="turn a package JSON back into the original file format")
    p.add_argument("package", type=Path)
    p.add_argument("-o", "--out", type=Path, required=True)

    p = sub.add_parser("check", help="re-verify a certificate against its artifacts")
    p.add_argument("certificate", type=Path)
    p.add_argument("--source", type=Path, help="also check the original file")

    args = parser.parse_args(argv)
    try:
        if args.command == "convert":
            out_dir = args.out or args.file.with_name(args.file.name + ".provenance")
            cert = convert(args.file, out_dir, args.fuzz, args.seed, args.engine)
            if args.json:
                print(json.dumps(cert, indent=1, ensure_ascii=False))
            else:
                _print_summary(cert, out_dir)
            return EXIT[cert["verdict"]["level"]]
        if args.command == "rebuild":
            args.out.write_bytes(build_package(json.loads(args.package.read_bytes())))
            print(f"wrote {args.out}")
            return 0
        problems = check(args.certificate, args.source)
        for problem in problems:
            print(f"FAIL: {problem}")
        if not problems:
            print("OK: certificate is intact and the artifacts match it")
            if "signature" not in json.loads(args.certificate.read_text(encoding="utf-8")):
                print("WARNING: the certificate is not signed. Its checksum catches accidental changes, but anyone "
                      f"can edit it and recompute the checksum. Set {certificate.KEY_ENV} to sign certificates.")
        return 1 if problems else 0
    except (ConversionError, OSError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
