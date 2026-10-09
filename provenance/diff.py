"""Differential testing of the compiled workbook.

Two independent checks:

* replay - recompute every formula from the stored inputs and compare with
  the result Excel itself saved in the file;
* fuzz   - change the numeric inputs at random and compare the compiled module
  with a separate spreadsheet engine recalculating the same changed workbook.

A comparison is `exact` (equal values), `close` (floats equal to 1e-9
relative: same number, different last bits) or `mismatch`.
"""

from __future__ import annotations

import copy
import functools
import math
import os
import random
import shutil
import subprocess
import tempfile
from pathlib import Path

from .canon import attr, build_package, find_part, kid, kids, read_package, root
from .model import formula_cells, key, read_workbook
from .runtime import Context, Unsupported, XlError, circular_cells, error
from .runtime import _finish as finish

RANK = {"exact": 0, "close": 1, "mismatch": 2}


def plain(v):
    """Normalise a runtime or model value for comparison and for JSON."""
    if isinstance(v, XlError):
        return {"error": v.code}
    if isinstance(v, bool) or v is None or isinstance(v, str):
        return v
    return float(v)


def cached(cell: dict):
    """Excel's stored result for a cell; (False, None) if there is none."""
    if "e" in cell:
        return True, {"error": cell["e"]}
    if "v" in cell:
        return True, cell["v"]
    return False, None


def compare(a, b, bool_as_number: bool = False) -> str:
    """exact / close (floats within 1e-9 relative) / mismatch.

    bool_as_number: the file was saved by a program that stores TRUE/FALSE as
    1/0 (LibreOffice does), so a logical result may match a stored number.
    """
    if bool_as_number and isinstance(a, bool) != isinstance(b, bool) and isinstance(a, (bool, float)) and isinstance(b, (bool, float)):
        return "close" if float(a) == float(b) else "mismatch"
    if isinstance(a, float) and isinstance(b, float) and not isinstance(a, bool) and not isinstance(b, bool):
        if a == b:
            return "exact"
        return "close" if math.isclose(a, b, rel_tol=1e-9, abs_tol=1e-12) else "mismatch"
    return "exact" if type(a) is type(b) and a == b else "mismatch"


def saved_by_excel(model: dict) -> bool:
    return "Microsoft" in (model.get("application") or "") and "Excel" in (model.get("application") or "")


# ---------------------------------------------------------------------- replay

def _runtime_value(v):
    return error(v["error"]) if isinstance(v, dict) else v


def replay(model: dict, namespace: dict) -> dict:
    """Compare compiled formulas with the results stored in the file, two ways.

    Per formula: each formula is computed from the *stored* values of the cells
    it reads, so it is judged on its own logic; this decides the cell's status.
    End to end: compute() recalculates everything from the inputs; a cell that
    passes on its own but differs here has picked up floating-point rounding
    from the cells before it ("drift").
    """
    lenient = not saved_by_excel(model)
    stored = {}
    for cell_key, _sheet, _ref, cell in formula_cells(model):
        has, value = cached(cell)
        if has:
            stored[cell_key] = value
    local_values = dict(namespace["INPUTS"])
    local_values.update({k: _runtime_value(v) for k, v in stored.items()})
    ctx = Context(local_values, namespace["FORMULAS"], namespace["UNSUPPORTED"])

    values, failed = namespace["compute"]()
    circular = circular_cells(namespace.get("DEPS", {}))
    cells: dict[str, str] = {}
    mismatches, drift = [], []
    for cell_key, _sheet, _ref, cell in formula_cells(model):
        f = namespace["FORMULAS"].get(cell_key)
        if f is None:
            cells[cell_key] = "unverified"
            continue
        if cell_key not in stored:
            cells[cell_key] = "unverified" if cell_key in failed else "no_stored_value"
            continue
        if cell_key in circular or "circular reference" in failed.get(cell_key, ""):
            # Excel stores 0 (or a partial iteration) for unresolved circular cells,
            # so the stored values in and after a cycle are not answers to check.
            cells[cell_key] = "unverified"
            continue
        try:
            got = plain(finish(f(ctx)))
        except Unsupported:
            cells[cell_key] = "unverified"
            continue
        except Exception:  # same rule as the evaluator: a runtime gap is not a value
            cells[cell_key] = "unverified"
            continue
        expected = stored[cell_key]
        cells[cell_key] = result = compare(got, expected, lenient)
        if result == "mismatch":
            mismatches.append({"cell": cell_key, "formula": cell["f"], "excel": expected, "compiled": got})
        elif cell_key in values:
            end_to_end = plain(values[cell_key])
            if compare(end_to_end, expected, lenient) == "mismatch":
                drift.append({"cell": cell_key, "formula": cell["f"], "excel": expected, "end_to_end": end_to_end})
    counts = {s: sum(1 for v in cells.values() if v == s) for s in ("exact", "close", "mismatch", "no_stored_value", "unverified")}
    failed = dict(failed)
    for cell_key in circular:
        failed.setdefault(cell_key, "on or after a circular reference")
    return {"counts": counts, "mismatches": mismatches[:100], "drift": drift[:100], "drift_count": len(drift),
            "cells": cells, "failed": failed, "stored_by": model.get("application")}


# ----------------------------------------------------------------------- cases

def variable_inputs(model: dict) -> dict[str, float | bool | str]:
    """Input cells (no formula) holding a number, TRUE/FALSE or text."""
    out = {}
    for sheet in model["sheets"]:
        for ref, cell in sheet["cells"].items():
            v = cell.get("v")
            if "f" not in cell and "unreadable" not in cell and isinstance(v, (float, bool, str)):
                out[key(sheet["name"], ref)] = v
    return out


def make_cases(model: dict, count: int, seed: int) -> list[dict]:
    """Case 0 changes nothing (it calibrates the engine); the rest perturb inputs.

    Numbers are scaled, zeroed or negated; TRUE/FALSE are flipped; text is
    swapped for other text from the same column (so lookups and IF tests take
    other branches) or occasionally emptied.
    """
    rng = random.Random(seed)
    inputs = sorted(variable_inputs(model).items())
    texts_by_column: dict[tuple, list[str]] = {}
    for cell_key, v in inputs:
        if isinstance(v, str):
            texts_by_column.setdefault(_column(cell_key), []).append(v)
    cases: list[dict] = [{}]
    for _ in range(count):
        case = {}
        for cell_key, v in inputs:
            roll = rng.random()
            if isinstance(v, bool):
                if roll < 0.3:
                    case[cell_key] = not v
                continue
            if isinstance(v, str):
                others = [t for t in texts_by_column[_column(cell_key)] if t != v]
                if roll < 0.25 and others:
                    case[cell_key] = rng.choice(others)
                elif roll < 0.28:
                    case[cell_key] = ""
                continue
            if roll < 0.25:
                continue
            if roll < 0.33:
                new = 0.0
            elif roll < 0.41:
                new = -v
            elif v == 0:
                new = float(rng.randint(-3, 3))
            else:
                scaled = v * rng.uniform(0.5, 1.5)
                new = float(round(scaled)) if v.is_integer() else round(scaled, 4)
            if new != v:
                case[cell_key] = new + 0.0
        cases.append(case)
    return cases


def _column(cell_key: str) -> tuple[str, str]:
    sheet, ref = cell_key.rsplit("!", 1)
    return sheet, ref.rstrip("0123456789")


# --------------------------------------------------------------------- engines

class EngineError(Exception):
    pass


def _set_value(c: dict, value) -> None:
    """Rewrite an input cell element to hold `value` (number, bool or text)."""
    prefix = c["tag"][:-1]  # "c" or "x:c" -> "" or "x:"
    attrs = [a for a in c.get("attrs", []) if a[0] != "t"]
    keep = [ch for ch in c["children"] if not (isinstance(ch, dict) and ch.get("tag", "").rsplit(":", 1)[-1] in ("v", "is"))]
    if isinstance(value, bool):
        attrs.append(["t", "b"])
        keep.append({"tag": prefix + "v", "children": ["1" if value else "0"]})
    elif isinstance(value, str):
        attrs.append(["t", "inlineStr"])
        t = {"tag": prefix + "t", "attrs": [["xml:space", "preserve"]], "children": [value] if value else []}
        keep.append({"tag": prefix + "is", "children": [t]})
    else:
        keep.append({"tag": prefix + "v", "children": [repr(float(value))]})
    c["attrs"], c["children"] = attrs, keep


def mutated_file(pkg: dict, model: dict, overrides: dict[str, float]) -> bytes:
    """The workbook with new input values and every stored formula result removed."""
    out = {"format": pkg["format"], "parts": list(pkg["parts"])}
    by_part = {s["part"]: s["name"] for s in model["sheets"] if s["part"]}
    pending = set(overrides)
    for i, part in enumerate(out["parts"]):
        sheet = by_part.get(part["name"])
        if sheet is None or part["kind"] != "xml":
            continue
        part = out["parts"][i] = copy.deepcopy(part)
        data = kid(root(part["xml"]), "sheetData")
        for row in kids(data, "row") if data else ():
            for c in kids(row, "c"):
                is_formula = kid(c, "f") is not None
                cell_key = key(sheet, (attr(c, "r") or "").upper())
                if is_formula:
                    c["children"] = [ch for ch in c["children"] if not (isinstance(ch, dict) and ch.get("tag", "").rsplit(":", 1)[-1] == "v")]
                    c["attrs"] = [a for a in c.get("attrs", []) if a[0] != "t"]
                elif cell_key in overrides:
                    _set_value(c, overrides[cell_key])
                    pending.discard(cell_key)
    if pending:
        raise EngineError(f"could not set inputs: {sorted(pending)[:5]}")
    return build_package(out)


def _formula_values(model: dict) -> dict:
    out = {}
    for cell_key, _sheet, _ref, cell in formula_cells(model):
        has, value = cached(cell)
        if has:
            out[cell_key] = value
    return out


@functools.lru_cache(maxsize=None)
def _responds(binary: str) -> bool:
    try:
        return subprocess.run([binary, "--headless", "--version"], capture_output=True, timeout=8).returncode == 0
    except (subprocess.TimeoutExpired, OSError):
        return False


class LibreOffice:
    """Recalculate with LibreOffice Calc running headless."""

    def __init__(self, binary: str):
        self.binary = binary
        self.name = "LibreOffice"

    @staticmethod
    def find() -> "LibreOffice | None":
        candidates = [os.environ.get("PROVENANCE_SOFFICE"), shutil.which("soffice"), shutil.which("libreoffice"),
                      "/Applications/LibreOffice.app/Contents/MacOS/soffice"]
        for path in candidates:
            if path and os.path.exists(path):
                return LibreOffice(path)
        return None

    def responds(self) -> bool:
        """False if the program hangs (first-launch approval pending, or blocked by a sandbox)."""
        return _responds(self.binary)

    def run(self, pkg: dict, model: dict, cases: list[dict]) -> list[dict]:
        with tempfile.TemporaryDirectory(prefix="provenance-lo-") as tmp:
            tmp_path = Path(tmp)
            (tmp_path / "in").mkdir()
            names = []
            for i, case in enumerate(cases):
                name = f"case{i:04d}.xlsx"
                (tmp_path / "in" / name).write_bytes(mutated_file(pkg, model, case))
                names.append(name)
            cmd = [self.binary, "--headless", "--norestore", f"-env:UserInstallation={(tmp_path / 'profile').as_uri()}",
                   "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(tmp_path / "out"),
                   *(str(tmp_path / "in" / n) for n in names)]
            try:
                subprocess.run(cmd, capture_output=True, timeout=120 + 5 * len(cases), check=False)
            except subprocess.TimeoutExpired:
                raise EngineError("LibreOffice did not finish (if it was just installed, open it once by hand first)") from None
            results = []
            for name in names:
                path = tmp_path / "out" / name
                if not path.exists():
                    raise EngineError(f"LibreOffice produced no output for {name}")
                results.append(_formula_values(read_workbook(read_package(path.read_bytes()))))
            return results


class FormulasLibrary:
    """Recalculate with the pure-Python `formulas` interpreter (pip install formulas)."""

    name = "formulas (Python library)"

    @staticmethod
    def find() -> "FormulasLibrary | None":
        try:
            import formulas  # noqa: F401
        except Exception:
            return None
        return FormulasLibrary()

    def run(self, pkg: dict, model: dict, cases: list[dict]) -> list[dict]:
        import formulas
        import logging
        logging.getLogger("formulas").setLevel(logging.ERROR)
        with tempfile.TemporaryDirectory(prefix="provenance-fx-") as tmp:
            path = Path(tmp) / "book.xlsx"
            path.write_bytes(mutated_file(pkg, model, {}))
            try:
                xl = formulas.ExcelModel().loads(str(path)).finish()
            except Exception as exc:
                raise EngineError(f"formulas could not load the workbook: {type(exc).__name__}: {exc}") from None

            def ref(cell_key: str) -> str:
                sheet, a1_ref = cell_key.rsplit("!", 1)
                return f"'[book.xlsx]{sheet.upper()}'!{a1_ref}"

            wanted = {ref(k): k for k, *_ in formula_cells(model)}
            results = []
            for case in cases:
                try:
                    solution = xl.calculate(inputs={ref(k): v for k, v in case.items()})
                except Exception as exc:
                    raise EngineError(f"formulas failed to calculate: {type(exc).__name__}: {exc}") from None
                out = {}
                for name, cell_key in wanted.items():
                    if name in solution:
                        try:
                            out[cell_key] = self._plain(solution[name].value[0][0])
                        except Exception:
                            pass
                results.append(out)
            return results

    @staticmethod
    def _plain(v):
        if type(v).__name__ == "XlError":
            return {"error": str(v)}
        if hasattr(v, "item"):
            v = v.item()
        if isinstance(v, bool) or isinstance(v, str):
            return v
        if isinstance(v, (int, float)):
            return float(v)
        if v is None or type(v).__name__ == "Empty":
            return 0.0
        raise TypeError(type(v).__name__)


def find_engine(choice: str = "auto"):
    """Return (engine or None, note or None)."""
    if choice == "none":
        return None, None
    note = None
    if choice in ("auto", "libreoffice"):
        engine = LibreOffice.find()
        if engine and engine.responds():
            return engine, None
        if engine:
            note = ("LibreOffice is installed but did not start within 8 seconds (a pending first-launch prompt, or a "
                    "sandbox blocking it); the formulas library was used instead.")
        if choice == "libreoffice":
            return None, note or "LibreOffice was not found."
    engine = FormulasLibrary.find()
    if engine is None:
        note = (note + " " if note else "") + "The `formulas` package is not installed."
    return engine, note


# ------------------------------------------------------------------------ fuzz

def fuzz(pkg: dict, model: dict, namespace: dict, engine, count: int, seed: int) -> dict:
    cases = make_cases(model, count, seed)
    report: dict = {"engine": engine.name, "seed": seed, "cases": count, "inputs_varied": len(variable_inputs(model))}
    try:
        oracle = engine.run(pkg, model, cases)
    except EngineError as exc:
        report.update(status="engine_failed", reason=str(exc))
        return report

    # Calibration: does the engine itself reproduce Excel's stored results?
    stored = _formula_values(model)
    agree = sum(1 for k, v in stored.items() if k in oracle[0] and compare(oracle[0][k], v) != "mismatch")
    report["calibration"] = {"cells_with_stored_value": len(stored), "engine_agrees_with_excel": agree}

    counts = {"exact": 0, "close": 0, "mismatch": 0, "engine_gave_no_value": 0}
    cells: dict[str, str] = {}
    mismatches = []
    for i, (case, expected) in enumerate(zip(cases, oracle)):
        if i == 0:
            continue
        values, failed = namespace["compute"](case)
        for cell_key in namespace["FORMULAS"]:
            if cell_key in failed:
                continue
            if cell_key not in expected:
                counts["engine_gave_no_value"] += 1
                continue
            got = plain(values[cell_key])
            result = compare(got, expected[cell_key])
            counts[result] += 1
            if RANK[result] > RANK.get(cells.get(cell_key, "exact"), 0) or cell_key not in cells:
                cells[cell_key] = result
            if result == "mismatch" and len(mismatches) < 100:
                mismatches.append({"case": i, "cell": cell_key, "engine": expected[cell_key], "compiled": got, "inputs_changed": case})
    report.update(status="ran", counts=counts, cells=cells, mismatches=mismatches)
    return report
