"""Readable workbook model derived from the canonical package.

This is a *view*: cell values, formula text and Excel's cached results per
sheet. Styles, charts and layout stay in the canonical package only. Values
are read straight from the sheet XML, so numbers are exactly what Excel
stored (dates stay serial numbers).
"""

from __future__ import annotations

import posixpath

from .canon import ConversionError, attr, find_part, kid, kids, root, text_of
from .formula import FormulaError, a1, col_letters, shift_formula, split_a1

FORMAT = "provenance.workbook/1"
MAX_MODEL_TEXT = 200_000_000  # guards against one huge shared string referenced by many cells


def _rels(pkg: dict, part_name: str) -> dict[str, str]:
    folder, base = posixpath.split(part_name)
    rels = find_part(pkg, posixpath.join(folder, "_rels", base + ".rels"))
    out = {}
    if rels and rels["kind"] == "xml":
        for rel in kids(root(rels["xml"]), "Relationship"):
            target = attr(rel, "Target", "")
            if attr(rel, "TargetMode") == "External":
                continue
            out[attr(rel, "Id")] = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(folder, target))
    return out


def _workbook_part(pkg: dict) -> str:
    for target in _rels(pkg, "").values():
        if posixpath.basename(target).startswith("workbook"):
            return target
    if find_part(pkg, "xl/workbook.xml"):
        return "xl/workbook.xml"
    raise ConversionError("no workbook part")


def _rich_text(el: dict) -> str:
    """Text of an <si> or <is>: plain <t> plus rich-text runs, minus phonetics."""
    out = []
    for child in kids(el):
        name = child["tag"].rsplit(":", 1)[-1]
        if name == "t":
            out.append(text_of(child))
        elif name == "r":
            out.extend(text_of(t) for t in kids(child, "t"))
    return "".join(out)


def _shared_strings(pkg: dict, wb_rels: dict[str, str]) -> list[str]:
    name = next((t for t in wb_rels.values() if posixpath.basename(t).lower().startswith("sharedstrings")), None)
    part = find_part(pkg, name) if name else None
    if not part or part["kind"] != "xml":
        return []
    return [_rich_text(si) for si in kids(root(part["xml"]), "si")]


def _read_sheet(tree: dict, strings: list[str]) -> tuple[dict, list[str]]:
    cells: dict[str, dict] = {}
    notes: list[str] = []
    shared: dict[str, tuple[str, int, int]] = {}
    blocks: list[tuple[str, str, str, str]] = []  # array formulas and data tables: (range, master, formula, kind)
    data = kid(root(tree), "sheetData")
    next_row = 1
    for row in kids(data, "row") if data else ():
        row_no = int(attr(row, "r", next_row))
        next_row = row_no + 1
        next_col = 1
        for c in kids(row, "c"):
            ref = attr(c, "r")
            if ref:
                ref = ref.upper()
                _, col_no = split_a1(ref)
            else:
                col_no = next_col
                ref = a1(row_no, col_no)
            next_col = col_no + 1
            cell: dict = {}
            kind = attr(c, "t", "n")
            v = kid(c, "v")
            raw = text_of(v) if v is not None else None
            if raw is not None and not raw.strip() and kind not in ("s", "str", "inlineStr"):
                raw = None  # some writers (openpyxl) emit an empty <v/>: no stored value
            try:
                if kind == "inlineStr":
                    is_el = kid(c, "is")
                    cell["v"] = _rich_text(is_el) if is_el is not None else ""
                elif raw is None:
                    pass
                elif kind == "s":
                    cell["v"] = strings[int(raw)]
                elif kind == "str":
                    cell["v"] = raw
                elif kind == "b":
                    cell["v"] = raw.strip() == "1"
                elif kind == "e":
                    cell["e"] = raw
                elif kind == "n":
                    cell["v"] = float(raw)
                else:
                    cell["v"] = raw
                    notes.append(f"{ref}: cell type {kind!r} kept as text")
            except (ValueError, IndexError):
                cell["unreadable"] = f"stored value {raw!r} (type {kind!r}) could not be read"
                notes.append(f"{ref}: {cell['unreadable']}")

            f = kid(c, "f")
            if f is not None:
                f_type = attr(f, "t", "normal")
                f_text = text_of(f)
                if f_type == "shared":
                    si = attr(f, "si")
                    if f_text:
                        shared[si] = (f_text, row_no, col_no)
                    elif si in shared:
                        base, base_row, base_col = shared[si]
                        try:
                            f_text = shift_formula(base, row_no - base_row, col_no - base_col)
                        except FormulaError as exc:
                            cell["unsupported"] = f"shared formula could not be expanded: {exc}"
                            f_text = base
                    else:
                        cell["unsupported"] = "shared formula has no master cell"
                elif f_type == "array":
                    cell["unsupported"] = "array formula (Ctrl+Shift+Enter)"
                elif f_type == "dataTable":
                    cell["unsupported"] = "what-if data table"
                if attr(c, "cm"):
                    cell["unsupported"] = "dynamic array formula"
                if f_type in ("array", "dataTable") and attr(f, "ref"):
                    blocks.append((attr(f, "ref").upper(), ref, f_text, cell["unsupported"]))
                cell["f"] = f_text
            if cell:
                cells[ref] = cell

    # Every cell an array formula or data table fills is a result of that formula,
    # not an input, even though the file stores no formula for it.
    for block, master, f_text, reason in blocks:
        first, _, last = block.partition(":")
        try:
            r1, c1 = split_a1(first)
            r2, c2 = split_a1(last or first)
        except FormulaError:
            notes.append(f"{master}: unreadable formula range {block!r}")
            continue
        if (r2 - r1 + 1) * (c2 - c1 + 1) > 1_000_000:
            notes.append(f"{master}: formula range {block} is too large to model")
            continue
        for r in range(r1, r2 + 1):
            for col in range(c1, c2 + 1):
                member = a1(r, col)
                if member != master:
                    cell = cells.setdefault(member, {})
                    cell["f"] = f_text
                    cell["unsupported"] = f"part of the {reason} at {master}"
    return cells, notes


def read_workbook(pkg: dict) -> dict:
    """Build the workbook model from a canonical package."""
    wb_name = _workbook_part(pkg)
    wb_part = find_part(pkg, wb_name)
    if not wb_part or wb_part["kind"] != "xml":
        raise ConversionError("workbook part is missing or unreadable")
    wb = root(wb_part["xml"])
    wb_rels = _rels(pkg, wb_name)
    strings = _shared_strings(pkg, wb_rels)

    pr = kid(wb, "workbookPr")
    date1904 = bool(pr) and attr(pr, "date1904", "0") in ("1", "true")

    sheets = []
    sheets_el = kid(wb, "sheets")
    for s in kids(sheets_el, "sheet") if sheets_el else ():
        target = wb_rels.get(attr(s, "id"))
        part = find_part(pkg, target) if target else None
        sheet = {"name": attr(s, "name"), "state": attr(s, "state", "visible"), "part": target, "cells": {}, "notes": []}
        if part is None or part["kind"] != "xml" or kid(root(part["xml"]), "sheetData") is None:
            sheet["notes"].append("not a worksheet (chart sheet, macro sheet, missing part or kept as bytes)")
        else:
            try:
                sheet["cells"], sheet["notes"] = _read_sheet(part["xml"], strings)
            except FormulaError as exc:
                raise ConversionError(f"sheet {sheet['name']!r}: {exc}") from None
        rows = [split_a1(r) for r in sheet["cells"]]
        sheet["max_row"] = max((r for r, _ in rows), default=0)
        sheet["max_col"] = max((c for _, c in rows), default=0)
        sheets.append(sheet)

    names = []
    names_el = kid(wb, "definedNames")
    for n in kids(names_el, "definedName") if names_el else ():
        local_id = attr(n, "localSheetId")
        names.append({
            "name": attr(n, "name"),
            "scope": sheets[int(local_id)]["name"] if local_id is not None and int(local_id) < len(sheets) else None,
            "refers_to": text_of(n),
        })

    text = sum(len(c["v"]) for s in sheets for c in s["cells"].values() if isinstance(c.get("v"), str))
    if text > MAX_MODEL_TEXT:
        raise ConversionError(f"cells hold {text:,} characters of text once shared strings are expanded "
                              f"(limit {MAX_MODEL_TEXT:,}); the readable views are not produced")

    macro_parts = [p["name"] for p in pkg["parts"] if posixpath.basename(p["name"]).lower() == "vbaproject.bin"]
    external = [p["name"] for p in pkg["parts"] if "/externalLinks/" in p["name"] and "_rels" not in p["name"]]
    app_part = find_part(pkg, "docProps/app.xml")
    application = None
    if app_part and app_part["kind"] == "xml":
        app_el = kid(root(app_part["xml"]), "Application")
        version = kid(root(app_part["xml"]), "AppVersion")
        if app_el is not None:
            application = text_of(app_el) + (f" {text_of(version)}" if version is not None else "")
    return {
        "format": FORMAT,
        "application": application,
        "date1904": date1904,
        "sheets": sheets,
        "defined_names": names,
        "macro_parts": macro_parts,
        "external_link_parts": external,
    }


def key(sheet: str, ref: str) -> str:
    return f"{sheet}!{ref}"


def formula_cells(model: dict):
    """Yield (key, sheet, ref, cell) for every formula cell, sheet by sheet, row-major."""
    for sheet in model["sheets"]:
        for ref in sorted((r for r, c in sheet["cells"].items() if "f" in c), key=split_a1):
            yield key(sheet["name"], ref), sheet, ref, sheet["cells"][ref]


# -------------------------------------------------------------------- Markdown

def _show(cell: dict | None) -> str:
    if not cell:
        return ""
    if "e" in cell:
        return cell["e"]
    v = cell.get("v")
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() and abs(v) < 1e15 else repr(v)
    return str(v).replace("|", "\\|").replace("\n", " ")


def to_markdown(model: dict, source_name: str, source_sha256: str, max_rows: int = 200, max_cols: int = 30) -> str:
    out = [
        f"# {source_name}",
        "",
        f"Derived view of the workbook with SHA-256 `{source_sha256}`. "
        "Values are Excel's stored results; this view omits styles and layout and is not the source of truth.",
    ]
    for sheet in model["sheets"]:
        out += ["", f"## {sheet['name']}", ""]
        if not sheet["cells"]:
            out.append("_(empty)_")
            continue
        rows, cols = min(sheet["max_row"], max_rows), min(sheet["max_col"], max_cols)
        out.append("| | " + " | ".join(col_letters(c) for c in range(1, cols + 1)) + " |")
        out.append("|---" * (cols + 1) + "|")
        for r in range(1, rows + 1):
            out.append(f"| **{r}** | " + " | ".join(_show(sheet["cells"].get(a1(r, c))) for c in range(1, cols + 1)) + " |")
        if rows < sheet["max_row"] or cols < sheet["max_col"]:
            out += ["", f"_Truncated to {rows} rows x {cols} columns of {sheet['max_row']} x {sheet['max_col']}._"]
        formulas = [(ref, c["f"]) for ref, c in sorted(sheet["cells"].items(), key=lambda kv: split_a1(kv[0])) if "f" in c]
        if formulas:
            out += ["", "### Formulas", ""]
            out += [f"- `{ref}` = `{f}`" for ref, f in formulas[:500]]
            if len(formulas) > 500:
                out.append(f"- ... and {len(formulas) - 500} more")
    return "\n".join(out) + "\n"
