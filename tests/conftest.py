"""Builds small .xlsx files by hand so stored formula results are under test control."""

import io
import zipfile
from xml.sax.saxutils import escape

import pytest

NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _cell(ref, spec):
    formula, value = spec if isinstance(spec, tuple) else (None, spec)
    f = "" if formula is None else f"<f>{escape(formula)}</f>"
    if value is None:
        return f'<c r="{ref}">{f}</c>'
    if isinstance(value, bool):
        return f'<c r="{ref}" t="b">{f}<v>{int(value)}</v></c>'
    if isinstance(value, (int, float)):
        return f'<c r="{ref}">{f}<v>{value!r}</v></c>'
    if value.startswith("#"):
        return f'<c r="{ref}" t="e">{f}<v>{escape(value)}</v></c>'
    if formula is None:
        return f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{escape(value)}</t></is></c>'
    return f'<c r="{ref}" t="str">{f}<v>{escape(value)}</v></c>'


def _split(ref):
    letters = "".join(ch for ch in ref if ch.isalpha())
    col = 0
    for ch in letters:
        col = col * 26 + ord(ch) - 64
    return int(ref[len(letters):]), col


def make_xlsx(sheets, names=None, extra=None):
    """sheets: {name: {ref: value | (formula, stored_result)}} -> bytes of an .xlsx"""
    parts = {}
    overrides, sheet_els, rels = [], [], []
    for i, (name, cells) in enumerate(sheets.items(), 1):
        rows = {}
        for ref in sorted(cells, key=_split):
            rows.setdefault(_split(ref)[0], []).append(_cell(ref, cells[ref]))
        body = "".join(f'<row r="{r}">{"".join(cs)}</row>' for r, cs in sorted(rows.items()))
        parts[f"xl/worksheets/sheet{i}.xml"] = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet {NS}><sheetData>{body}</sheetData></worksheet>'
        overrides.append(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
        sheet_els.append(f'<sheet name="{escape(name)}" sheetId="{i}" r:id="rId{i}"/>')
        rels.append(f'<Relationship Id="rId{i}" Type="{REL}/worksheet" Target="worksheets/sheet{i}.xml"/>')
    defined = "".join(f'<definedName name="{n}">{escape(t)}</definedName>' for n, t in (names or {}).items())
    defined = f"<definedNames>{defined}</definedNames>" if defined else ""
    parts["xl/workbook.xml"] = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook {NS} xmlns:r="{REL}"><sheets>{"".join(sheet_els)}</sheets>{defined}</workbook>'
    parts["xl/_rels/workbook.xml.rels"] = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{"".join(rels)}</Relationships>'
    parts["_rels/.rels"] = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    content_types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                     '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
                     '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>'
                     '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                     + "".join(overrides) + "</Types>")
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        for name, text in parts.items():
            z.writestr(name, text)
        for name, data in (extra or {}).items():
            z.writestr(name, data)
    return buf.getvalue()


@pytest.fixture
def xlsx():
    return make_xlsx


@pytest.fixture
def budget():
    """A small realistic workbook whose stored results are what Excel shows."""
    return make_xlsx({
        "Budget": {
            "A1": "Item", "B1": "Qty", "C1": "Price", "D1": "Total",
            "A2": "Widget", "B2": 4, "C2": 2.5, "D2": ("B2*C2", 10),
            "A3": "Gadget", "B3": 10, "C3": 1.2, "D3": ("B3*C3", 12),
            "A4": "Gizmo", "B4": 0, "C4": 99.99, "D4": ("B4*C4", 0),
            "A6": "Sum", "D6": ("SUM(D2:D4)", 22),
            "A7": "Average price", "D7": ("AVERAGE(C2:C4)", 34.56333333333333),
            "A8": "Big order?", "D8": ('IF(D6>20,"yes","no")', "yes"),
            "A9": "Per unit", "D9": ("D6/SUM(B2:B4)", 1.5714285714285714),
            "A10": "Gadget price", "D10": ('VLOOKUP("Gadget",A2:C4,3,FALSE)', 1.2),
            "A11": "Label", "D11": ('A2&" x"&B2', "Widget x4"),
            "A12": "Tax", "D12": ("ROUND(D6*Rates!B1,2)", 4.4),
            "A13": "Safe", "D13": ("IFERROR(D6/B4,-1)", -1),
        },
        "Rates": {"A1": "VAT", "B1": 0.2},
    })
