import json

import pytest

from provenance import certificate, diff
from provenance.canon import build_package, dumps, read_package, tree_to_xml, xml_to_tree
from provenance.cli import check, convert, main
from provenance.compiler import compile_workbook, load_module
from provenance.formula import FormulaError, parse, shift_formula
from provenance.model import read_workbook
from provenance.runtime import DIV0, NA, NUM, VALUE, XlError
from provenance.verify import round_trip


# ------------------------------------------------------------------ structure

def test_xml_tree_keeps_everything():
    raw = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           b'<a:root xmlns:a="urn:a" xmlns:unused="urn:u" z="1" b="2"><!-- note --><?pi data?>'
           b'<a:t xml:space="preserve">  x &amp; y&#13;\n</a:t><empty a:q="&quot;&#10;"/>tail</a:root>')
    tree = xml_to_tree(raw)
    again = xml_to_tree(tree_to_xml(tree))
    assert again == tree
    assert tree["nodes"][0]["attrs"][:2] == [["xmlns:a", "urn:a"], ["xmlns:unused", "urn:u"]]


def test_round_trip_is_proven_and_deterministic(budget):
    pkg = read_package(budget)
    report = round_trip(budget, pkg)
    assert report["status"] == "proven", report["problems"]
    assert all(p["match"] for p in report["parts"])
    assert build_package(pkg) == build_package(json.loads(dumps(pkg)))


def test_round_trip_detects_a_lossy_package(budget):
    pkg = read_package(budget)
    sheet = next(p for p in pkg["parts"] if p["name"].endswith("sheet1.xml"))
    broken = json.loads(dumps(pkg))
    target = next(p for p in broken["parts"] if p["name"] == sheet["name"])
    target["xml"]["nodes"][0]["children"][0]["children"].pop()  # lose a row
    assert round_trip(budget, broken)["status"] == "failed"


def test_binary_parts_survive_byte_for_byte(xlsx):
    blob = bytes(range(256)) * 4
    data = xlsx({"S": {"A1": 1}}, extra={"xl/vbaProject.bin": blob})
    pkg = read_package(data)
    assert round_trip(data, pkg)["status"] == "proven"
    assert read_workbook(pkg)["macro_parts"] == ["xl/vbaProject.bin"]


# -------------------------------------------------------------------- formulas

@pytest.mark.parametrize("text, expected", [
    ("-2^2", 4.0), ("2^3^2", 64.0), ("1+2*3", 7.0), ("(1+2)*3", 9.0), ("50%*4", 2.0), ("2^-1", 0.5),
    ('"a"&1+1', "a2"), ("1=1", True), ('"a"<"B"', True), ('1<"a"', True), ("10/4", 2.5),
    ("1/0", DIV0), ('"x"+1', VALUE), ('"3"+1', 4.0), ("TRUE+1", 2.0),
    ("SUM(1,2,3)", 6.0), ('SUM("5",TRUE)', 6.0), ("ROUND(2.675,2)", 2.68), ("ROUND(-0.5,0)", -1.0),
    ("ROUND(1234.5,-2)", 1200.0), ("MOD(-3,2)", 1.0), ("INT(-1.5)", -2.0),
    ('IF(1>2,"a","b")', "b"), ("IF(0,1)", False), ("AND(TRUE,1,2)", True), ("OR(FALSE,0)", False),
    ("IFERROR(1/0,7)", 7.0), ("ISNA(NA())", True), ("MATCH(3,{1,2,3,4},0)", 3.0), ("MATCH(2.5,{1,2,3},1)", 2.0),
    ('VLOOKUP("b",{"a",1;"b",2},2,FALSE)', 2.0), ('VLOOKUP("z",{"a",1;"b",2},2,FALSE)', NA),
    ("INDEX({1,2;3,4},2,1)", 3.0), ("SUMPRODUCT({1,2,3},{4,5,6})", 32.0),
    ('LEFT("hello",2)&RIGHT("hello",2)&MID("hello",2,3)', "heloell"), ('LEN(TRIM("  a   b "))', 3.0),
    ("0.1+0.2", 0.30000000000000004), ('1/3&""', "0.333333333333333"), ('1E+20&""', "1E+20"),
    ("YEAR(DATE(2024,2,29))", 2024.0), ("DATE(2024,1,1)", 45292.0), ("DAY(45351)", 29.0),
    ("PMT(0.05/12,60,10000)", -188.71233644010877), ('COUNTIF({1,2,3,4},">2")', 2.0),
    ('SUMIF({"a","b","a"},"a",{1,2,3})', 4.0), ("_xlfn.IFNA(NA(),5)", 5.0),
    # found by the differential sweep (scripts/sweep.py)
    ("VALUE(TRUE)", VALUE), ('"1,000"+0', 1000.0), ('"$5"+1', 6.0), ('"50%"*2', 1.0), ('VALUE(" -$1,234.5 ")', -1234.5),
    ('"abc"+0', VALUE), ("ROUND(1.005,2)", 1.01), ("ROUND(0.285,2)", 0.29), ("0^0", NUM), ('LEN(1/3)', 17.0),
    ('2^53&""', "9.00719925474099E+15"), ('TRIM("  a  b  ")', "a b"),
])
def test_formula_semantics(xlsx, text, expected):
    model = read_workbook(read_package(xlsx({"S": {"A1": (text, None)}})))
    source, report = compile_workbook(model, "0" * 64)
    assert not report["unsupported"], report["unsupported"]
    values, failed = load_module(source)["compute"]()
    assert not failed
    got = values["S!A1"]
    if isinstance(expected, float):
        assert got == pytest.approx(expected, rel=1e-12)
    else:
        assert got == expected and type(got) is type(expected)


def test_parser_rejects_what_it_does_not_model():
    for text in ("Table1[Amount]", "[1]Sheet1!A1", "Sheet1:Sheet3!A1", "A1:A3 B1:B3"):
        with pytest.raises(FormulaError):
            parse(text)


def test_shared_formula_shift():
    assert shift_formula("A1+$B$1+C$2+SUM($A1:B2)", 2, 1) == "B3+$B$1+D$2+SUM($A3:C4)"
    assert shift_formula("A1", -1, 0) == "#REF!"


def test_ranges_ignore_text_but_direct_arguments_do_not(xlsx):
    model = read_workbook(read_package(xlsx({"S": {
        "A1": "5", "A2": 3, "A3": True,
        "B1": ("SUM(A1:A3)", 3), "B2": ("SUM(A1,A2)", 3), "B3": ("A1+A2", 8), "B4": ("COUNT(A1:A3)", 1),
    }})))
    source, _ = compile_workbook(model, "0" * 64)
    assert diff.replay(model, load_module(source))["counts"]["exact"] == 4


def test_unsupported_cells_are_reported_not_guessed(xlsx):
    model = read_workbook(read_package(xlsx({"S": {
        "A1": 1, "A2": ("NOW()", 45000.5), "A3": ("A2+1", 45001.5), "A4": ("A1+1", 2), "A5": ('TEXT(A1,"0.00")', "1.00"),
    }})))
    source, report = compile_workbook(model, "0" * 64)
    assert set(report["unsupported"]) == {"S!A2", "S!A5"}
    values, failed = load_module(source)["compute"]()
    assert values == {"S!A4": 2.0}
    assert "S!A3" in failed and "S!A2" in failed["S!A3"]


def test_text_excel_may_read_as_a_date_is_not_guessed(xlsx):
    model = read_workbook(read_package(xlsx({"S": {
        "A1": ('"1/1/2024"+0', 45292), "A2": ('SUMIF(B1:B2,">=1/1/2024",C1:C2)', 1), "A3": ('COUNTIF(B1:B2,"-")', 0),
        "B1": 45292, "B2": 1, "C1": 1, "C2": 2,
    }})))
    values, failed = load_module(compile_workbook(model, "0" * 64)[0])["compute"]()
    assert set(failed) == {"S!A1", "S!A2"}
    assert values["S!A3"] == 0.0


def test_countif_blank_and_error_criteria(xlsx):
    model = read_workbook(read_package(xlsx({"S": {
        "A1": "", "A2": 0, "A3": "#N/A", "A4": "x",  # A5 is truly empty
        "B1": ('COUNTIF(A1:A5,"")', 2), "B2": ('COUNTIF(A1:A5,"=")', 1), "B3": ('COUNTIF(A1:A5,"<>")', 4),
        "B4": ('COUNTIF(A1:A5,"#N/A")', 1), "B5": ('COUNTIF(A1:A5,"<>#N/A")', 4), "B6": ("COUNTIF(A1:A5,C1)", 1),
    }})))
    source, _ = compile_workbook(model, "0" * 64)
    assert diff.replay(model, load_module(source))["counts"]["exact"] == 6


def test_circular_reference_is_unverified(xlsx):
    model = read_workbook(read_package(xlsx({"S": {"A1": ("A2+1", 0), "A2": ("A1+1", 0)}})))
    values, failed = load_module(compile_workbook(model, "0" * 64)[0])["compute"]()
    assert not values and set(failed) == {"S!A1", "S!A2"}


def test_defined_names_and_other_sheets(xlsx):
    data = xlsx({"Main": {"A1": ("Rate*'Other Sheet'!A1", 20), "A2": ("SUM(Vals)", 6)},
                 "Other Sheet": {"A1": 100, "B1": 1, "B2": 2, "B3": 3}},
                names={"Rate": "0.2", "Vals": "'Other Sheet'!$B$1:$B$3"})
    model = read_workbook(read_package(data))
    source, report = compile_workbook(model, "0" * 64)
    assert not report["unsupported"]
    assert diff.replay(model, load_module(source))["counts"]["exact"] == 2


# ------------------------------------------------------------------ end to end

def test_convert_budget_end_to_end(tmp_path, budget):
    src = tmp_path / "budget.xlsx"
    src.write_bytes(budget)
    cert = convert(src, tmp_path / "out", fuzz_cases=0)
    assert cert["layers"]["structure"]["status"] == "proven"
    formulas = cert["layers"]["formulas"]
    assert formulas["status"] == "verified", formulas
    assert formulas["replay"]["counts"]["mismatch"] == 0
    assert formulas["counts"] == {"reproduced": 11}
    assert cert["verdict"]["level"] == "verified"
    assert check(tmp_path / "out" / "budget.certificate.json", src) == []

    rebuilt = tmp_path / "rebuilt.xlsx"
    assert main(["rebuild", str(tmp_path / "out" / "budget.package.json"), "-o", str(rebuilt)]) == 0
    assert round_trip(budget, read_package(rebuilt.read_bytes()))["status"] == "proven"


def test_wrong_stored_value_is_a_mismatch(tmp_path, xlsx):
    src = tmp_path / "bad.xlsx"
    src.write_bytes(xlsx({"S": {"A1": 2, "A2": ("A1*2", 5)}}))
    cert = convert(src, tmp_path / "out", fuzz_cases=0)
    assert cert["verdict"]["level"] == "failed"
    assert cert["layers"]["formulas"]["replay"]["mismatches"][0]["cell"] == "S!A2"


def test_tampered_certificate_or_package_is_caught(tmp_path, budget):
    src = tmp_path / "budget.xlsx"
    src.write_bytes(budget)
    convert(src, tmp_path / "out", fuzz_cases=0)
    cert_path = tmp_path / "out" / "budget.certificate.json"
    cert = json.loads(cert_path.read_text())
    cert["verdict"]["level"] = "verified!"
    cert_path.write_text(json.dumps(cert))
    assert any("altered" in p for p in check(cert_path, None))

    pkg_path = tmp_path / "out" / "budget.package.json"
    pkg_path.write_bytes(pkg_path.read_bytes().replace(b"Widget", b"Wodget"))
    assert any("package JSON" in p for p in check(cert_path, None))


def test_macros_make_the_verdict_partial(tmp_path, xlsx):
    src = tmp_path / "m.xlsm"
    src.write_bytes(xlsx({"S": {"A1": 1, "A2": ("A1+1", 2)}}, extra={"xl/vbaProject.bin": b"\x00\x01binary"}))
    cert = convert(src, tmp_path / "out", fuzz_cases=0)
    assert cert["layers"]["macros"]["status"] == "unverified"
    assert cert["verdict"]["level"] == "partial"


def test_fuzz_against_formulas_library(tmp_path, budget):
    pytest.importorskip("formulas")
    src = tmp_path / "budget.xlsx"
    src.write_bytes(budget)
    cert = convert(src, tmp_path / "out", fuzz_cases=10, seed=1, engine="formulas")
    fz = cert["layers"]["formulas"]["fuzz"]
    assert fz["status"] == "ran", fz
    assert fz["counts"]["mismatch"] == 0, fz["mismatches"][:3]
    assert fz["counts"]["exact"] + fz["counts"]["close"] > 50


def test_seal_detects_change():
    cert = certificate.seal({"a": 1})
    assert certificate.check_seal(cert) == []
    cert["a"] = 2
    assert certificate.check_seal(cert)


def test_empty_and_unreadable_stored_values(xlsx):
    data = xlsx({"S": {"A1": 1, "A2": ("A1+1", 2), "A3": ("A1*5", 5)}})
    import io, zipfile
    src = zipfile.ZipFile(io.BytesIO(data))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as z:
        for info in src.infolist():
            raw = src.read(info)
            if info.filename.endswith("sheet1.xml"):
                raw = raw.replace(b"<v>2</v>", b"<v></v>").replace(b'<c r="A1"><v>1</v>', b'<c r="A1"><v>one</v>')
            z.writestr(info, raw)
    model = read_workbook(read_package(out.getvalue()))
    assert "unreadable" in model["sheets"][0]["cells"]["A1"]
    source, report = compile_workbook(model, "0" * 64)
    assert report["unreadable_inputs"] and not report["unsupported"]
    values, failed = load_module(source)["compute"]()
    assert set(failed) >= {"S!A2", "S!A3"} and not values


# ------------------------------------------------- found by independent review

def test_deep_chain_where_early_cells_read_later_ones(xlsx):
    n = 30_000
    cells = {f"A{i}": (f"A{i + 1}+1", n - i) for i in range(1, n)}
    cells[f"A{n}"] = 0
    model = read_workbook(read_package(xlsx({"S": cells})))
    source, _ = compile_workbook(model, "0" * 64)
    import sys
    old = sys.getrecursionlimit()
    sys.setrecursionlimit(1000)  # the evaluator must not need deep recursion
    try:
        values, failed = load_module(source)["compute"]()
    finally:
        sys.setrecursionlimit(old)
    assert not failed and values["S!A1"] == float(n - 1)


def test_huge_and_whole_column_ranges(xlsx):
    model = read_workbook(read_package(xlsx({
        "S": {"A1": 1, "A2": 2, "B1": "x", "B2": "",
              "C2": ("SUM(A:A)", 3), "C3": ("COUNTBLANK(A:A)", 1048574),
              "C4": ('COUNTIF(B:B,"x")', 1), "C5": ('COUNTIF(B:B,"")', 1048575), "C6": ("SUMPRODUCT(A:A,Other!A:A)", 23),
              "C7": ("INDEX(A:A,2)", 2), "C8": ('SUMPRODUCT((B:B="x")*A:A)', 1), "C9": ('SUMPRODUCT((B:B="")*1)', 1048575)},
        "Other": {"A1": 3, "A2": 10, "B1": ("COUNTA(S!A1:XFD2)", 5), "B2": ("COUNTA(S!A1:B1048576)", 4)},
    })))
    source, report = compile_workbook(model, "0" * 64)
    values, failed = load_module(source)["compute"]()
    assert {k: values[k] for k in ("Other!B1", "Other!B2", "S!C2", "S!C4", "S!C6", "S!C7", "S!C8")} == {
        "Other!B1": 5.0, "Other!B2": 4.0, "S!C2": 3.0, "S!C4": 1.0, "S!C6": 23.0, "S!C7": 2.0, "S!C8": 1.0}
    assert set(failed) == {"S!C3", "S!C5", "S!C9"}  # depend on the unseen blank cells: declined


def test_text_ordering_with_punctuation_is_not_guessed(xlsx):
    model = read_workbook(read_package(xlsx({"S": {
        "A1": ('"a"<"b"', True), "A2": ('"_"<"a"', True), "A3": ('"co-op"<"coop"', False), "A4": ('"ab"="AB"', True),
    }})))
    values, failed = load_module(compile_workbook(model, "0" * 64)[0])["compute"]()
    assert values == {"S!A1": True, "S!A4": True} and set(failed) == {"S!A2", "S!A3"}


def test_round_of_huge_number(xlsx):
    model = read_workbook(read_package(xlsx({"S": {"A1": ("ROUND(1E+300,2)", 1e300)}})))
    values, failed = load_module(compile_workbook(model, "0" * 64)[0])["compute"]()
    assert values == {"S!A1": 1e300}


def test_fuzz_varies_text_and_booleans(xlsx):
    model = read_workbook(read_package(xlsx({"S": {
        "A1": "yes", "A2": "no", "A3": "maybe", "B1": True, "C1": 5, "D1": ('IF(A1="yes",C1,0)', 5)}})))
    cases = diff.make_cases(model, 40, 0)
    changed = {k for case in cases for k in case}
    assert {"S!A1", "S!B1", "S!C1"} <= changed
    assert all(isinstance(v, str) for case in cases for k, v in case.items() if k.startswith("S!A"))


def test_fuzz_with_text_inputs_against_formulas_library(tmp_path, xlsx):
    pytest.importorskip("formulas")
    src = tmp_path / "t.xlsx"
    src.write_bytes(xlsx({"S": {
        "A1": "yes", "A2": "no", "B1": True, "C1": 5, "C2": 7,
        "D1": ('IF(A1="yes",C1,C2)', 5), "D2": ('IF(B1,C1*2,0)', 10), "D3": ('A1&"-"&A2', "yes-no")}}))
    cert = convert(src, tmp_path / "out", fuzz_cases=20, seed=3, engine="formulas")
    fz = cert["layers"]["formulas"]["fuzz"]
    assert fz["status"] == "ran" and fz["counts"]["mismatch"] == 0, fz
    assert cert["verdict"]["level"] == "verified"


def test_unsigned_certificate_is_labelled(tmp_path, budget, capsys):
    src = tmp_path / "budget.xlsx"
    src.write_bytes(budget)
    convert(src, tmp_path / "out", fuzz_cases=0)
    assert main(["check", str(tmp_path / "out" / "budget.certificate.json")]) == 0
    assert "not signed" in capsys.readouterr().out


# ------------------------------------- found by the public corpus (Excel-saved answers)

@pytest.mark.parametrize("text, expected", [
    ('MATCH(4.1,{"c bau c",4.1,"a",1.1})', 2.0), ('MATCH(4.1,{TRUE,4.1,FALSE,1.1})', 2.0),
    ("MATCH(1.1,{-1.1,2.1,3.1,4.1},-1)", NA), ("MATCH(0,{0,1,2,3,4},-1)", 1.0), ("MATCH(0.5,{0,1,2,3,4},-1)", NA),
    ("MATCH(5,{4.1,3.1,2.1,1.1})", 4.0), ("MATCH(TRUE,{TRUE,4.1,FALSE,1.1})", 3.0),
])
def test_excel_lookup_and_logic_quirks(xlsx, text, expected):
    model = read_workbook(read_package(xlsx({"S": {"A1": (text, None)}})))
    values, failed = load_module(compile_workbook(model, "0" * 64)[0])["compute"]()
    assert not failed and values["S!A1"] == expected


def test_sumproduct_single_cell_text_and_external_links(xlsx):
    model = read_workbook(read_package(xlsx({"S": {
        "A1": "4", "B1": 2, "C1": ("SUMPRODUCT(A1:A1,B1:B1)", "#VALUE!"),
        "C2": ("SUM('[other.xlsx]Sheet0'!A1:B1)", 30)}})))
    source, report = compile_workbook(model, "0" * 64)
    assert "another workbook" in report["unsupported"]["S!C2"]
    assert diff.replay(model, load_module(source))["counts"]["exact"] == 1


def test_long_text_has_one_canonical_form(xlsx):
    data = xlsx({"S": {"A1": "x" * 50_000 + "&<" + "y" * 50_000}})
    pkg = read_package(data)
    assert round_trip(data, pkg)["status"] == "proven"
    assert read_package(build_package(pkg)) == pkg


def test_hostile_packages_are_rejected_or_kept_as_bytes(xlsx):
    import io, zipfile
    from provenance.canon import ConversionError
    deep = b"<a>" * 5000 + b"</a>" * 5000
    data = xlsx({"S": {"A1": 1}}, extra={"customXml/deep.xml": deep})
    pkg = read_package(data)
    assert next(p for p in pkg["parts"] if p["name"] == "customXml/deep.xml")["kind"] == "binary"
    assert round_trip(data, pkg)["status"] == "proven"
    corrupt = bytearray(data)
    i = corrupt.find(b"xl/workbook.xml") + 40  # inside that part's compressed data
    corrupt[i] ^= 0xFF
    with pytest.raises(ConversionError):
        read_package(bytes(corrupt))


def test_shared_string_amplification_is_refused(tmp_path, xlsx):
    from provenance import model as model_mod
    big = "z" * 1000
    data = xlsx({"S": {f"A{i}": big for i in range(1, 300)}})
    old = model_mod.MAX_MODEL_TEXT
    model_mod.MAX_MODEL_TEXT = 100_000
    try:
        src = tmp_path / "amp.xlsx"
        src.write_bytes(data)
        cert = convert(src, tmp_path / "out", fuzz_cases=0)
    finally:
        model_mod.MAX_MODEL_TEXT = old
    assert cert["layers"]["structure"]["status"] == "proven" and "formulas" not in cert["layers"]
    assert any("characters of text" in n for n in cert["notes"])


def test_each_formula_judged_on_stored_inputs_so_rounding_does_not_cascade(xlsx):
    # A1 is stored 1 ULP away from what we compute; B1 subtracts nearly equal numbers.
    a = 41922.06155493237
    model = read_workbook(read_package(xlsx({"S": {
        "X1": 41922.0, "X2": 0.06155493237,
        "A1": ("X1+X2", a), "C1": 41922.06, "B1": ("C1-A1", 41922.06 - a)}})))
    source, _ = compile_workbook(model, "0" * 64)
    report = diff.replay(model, load_module(source))
    assert report["counts"]["mismatch"] == 0


def test_libreoffice_booleans_stored_as_numbers(xlsx):
    data = xlsx({"S": {"A1": 1, "B1": ("ISNUMBER(A1)", 1)}},
                extra={"docProps/app.xml": '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>LibreOffice/7.0</Application></Properties>'})
    model = read_workbook(read_package(data))
    assert model["application"] == "LibreOffice/7.0"
    report = diff.replay(model, load_module(compile_workbook(model, "0" * 64)[0]))
    assert report["counts"]["close"] == 1 and report["stored_by"] == "LibreOffice/7.0"


def test_quirks_from_the_second_corpus_pass(xlsx):
    model = read_workbook(read_package(xlsx({"S": {
        "A1": ('LEFT("ciao",H1)', ""), "A2": ('MID("ciao",2,H1)', ""), "A3": ('FIND("ia","ciao",H1)', "#VALUE!"),
        "A4": ('LEFT("ciao")', "c"), "A5": ('MATCH("\u00df",{1,"ss",2.1},0)', "#N/A"), "A6": ('MATCH("SS",{1,"ss"},0)', 2),
        "A7": ("SUMPRODUCT(B1:C1,B2:D2)", "#NUM!"), "B1": "#NUM!", "C1": 1, "B2": 1, "C2": 2, "D2": 3,
        "A8": ('AND(TRUE,"0")', True)}})))
    source, report = compile_workbook(model, "0" * 64)
    result = diff.replay(model, load_module(source))
    assert result["counts"]["mismatch"] == 0 and result["counts"]["exact"] == 7
    assert result["cells"]["S!A8"] == "unverified"  # Excel-saved files disagree on text in AND/OR


def test_cells_on_or_after_a_cycle_are_not_checked(xlsx):
    # B1 -> C1 -> D1 (an array formula we do not compile) -> B1; E1 reads the cycle; F1 does not.
    import io, zipfile
    data = xlsx({"S": {"A1": True, "B1": ("C1", 0), "C1": ("D1", 0), "D1": ("IF(A1:A2,B1:B2,1)", 0),
                       "E1": ("B1+1", 0), "F1": ("A1", True)}})
    src = zipfile.ZipFile(io.BytesIO(data))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as z:
        for info in src.infolist():
            raw = src.read(info)
            if info.filename.endswith("sheet1.xml"):
                raw = raw.replace(b"<f>IF(A1:A2,B1:B2,1)</f>", b'<f t="array" ref="D1">IF(A1:A2,B1:B2,1)</f>')
            z.writestr(info, raw)
    model = read_workbook(read_package(out.getvalue()))
    result = diff.replay(model, load_module(compile_workbook(model, "0" * 64)[0]))
    assert result["counts"]["mismatch"] == 0
    assert {k for k, v in result["cells"].items() if v == "unverified"} == {"S!B1", "S!C1", "S!D1", "S!E1"}
    assert result["cells"]["S!F1"] == "exact"


def test_array_formula_results_are_not_treated_as_inputs(xlsx):
    import io, zipfile
    data = xlsx({"S": {"A1": 1, "A2": 2, "B1": ("A1:A2*10", 10), "B2": 20, "C1": ("B2+1", 21)}})
    src = zipfile.ZipFile(io.BytesIO(data))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as z:
        for info in src.infolist():
            raw = src.read(info)
            if info.filename.endswith("sheet1.xml"):
                raw = raw.replace(b"<f>A1:A2*10</f>", b'<f t="array" ref="B1:B2">A1:A2*10</f>')
            z.writestr(info, raw)
    model = read_workbook(read_package(out.getvalue()))
    assert "part of the array formula" in model["sheets"][0]["cells"]["B2"]["unsupported"]
    source, report = compile_workbook(model, "0" * 64)
    values, failed = load_module(source)["compute"]({"S!A2": 5})
    assert "S!C1" in failed and "S!B2" not in load_module(source)["INPUTS"]  # never a stale constant


def test_found_by_random_inputs_across_the_corpus(xlsx):
    model = read_workbook(read_package(xlsx({"S": {
        "A1": True, "A2": False, "A3": 1, "B1": 5, "B2": 7, "B3": 9,
        "C1": ('COUNTIF(A1:A3,"<=TRUE")', 2), "C2": ('SUMIF(A1:A3,">FALSE",B1:B3)', 5), "C3": ('COUNTIF(A1:A3,"<TRUE")', 1),
        "C4": ("PV(0.1,6,0,-1000,3)", 0), "C5": ("DATE(10000,1,-1)", "#NUM!")}})))
    source, _ = compile_workbook(model, "0" * 64)
    result = diff.replay(model, load_module(source))
    assert result["counts"]["mismatch"] == 0 and result["counts"]["exact"] == 4
    assert result["cells"]["S!C4"] == "unverified"
