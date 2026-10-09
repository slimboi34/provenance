"""Differential sweep: every supported function x awkward inputs, compared with
the independent `formulas` interpreter.

    .venv/bin/python scripts/sweep.py [--show N]

A disagreement means one of the two engines is wrong; it does not say which.
Each one has to be settled against Excel's documented behaviour.
"""

from __future__ import annotations

import argparse
import logging
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT), str(ROOT / "tests")]

from conftest import make_xlsx  # noqa: E402
from provenance import diff  # noqa: E402
from provenance.canon import read_package  # noqa: E402
from provenance.compiler import compile_workbook, load_module  # noqa: E402
from provenance.model import read_workbook  # noqa: E402

# Column A of sheet In: one awkward value per row. Row 30 is left empty.
INPUTS = [0, 1, -1, 2.5, -2.5, 0.5, 1.5, 2.675, 1e15, 1e-10, 123456789.123, -0.0001,
          "", "abc", "5", " 5 ", "1e3", "TRUE", "Abc", True, False, "#N/A", "#DIV/0!",
          45351, 3, 10, 100, 0.1, -3.7]
EMPTY_ROW = 30
ROWS = list(range(1, len(INPUTS) + 1)) + [EMPTY_ROW]

UNARY = [
    "ABS(x)", "INT(x)", "SIGN(x)", "SQRT(x)", "LN(x)", "LOG10(x)", "LOG(x,2)", "EXP(x)",
    "ROUND(x,0)", "ROUND(x,1)", "ROUND(x,2)", "ROUND(x,-1)", "ROUNDUP(x,0)", "ROUNDUP(x,1)",
    "ROUNDDOWN(x,0)", "ROUNDDOWN(x,1)", "TRUNC(x)", "TRUNC(x,1)", "MOD(x,3)", "MOD(7,x)", "MOD(x,-2)",
    "POWER(x,2)", "POWER(2,x)", "x^0.5", "-x", "+x", "x%", "x+0", "x*1", "x/2", "1/x", "x-x",
    'x&""', 'x&"|"', "LEN(x)", "LEFT(x,2)", "LEFT(x)", "RIGHT(x,2)", "MID(x,2,2)", "UPPER(x)",
    "LOWER(x)", "TRIM(x)", "VALUE(x)", "REPT(x,2)", 'FIND("b",x)', 'SEARCH("B",x)',
    'SUBSTITUTE(x,"b","X")', 'EXACT(x,"abc")', "CONCATENATE(x,x)", "CONCAT(x,x)",
    "ISNUMBER(x)", "ISTEXT(x)", "ISBLANK(x)", "ISERROR(x)", "ISERR(x)", "ISNA(x)", "ISLOGICAL(x)",
    "NOT(x)", "IF(x,1,2)", "AND(x)", "OR(x)", "AND(x,TRUE)", "OR(x,FALSE)", "IFERROR(x,99)", "IFNA(x,99)",
    "x=0", 'x=""', 'x="abc"', "x=TRUE", "x<1", 'x<"a"', "x>FALSE", "x<>5", "x>=0",
    "DATE(2024,x,1)", "DATE(x,1,1)", "YEAR(x)", "MONTH(x)", "DAY(x)",
    "SUM(x)", "SUM(x,1)", "AVERAGE(x)", "COUNT(x)", "COUNTA(x)", "MIN(x)", "MAX(x)", "PRODUCT(x)",
    "CHOOSE(x,10,20,30)", "INDEX(B1:B5,x)", "MATCH(x,In!$A$1:$A$29,0)",
    'COUNTIF(In!$A$1:$A$30,x)', "PMT(0.01,x,1000)", "FV(0.01,12,x)", "PV(0.01,12,x)", "NPV(0.1,x,2)",
]

# Range-level formulas over the whole input column and a clean numeric range.
FIXED = [
    "SUM(In!A1:A12)", "SUM(In!A1:A30)", "AVERAGE(In!A1:A12)", "COUNT(In!A1:A30)", "COUNTA(In!A1:A30)",
    "COUNTBLANK(In!A1:A30)", "MIN(In!A1:A12)", "MAX(In!A1:A12)", "MEDIAN(In!A1:A12)", "PRODUCT(In!A1:A7)",
    "SUMPRODUCT(In!A1:A12,In!A1:A12)", "SUMPRODUCT(In!A1:A20,In!A1:A20)", "CONCAT(In!A13:A20)",
    "AND(In!A1:A3)", "OR(In!A13:A21)", "AND(In!A13:A18)", "SUM(In!A13:A20)", "AVERAGE(In!A13:A20)",
    "MAX(In!A13:A18)", "MIN(In!A30:A30)", "AVERAGE(In!A30:A30)",
    *[f'COUNTIF(In!A1:A30,{c})' for c in ['"5"', '">2"', '"<>abc"', '"a*"', '"?b?"', '""', '"="', '"<>"',
                                           '">=TRUE"', "TRUE", "0", '"=5"', '"<>5"', '"~*"', '"<b"', '">a"',
                                           '"*"', '"<=0"', '"#N/A"', '"1e3"', '">=45351"', '"ABC"']],
    *[f'SUMIF(In!A1:A30,{c},In!A1:A30)' for c in ['">0"', '"<0"', '"<>0"', '"5"', '"abc"']],
    'SUMIFS(In!A1:A12,In!A1:A12,">0",In!A1:A12,"<3")', 'COUNTIFS(In!A1:A30,">0",In!A1:A30,"<100")',
    'AVERAGEIF(In!A1:A12,">0")', 'AVERAGEIFS(In!A1:A12,In!A1:A12,"<0")', 'MAXIFS(In!A1:A12,In!A1:A12,"<1")',
    'MINIFS(In!A1:A12,In!A1:A12,">1")', 'AVERAGEIF(In!A1:A12,">1e20")',
    # lookups over a small table in Tab!A1:C6 (sorted keys) and Tab!E1:E6 (unsorted)
    *[f'VLOOKUP({k},Tab!A1:C6,2,FALSE)' for k in ['"b"', '"B"', '"z"', "3", '"3"', '"b*"', '"?"', "TRUE"]],
    *[f'VLOOKUP({k},Tab!A1:C6,3,TRUE)' for k in ["0", "1", "2.5", "100", '"a"']],
    'VLOOKUP("b",Tab!A1:C6,4,FALSE)', 'VLOOKUP("b",Tab!A1:C6,0,FALSE)', 'HLOOKUP(2,Tab!G1:J2,2,FALSE)',
    'HLOOKUP(2.5,Tab!G1:J2,2)', *[f'MATCH({k},Tab!G1:J1,{m})' for k in ["2", "2.5", "0", "9"] for m in ("0", "1")],
    'MATCH(3,Tab!L1:L4,-1)', 'MATCH(9,Tab!L1:L4,-1)', 'MATCH("B*",Tab!A1:A6,0)', 'MATCH(5,Tab!E1:E6,0)',
    'INDEX(Tab!A1:C6,2,3)', 'INDEX(Tab!A1:C6,7,1)', 'INDEX(Tab!G1:J1,3)', 'INDEX(Tab!A1:A6,0,1)',
    'SUM(INDEX(Tab!A1:C6,0,3))', 'SUM(INDEX(Tab!A1:C6,2,0))', 'XLOOKUP("c",Tab!A1:A6,Tab!C1:C6)',
    'XLOOKUP("q",Tab!A1:A6,Tab!C1:C6,"none")', 'CHOOSE(2,"a","b")', 'CHOOSE(0,"a")',
    # operators and precedence
    "-2^2", "2^-2", "-2^-2", "2^3^2", "1-2-3", "8/2/2", "2*3%", '"a"&1+1', "1+2&3", "1=1=TRUE",
    '"A"="a"', '"a"<"B"', "TRUE>1", '"1"=1', "0.1+0.2=0.3", "0.1+0.2-0.3", "1/3*3=1", "(1/3)&\"\"",
    '1E+20&""', '1E-5&""', '123456789012345678&""', '0.000123&""', '1234567.891&""', '-0.5&""',
    '12345678901234.5&""', '2^53&""', '1/7&""', '100000000000000000000&""', '1E+15&""',
    "ROUND(2.5,0)", "ROUND(-2.5,0)", "ROUND(0.285,2)", "ROUND(1.005,2)", "ROUND(1234.5678,-2)",
    "ROUNDUP(-1.1,0)", "ROUNDDOWN(-1.9,0)", "INT(-0.5)", "MOD(-7,3)", "MOD(7,-3)", "MOD(5.5,2)",
    "MOD(1,0)", "SQRT(-1)", "LN(0)", "0^0", "0^-1", "(-8)^(1/3)", "10^400", "-10^400",
    "DATE(2024,13,1)", "DATE(2024,0,1)", "DATE(2024,2,30)", "DATE(99,1,1)", "DATE(1900,3,1)",
    "DATE(2024,1,-5)", "YEAR(0)", "MONTH(0)", "DAY(0)", "DAY(60)", "MONTH(61)", "YEAR(2958465)", "YEAR(-1)",
    'LEFT("abc",-1)', 'MID("abc",0,1)', 'MID("abc",5,1)', 'RIGHT("abc",10)', 'FIND("",  "abc")',
    'FIND("c","abc",4)', 'SUBSTITUTE("aaa","a","b",2)', 'SUBSTITUTE("aaa","a","b",0)', 'REPT("a",-1)',
    'TRIM("  a  b  ")', 'VALUE("1,000")', 'VALUE("$5")', 'VALUE("50%")', 'VALUE("")', 'VALUE("1e400")',
    'LEN(1/3)', 'UPPER(TRUE)', "IFS(FALSE,1,TRUE,2)", "IFS(FALSE,1)", 'SWITCH(2,1,"a",2,"b")',
    'SWITCH(3,1,"a","z")', 'SWITCH(3,1,"a")', "PMT(0,12,1200)", "PMT(0.05/12,360,-200000,0,1)",
    "FV(0.05,10,-100,-1000,1)", "PV(0.08/12,240,500)", "NPV(0.1,-1000,300,400,500)", "NPV(-1,1)",
    "SUM()", "COUNT()", "AND()", "IF(TRUE,)", "IF(FALSE,1,)", "SUM(,1)", "MAX(,-1)",
]

TABLE = {
    "A1": "a", "B1": 1, "C1": 10, "A2": "b", "B2": 2, "C2": 20, "A3": "c", "B3": 3, "C3": 30,
    "A4": 3, "B4": 4, "C4": 40, "A5": "B", "B5": 5, "C5": 50, "A6": True, "B6": 6, "C6": 60,
    "E1": 9, "E2": 2, "E3": 7, "E4": 5, "E5": 1, "E6": 5,
    "G1": 1, "H1": 2, "I1": 3, "J1": 4, "G2": "w", "H2": "x", "I2": "y", "J2": "z",
    "L1": 9, "L2": 5, "L3": 3, "L4": 1,
}


def build():
    from provenance.formula import a1
    # Strings starting with '#' become error cells (make_xlsx convention).
    inputs = {f"A{i + 1}": v for i, v in enumerate(INPUTS)}
    out: dict[str, object] = {f"B{i}": i * 10 for i in range(1, 6)}
    cases: dict[str, str] = {}
    col = 3
    for template in UNARY:
        for r in ROWS:
            ref = a1(r, col)
            # Function names are upper case, so lower-case x is only ever the placeholder.
            text = template.replace("x", f"In!A{r}").replace("INDEX(B1:B5", "INDEX(Out!$B$1:$B$5")
            out[ref] = (text, None)
            cases[f"Out!{ref}"] = f"{template}  [x = {describe(r)}]"
        col += 1
    for i, text in enumerate(FIXED, 1):
        out[a1(i, col)] = (text, None)
        cases[f"Out!{a1(i, col)}"] = text
    sheets = {"In": inputs, "Out": out, "Tab": TABLE}
    return make_xlsx(sheets), cases


def describe(row: int) -> str:
    if row == EMPTY_ROW:
        return "empty cell"
    return repr(INPUTS[row - 1])


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--show", type=int, default=400)
    args = parser.parse_args()
    logging.disable(logging.CRITICAL)

    data, cases = build()
    pkg = read_package(data)
    model = read_workbook(pkg)
    source, report = compile_workbook(model, "0" * 64)
    ours, failed = load_module(source)["compute"]()
    theirs = diff.FormulasLibrary().run(pkg, model, [{}])[0]

    tally = defaultdict(lambda: defaultdict(int))
    disagreements = []
    for cell_key, label in cases.items():
        func = label.split("  [")[0] if "  [" in label else label
        if cell_key in report["unsupported"] or cell_key in failed:
            tally[func]["we_decline"] += 1
            continue
        if cell_key not in theirs:
            tally[func]["they_decline"] += 1
            continue
        a, b = diff.plain(ours[cell_key]), theirs[cell_key]
        result = diff.compare(a, b)
        tally[func]["agree" if result != "mismatch" else "differ"] += 1
        if result == "mismatch":
            disagreements.append((label, a, b))

    total = sum(sum(v.values()) for v in tally.values())
    agree = sum(v["agree"] for v in tally.values())
    differ = len(disagreements)
    print(f"{total} formulas: {agree} agree, {differ} differ, "
          f"{sum(v['we_decline'] for v in tally.values())} we decline, {sum(v['they_decline'] for v in tally.values())} formulas-library declines")
    for label, a, b in disagreements[:args.show]:
        print(f"DIFF  {label:60}  ours={a!r:28} formulas={b!r}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
