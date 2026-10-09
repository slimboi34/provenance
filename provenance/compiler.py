"""Deterministic compiler: workbook model -> Python module.

One Python function per formula cell, built by fixed rules from the parsed
formula. No model, no heuristics: the same workbook always gives the same
source. A formula that uses anything outside the supported subset is not
compiled; it is listed in UNSUPPORTED with the reason.
"""

from __future__ import annotations

from . import __version__
from .formula import MAX_COL, MAX_ROW, FormulaError, Ref, a1, parse, sheet_name
from .model import formula_cells, key
from .runtime import FUNCTIONS, VOLATILE

_OPS = {"+": "op_add", "-": "op_sub", "*": "op_mul", "/": "op_div", "^": "op_pow", "&": "op_concat",
        "=": "op_eq", "<>": "op_ne", "<": "op_lt", ">": "op_gt", "<=": "op_le", ">=": "op_ge"}
_DATE_FUNCTIONS = {"DATE", "YEAR", "MONTH", "DAY"}


class _Unsupported(Exception):
    pass


class _Compiler:
    def __init__(self, model: dict):
        self.model = model
        self.sheets = {s["name"].upper(): s for s in model["sheets"]}
        self.names: dict[tuple[str | None, str], str] = {
            (n["scope"], n["name"].upper()): n["refers_to"] for n in model["defined_names"]
        }
        self.max_row = max([s["max_row"] for s in model["sheets"]] + [1])
        self.max_col = max([s["max_col"] for s in model["sheets"]] + [1])
        self.deps: list[str] = []

    def ref(self, ref: Ref, here: str, as_arg: bool) -> str:
        name = sheet_name(ref.sheet) if ref.sheet else here
        if name.startswith("[") or "]" in name:
            raise _Unsupported(f"reference to another workbook ({ref.sheet})")
        sheet = self.sheets.get(name.upper())
        if sheet is None:
            return "REF"
        r1, c1, r2, c2 = ref.r1, ref.c1, ref.r2, ref.c2
        if r1 is None:
            r1, r2 = 1, MAX_ROW
        if c1 is None:
            c1, c2 = 1, MAX_COL
        r1, r2, c1, c2 = min(r1, r2), max(r1, r2), min(c1, c2), max(c1, c2)
        # Anything past the last used row/column of the whole workbook is empty,
        # so it is not materialised; the runtime is told the range was clamped.
        # One limit for every sheet keeps same-sized ranges the same size.
        clamp_r2, clamp_c2 = min(r2, max(self.max_row, r1)), min(c2, max(self.max_col, c1))
        clamped = (clamp_r2, clamp_c2) != (r2, c2)
        r2, c2 = clamp_r2, clamp_c2
        if (r1, c1) == (r2, c2) and not clamped:
            cell_key = key(sheet["name"], a1(r1, c1))
            self.deps.append(cell_key)
            return f"x.{'ref' if as_arg else 'cell'}({cell_key!r})"
        range_key = key(sheet["name"], a1(r1, c1) + ":" + a1(r2, c2))
        self.deps.append(range_key)
        return f"x.rng({range_key!r}{', True' if clamped else ''})"

    def references(self, text: str, here: str) -> tuple:
        """Cells and ranges a formula reads, even if it cannot be compiled (best effort)."""
        try:
            tree = parse(text)
        except FormulaError:
            return ()
        self.deps = []
        todo = [tree]
        while todo:
            node = todo.pop()
            if node[0] == "ref":
                try:
                    self.ref(node[1], here, True)
                except _Unsupported:
                    pass
            elif node[0] == "name":
                try:
                    self.name(node[1], here, True, 0)
                except _Unsupported:
                    pass
            elif node[0] == "func":
                todo.extend(node[2])
            elif node[0] == "bin":
                todo.extend(node[2:])
            elif node[0] in ("un", "pct"):
                todo.append(node[-1])
        return tuple(dict.fromkeys(self.deps))

    def name(self, text: str, here: str, as_arg: bool, depth: int) -> str:
        target = self.names.get((here, text.upper())) or self.names.get((None, text.upper()))
        if target is None:
            raise _Unsupported(f"unknown name {text!r} (table reference or undefined name)")
        if depth > 8:
            raise _Unsupported(f"defined name {text!r} is nested too deeply")
        try:
            tree = parse(target)
        except FormulaError as exc:
            raise _Unsupported(f"defined name {text!r} refers to something unsupported: {exc}") from None
        if tree[0] == "ref" and tree[1].sheet is None:
            raise _Unsupported(f"defined name {text!r} has a relative reference")
        return self.expr(tree, here, as_arg, depth + 1)

    def expr(self, node, here: str, as_arg: bool = False, depth: int = 0) -> str:
        kind = node[0]
        if kind == "num":
            return repr(node[1])
        if kind == "str":
            return repr(node[1])
        if kind == "bool":
            return "True" if node[1] else "False"
        if kind == "err":
            return f"error({node[1]!r})"
        if kind == "missing":
            return "None"
        if kind == "ref":
            return self.ref(node[1], here, as_arg)
        if kind == "name":
            return self.name(node[1], here, as_arg, depth)
        if kind == "bin":
            return f"{_OPS[node[1]]}({self.expr(node[2], here, depth=depth)}, {self.expr(node[3], here, depth=depth)})"
        if kind == "un":
            return f"{'op_neg' if node[1] == '-' else 'op_pos'}({self.expr(node[2], here, depth=depth)})"
        if kind == "pct":
            return f"op_pct({self.expr(node[1], here, depth=depth)})"
        if kind == "array":
            rows = ", ".join("[" + ", ".join(self.expr(v, here) for v in row) + "]" for row in node[1])
            return f"Arr([{rows}])"
        if kind == "func":
            fname = node[1]
            for prefix in ("_XLFN._XLWS.", "_XLFN."):
                if fname.startswith(prefix):
                    fname = fname[len(prefix):]
            if fname in VOLATILE:
                raise _Unsupported(f"{fname} is volatile or builds references at run time")
            if fname not in FUNCTIONS:
                raise _Unsupported(f"function {fname} is not supported yet")
            if fname in _DATE_FUNCTIONS and self.model["date1904"]:
                raise _Unsupported("date functions in a 1904-date-system workbook")
            args = ", ".join(self.expr(a, here, as_arg=True, depth=depth) for a in node[2])
            return f"fn_{fname}({args})"
        raise _Unsupported(f"unhandled syntax node {kind}")


def compile_workbook(model: dict, source_sha256: str) -> tuple[str, dict]:
    """Return (python_source, report).

    report = {"compiled": [cell keys], "unsupported": {cell key: reason},
              "unreadable_inputs": {cell key: reason}}
    """
    comp = _Compiler(model)
    inputs: list[str] = []
    bad_inputs: dict[str, str] = {}
    for sheet in model["sheets"]:
        for ref, cell in sheet["cells"].items():
            if "f" in cell:
                continue
            if "unreadable" in cell:
                bad_inputs[key(sheet["name"], ref)] = cell["unreadable"]
                continue
            value = f"error({cell['e']!r})" if "e" in cell else repr(cell.get("v"))
            if "e" in cell or cell.get("v") is not None:
                inputs.append(f"    {key(sheet['name'], ref)!r}: {value},")

    functions: list[str] = []
    table: list[str] = []
    compiled: list[str] = []
    deps: list[str] = []
    unsupported: dict[str, str] = {}
    index = {s["name"]: i for i, s in enumerate(model["sheets"])}
    for cell_key, sheet, ref, cell in formula_cells(model):
        comp.deps = []
        try:
            if "unsupported" in cell:
                raise _Unsupported(cell["unsupported"])
            body = comp.expr(parse(cell["f"]), sheet["name"])
        except FormulaError as exc:
            unsupported[cell_key] = f"formula syntax not supported: {exc}"
            continue
        except _Unsupported as exc:
            unsupported[cell_key] = str(exc)
            refs = comp.references(cell["f"], sheet["name"])  # still needed to find circular references
            if refs:
                deps.append(f"    {cell_key!r}: {refs!r},")
            continue
        fname = f"s{index[sheet['name']]}_{ref}"
        comment = cell["f"].replace("\n", " ")
        functions.append(f"def {fname}(x):  # {cell_key} = {comment}\n    return {body}\n")
        table.append(f"    {cell_key!r}: {fname},")
        deps.append(f"    {cell_key!r}: {tuple(dict.fromkeys(comp.deps))!r},")
        compiled.append(cell_key)

    source = "\n".join([
        f'"""Compiled from the workbook with SHA-256 {source_sha256}.',
        "",
        f"Generated by provenance {__version__}; do not edit. Each function is one formula cell.",
        "compute(overrides) recalculates the workbook; overrides maps 'Sheet!A1' to a new input value.",
        '"""',
        "",
        "from provenance.runtime import *  # noqa: F401,F403",
        "from provenance.runtime import evaluate",
        "",
        f"SOURCE_SHA256 = {source_sha256!r}",
        "",
        "INPUTS = {",
        *inputs,
        "}",
        "",
        "",
        "\n\n".join(functions),
        "",
        "FORMULAS = {",
        *table,
        "}",
        "",
        "# Cells and ranges each formula reads; used to compute cells in dependency order.",
        "DEPS = {",
        *deps,
        "}",
        "",
        "UNSUPPORTED = {",
        *(f"    {k!r}: {v!r}," for k, v in {**bad_inputs, **unsupported}.items()),
        "}",
        "",
        "",
        "def compute(overrides=None):",
        '    """Return (values, failed) for every formula cell."""',
        "    return evaluate(INPUTS, FORMULAS, UNSUPPORTED, overrides, DEPS)",
        "",
    ])
    return source, {"compiled": compiled, "unsupported": unsupported, "unreadable_inputs": bad_inputs}


def load_module(source: str, chunk_lines: int = 4000):
    """Execute generated source and return its namespace.

    The source is run in order, in pieces split at top-level statements, which
    is equivalent to running it whole: compiling a large workbook's module in
    one piece needs several GB of memory. Generated code never puts a line
    break inside a string, so a line starting in column 0 always begins a new
    top-level statement.
    """
    namespace: dict = {"__name__": "provenance_compiled"}
    lines = source.splitlines(keepends=True)
    start = 0
    while start < len(lines):
        end = min(start + chunk_lines, len(lines))
        while end < len(lines) and (lines[end][:1] in (" ", "}", "\n", "") or lines[end].startswith(")")):
            end += 1
        chunk = "\n" * start + "".join(lines[start:end])  # keep line numbers in tracebacks
        exec(compile(chunk, "<provenance compiled workbook>", "exec"), namespace)
        start = end
    return namespace
