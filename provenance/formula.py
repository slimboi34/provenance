"""Excel formula lexer and parser.

Works on formulas exactly as stored in the file: no leading '=', en-US
function names, ',' between arguments. Anything outside the supported grammar
raises `FormulaError`; callers report the cell as unsupported rather than
guess.

AST nodes are tuples:
    ("num", float) ("str", s) ("bool", b) ("err", code) ("missing",)
    ("ref", Ref) ("name", text) ("func", NAME, [args])
    ("bin", op, left, right) ("un", op, operand) ("pct", operand)
    ("array", [[literal, ...], ...])
"""

from __future__ import annotations

import re
from dataclasses import dataclass, replace

MAX_ROW = 1_048_576
MAX_COL = 16_384


class FormulaError(Exception):
    pass


def col_number(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + ord(ch) - 64
    return n


def col_letters(n: int) -> str:
    out = ""
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def split_a1(ref: str) -> tuple[int, int]:
    """'B3' -> (row 3, col 2)."""
    m = re.fullmatch(r"\$?([A-Za-z]{1,3})\$?(\d+)", ref)
    if not m:
        raise FormulaError(f"bad cell reference {ref!r}")
    return int(m.group(2)), col_number(m.group(1))


def a1(row: int, col: int) -> str:
    return f"{col_letters(col)}{row}"


@dataclass(frozen=True)
class Ref:
    """A cell, rectangle, whole-column or whole-row reference.

    Whole columns have r1 = r2 = None; whole rows have c1 = c2 = None.
    `sheet` is the prefix exactly as written (None when absent).
    """
    sheet: str | None
    r1: int | None
    c1: int | None
    r2: int | None
    c2: int | None
    abs_r1: bool = False
    abs_c1: bool = False
    abs_r2: bool = False
    abs_c2: bool = False
    single: bool = False

    def text(self) -> str:
        def corner(r, c, ar, ac):
            return (("$" if ac else "") + col_letters(c) if c else "") + (("$" if ar else "") + str(r) if r else "")
        body = corner(self.r1, self.c1, self.abs_r1, self.abs_c1)
        if not self.single:
            body += ":" + corner(self.r2, self.c2, self.abs_r2, self.abs_c2)
        return (self.sheet + "!" if self.sheet else "") + body

    def shifted(self, drow: int, dcol: int) -> "Ref | None":
        """Move the relative parts; None if the result falls off the sheet."""
        def move(v, absolute, delta, limit):
            if v is None or absolute:
                return v
            v += delta
            return v if 1 <= v <= limit else -1
        r1 = move(self.r1, self.abs_r1, drow, MAX_ROW)
        r2 = move(self.r2, self.abs_r2, drow, MAX_ROW)
        c1 = move(self.c1, self.abs_c1, dcol, MAX_COL)
        c2 = move(self.c2, self.abs_c2, dcol, MAX_COL)
        if -1 in (r1, r2, c1, c2):
            return None
        return replace(self, r1=r1, r2=r2, c1=c1, c2=c2)


def sheet_name(prefix: str) -> str:
    """Turn a written sheet prefix ('My Sheet' with quotes, or bare) into the name."""
    if prefix.startswith("'") and prefix.endswith("'"):
        return prefix[1:-1].replace("''", "'")
    return prefix


# ----------------------------------------------------------------------- lexer

_CELL = r"(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})"
_RE_REF = re.compile(
    r"(?:((?:'(?:[^']|'')*')|(?:[A-Za-z0-9_.¡-￿]+))!)?"
    rf"(?:{_CELL}(?::{_CELL})?"
    r"|(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})"
    r"|(\$?)([0-9]{1,7}):(\$?)([0-9]{1,7}))"
    r"(?![A-Za-z0-9_.(\[!])"
)
_RE_STRING = re.compile(r'"(?:[^"]|"")*"')
_RE_ERROR = re.compile(r"#(?:N/A|DIV/0!|REF!|VALUE!|NAME\?|NUM!|NULL!|SPILL!|CALC!|GETTING_DATA)")
_RE_NUMBER = re.compile(r"(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?")
_RE_BOOL = re.compile(r"(?:TRUE|FALSE)(?![A-Za-z0-9_.(])", re.I)
_RE_FUNC = re.compile(r"([A-Za-z_][A-Za-z0-9_.]*)\(")
_RE_NAME = re.compile(r"[A-Za-z_\\¡-￿][A-Za-z0-9_.\\?¡-￿]*")
_RE_OP = re.compile(r"<=|>=|<>|[-+*/^&=<>%(),;{}]")
_RE_SPACE = re.compile(r"[ \t\r\n]+")


@dataclass(frozen=True)
class Token:
    kind: str      # str err ref num bool func name op space
    text: str
    value: object = None


def _ref_from_match(m: re.Match) -> Ref | None:
    g = m.groups()
    sheet = g[0]
    if g[2] is not None:  # cell or cell:cell
        c1, r1 = col_number(g[2]), int(g[4])
        if g[6] is not None:
            c2, r2 = col_number(g[6]), int(g[8])
            ref = Ref(sheet, r1, c1, r2, c2, bool(g[3]), bool(g[1]), bool(g[7]), bool(g[5]))
        else:
            ref = Ref(sheet, r1, c1, r1, c1, bool(g[3]), bool(g[1]), bool(g[3]), bool(g[1]), single=True)
    elif g[10] is not None:  # A:C
        ref = Ref(sheet, None, col_number(g[10]), None, col_number(g[12]), abs_c1=bool(g[9]), abs_c2=bool(g[11]))
    else:  # 1:3
        ref = Ref(sheet, int(g[14]), None, int(g[16]), None, abs_r1=bool(g[13]), abs_r2=bool(g[15]))
    for r in (ref.r1, ref.r2):
        if r is not None and not 1 <= r <= MAX_ROW:
            return None
    for c in (ref.c1, ref.c2):
        if c is not None and not 1 <= c <= MAX_COL:
            return None
    return ref


def tokenize(text: str) -> list[Token]:
    tokens: list[Token] = []
    i = 0
    while i < len(text):
        m = _RE_SPACE.match(text, i)
        if m:
            tokens.append(Token("space", m.group()))
            i = m.end()
            continue
        m = _RE_STRING.match(text, i)
        if m:
            tokens.append(Token("str", m.group(), m.group()[1:-1].replace('""', '"')))
            i = m.end()
            continue
        m = _RE_ERROR.match(text, i)
        if m:
            tokens.append(Token("err", m.group(), m.group()))
            i = m.end()
            continue
        m = _RE_REF.match(text, i)
        if m:
            ref = _ref_from_match(m)
            if ref is not None:
                tokens.append(Token("ref", m.group(), ref))
                i = m.end()
                continue
        m = _RE_NUMBER.match(text, i)
        if m:
            tokens.append(Token("num", m.group(), float(m.group())))
            i = m.end()
            continue
        m = _RE_BOOL.match(text, i)
        if m:
            tokens.append(Token("bool", m.group(), m.group().upper() == "TRUE"))
            i = m.end()
            continue
        m = _RE_FUNC.match(text, i)
        if m:
            tokens.append(Token("func", m.group(), m.group(1).upper()))
            i = m.end()
            continue
        m = _RE_NAME.match(text, i)
        if m:
            tokens.append(Token("name", m.group(), m.group()))
            i = m.end()
            continue
        m = _RE_OP.match(text, i)
        if m:
            tokens.append(Token("op", m.group(), m.group()))
            i = m.end()
            continue
        raise FormulaError(f"unsupported syntax at {text[i:i + 12]!r}")
    return tokens


def shift_formula(text: str, drow: int, dcol: int) -> str:
    """Rewrite a formula as Excel does when filling it `drow` down, `dcol` right.

    Used to expand shared formulas. Only reference tokens change; every other
    character is copied through untouched.
    """
    out = []
    for tok in tokenize(text):
        if tok.kind == "ref":
            moved = tok.value.shifted(drow, dcol)
            out.append(moved.text() if moved else "#REF!")
        else:
            out.append(tok.text)
    return "".join(out)


# ---------------------------------------------------------------------- parser

_BINARY = {"=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1, "&": 2, "+": 3, "-": 3, "*": 4, "/": 4, "^": 5}
_PERCENT_BP = 6


class _Parser:
    def __init__(self, text: str):
        self.toks = [t for t in tokenize(text) if t.kind != "space"]
        self.pos = 0

    def peek(self) -> Token | None:
        return self.toks[self.pos] if self.pos < len(self.toks) else None

    def take(self) -> Token:
        tok = self.peek()
        if tok is None:
            raise FormulaError("formula ends unexpectedly")
        self.pos += 1
        return tok

    def expect(self, op: str) -> None:
        tok = self.take()
        if tok.kind != "op" or tok.text != op:
            raise FormulaError(f"expected {op!r}, found {tok.text!r}")

    def is_op(self, op: str) -> bool:
        tok = self.peek()
        return tok is not None and tok.kind == "op" and tok.text == op

    def expr(self, min_bp: int = 0):
        left = self.prefix()
        while True:
            tok = self.peek()
            if tok is None or tok.kind != "op":
                break
            if tok.text == "%" and _PERCENT_BP >= min_bp:
                self.pos += 1
                left = ("pct", left)
                continue
            bp = _BINARY.get(tok.text)
            if bp is None or bp < min_bp:
                break
            self.pos += 1
            left = ("bin", tok.text, left, self.expr(bp + 1))  # all Excel operators are left-associative
        return left

    def prefix(self):
        tok = self.take()
        if tok.kind == "op" and tok.text in "+-":
            return ("un", tok.text, self.prefix())  # negation binds tighter than ^
        if tok.kind in ("num", "str", "bool", "err"):
            return (tok.kind, tok.value)
        if tok.kind == "ref":
            return ("ref", tok.value)
        if tok.kind == "name":
            return ("name", tok.value)
        if tok.kind == "func":
            args = []
            if self.is_op(")"):
                self.pos += 1
                return ("func", tok.value, args)
            while True:
                if self.is_op(",") or self.is_op(")"):
                    args.append(("missing",))
                else:
                    args.append(self.expr())
                if self.is_op(","):
                    self.pos += 1
                    continue
                self.expect(")")
                return ("func", tok.value, args)
        if tok.kind == "op" and tok.text == "(":
            inner = self.expr()
            self.expect(")")
            return inner
        if tok.kind == "op" and tok.text == "{":
            return self.array()
        raise FormulaError(f"unexpected {tok.text!r}")

    def array(self):
        rows, row = [], []
        while True:
            sign = 1.0
            tok = self.take()
            if tok.kind == "op" and tok.text in "+-":
                sign = -1.0 if tok.text == "-" else 1.0
                tok = self.take()
                if tok.kind != "num":
                    raise FormulaError("bad array constant")
            if tok.kind == "num":
                row.append(("num", sign * tok.value))
            elif tok.kind in ("str", "bool", "err"):
                row.append((tok.kind, tok.value))
            else:
                raise FormulaError("bad array constant")
            sep = self.take()
            if sep.kind != "op" or sep.text not in ",;}":
                raise FormulaError("bad array constant")
            if sep.text == ",":
                continue
            rows.append(row)
            row = []
            if sep.text == "}":
                break
        if len({len(r) for r in rows}) != 1:
            raise FormulaError("ragged array constant")
        return ("array", rows)


def parse(text: str):
    parser = _Parser(text)
    tree = parser.expr()
    if parser.peek() is not None:
        raise FormulaError(f"unexpected {parser.peek().text!r}")
    return tree
