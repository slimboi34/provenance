"""Excel semantics for compiled workbooks.

Generated modules call `op_*` for operators and `fn_*` for worksheet
functions. Values are: float, str, bool, None (empty cell), XlError, or Arr
(a rectangle of those). Errors are values, never exceptions. `Unsupported` is
raised when a formula needs behaviour this runtime does not model; the cell is
then reported as unverified instead of being given a guessed value.
"""

from __future__ import annotations

import math
import re
import sys
from datetime import date, timedelta
from decimal import ROUND_DOWN, ROUND_HALF_UP, ROUND_UP, Decimal, InvalidOperation
from bisect import bisect_left, bisect_right
from functools import lru_cache

from .formula import a1, split_a1


class Unsupported(Exception):
    pass


class XlError:
    __slots__ = ("code",)

    def __init__(self, code: str):
        self.code = code

    def __repr__(self):
        return self.code

    def __eq__(self, other):
        return isinstance(other, XlError) and other.code == self.code

    def __hash__(self):
        return hash(self.code)


NULL, DIV0, VALUE, REF, NAME, NUM, NA = (XlError(c) for c in ("#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"))
_ERRORS = {e.code: e for e in (NULL, DIV0, VALUE, REF, NAME, NUM, NA)}


def error(code: str) -> XlError:
    return _ERRORS.get(code) or XlError(code)


class Arr:
    """A rectangle of values (a range's contents or an array constant).

    A `clamped` Arr is a range that extends past the last used row/column of
    the workbook: only the used part is held, and every cell beyond it has the
    value `pad` (None for a plain range, or what an operator made of None).
    Functions whose answer depends on that unseen part must not use `flat`.
    """
    __slots__ = ("rows", "clamped", "pad")

    def __init__(self, rows, clamped=False, pad=None):
        self.rows = rows
        self.clamped = clamped
        self.pad = pad

    @property
    def shape(self):
        return len(self.rows), len(self.rows[0]) if self.rows else 0

    @property
    def single(self):
        return not self.clamped and self.shape == (1, 1)

    def flat(self):
        if self.clamped and self.pad is not None:
            raise Unsupported("a calculation over a very large range (beyond the used area)")
        for row in self.rows:
            yield from row

    def __repr__(self):
        return f"Arr({self.rows!r})"


# ------------------------------------------------------------------- coercions

_RE_NUM = re.compile(r"[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?")


def sc(v):
    """Scalar context: unwrap a 1x1 rectangle; larger ones are not modelled."""
    if isinstance(v, Arr):
        if v.single:
            return v.rows[0][0]
        raise Unsupported("a multi-cell range is used where a single value is expected")
    return v


def is_num(v) -> bool:
    return isinstance(v, float)


_RE_GROUPED = re.compile(r"[+-]?\d{1,3}(?:,\d{3})+(?:\.\d*)?")


def text_to_num(text: str):
    """Read text as a number the way Excel (en-US) does, for the forms modelled here.

    Plain numbers, thousands separators, a leading '$' and a trailing '%' are
    handled. Text without digits is #VALUE!. Anything else containing a digit
    (dates, times, fractions, accounting brackets) Excel may still read as a
    number, so it raises Unsupported instead of guessing.
    """
    s = text.strip(" ")
    percent = s.endswith("%")
    if percent:
        s = s[:-1].rstrip(" ")
    sign = ""
    if s[:1] in "+-" and s[1:2] == "$":
        sign, s = s[0], s[2:]
    elif s.startswith("$"):
        s = s[1:]
    s = sign + s
    if _RE_NUM.fullmatch(s):
        n = float(s)
    elif _RE_GROUPED.fullmatch(s):
        n = float(s.replace(",", ""))
    elif any(ch.isdigit() for ch in text):
        raise Unsupported(f"text {text!r} used as a number (Excel may read it as a date, time or number)")
    else:
        return VALUE
    return _check(n / 100.0 if percent else n)


def to_num(v):
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, float):
        return v
    if v is None:
        return 0.0
    if isinstance(v, str):
        return text_to_num(v)
    if isinstance(v, int):
        return float(v)
    return v  # XlError


def _check(x):
    if isinstance(x, XlError):
        return x
    return x if math.isfinite(x) else NUM


def num_to_text(x: float) -> str:
    """Excel's General rendering when a number is turned into text."""
    if x == 0:
        return "0"
    s = "%.15g" % x
    if "e" in s:
        mantissa, exp = s.split("e")
        exp_n = int(exp)
        if -10 < exp_n < 0:  # Excel still writes small decimals out in full
            return ("%.*f" % (14 - exp_n, x)).rstrip("0").rstrip(".")
        return f"{mantissa}E{'+' if exp_n >= 0 else '-'}{abs(exp_n):02d}"
    return s


def to_text(v):
    if isinstance(v, str):
        return v
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        return num_to_text(v)
    if v is None:
        return ""
    return v


def to_bool(v):
    if isinstance(v, bool):
        return v
    if isinstance(v, float):
        return v != 0
    if v is None:
        return False
    if isinstance(v, str):
        u = v.upper()
        return True if u == "TRUE" else False if u == "FALSE" else VALUE
    return v


def _err(*vals):
    for v in vals:
        if isinstance(v, XlError):
            return v
    return None


# ------------------------------------------------------------------- operators

def _broadcast(f):
    def op(a, b):
        if isinstance(a, Arr) and a.single:
            a = a.rows[0][0]
        if isinstance(b, Arr) and b.single:
            b = b.rows[0][0]
        if not isinstance(a, Arr) and not isinstance(b, Arr):
            return f(a, b)
        clamped = [x for x in (a, b) if isinstance(x, Arr) and x.clamped]
        if clamped:
            arrays = [x for x in (a, b) if isinstance(x, Arr)]
            if len(clamped) != len(arrays) or len({x.shape for x in arrays}) != 1:
                raise Unsupported("an operator combines a very large range with a differently sized one")
            pa, pb = (x.pad if isinstance(x, Arr) else x for x in (a, b))
            pick = lambda x, r, c: x.rows[r][c] if isinstance(x, Arr) else x  # noqa: E731
            rows, cols = arrays[0].shape
            return Arr([[f(pick(a, r, c), pick(b, r, c)) for c in range(cols)] for r in range(rows)], True, f(pa, pb))
        (ar, ac), (br, bc) = (x.shape if isinstance(x, Arr) else (1, 1) for x in (a, b))

        def pick(x, shape, r, c):
            if not isinstance(x, Arr):
                return x
            rr = 0 if shape[0] == 1 else r
            cc = 0 if shape[1] == 1 else c
            return x.rows[rr][cc] if rr < shape[0] and cc < shape[1] else NA

        return Arr([[f(pick(a, (ar, ac), r, c), pick(b, (br, bc), r, c)) for c in range(max(ac, bc))] for r in range(max(ar, br))])
    return op


def _map(f):
    def op(a):
        if isinstance(a, Arr):
            if a.single:
                return f(a.rows[0][0])
            return Arr([[f(v) for v in row] for row in a.rows], a.clamped, f(a.pad) if a.clamped else None)
        return f(a)
    return op


def _arith(f):
    @_broadcast
    def op(a, b):
        a, b = to_num(a), to_num(b)
        return _err(a, b) or f(a, b)
    return op


op_add = _arith(lambda a, b: _check(a + b))
op_sub = _arith(lambda a, b: _check(a - b))
op_mul = _arith(lambda a, b: _check(a * b))
op_div = _arith(lambda a, b: DIV0 if b == 0 else _check(a / b))


def _pow(a, b):
    if a == 0:
        return NUM if b == 0 else DIV0 if b < 0 else 0.0
    if a < 0 and not b.is_integer():
        return NUM
    try:
        return _check(a ** b)
    except (OverflowError, ZeroDivisionError):
        return NUM


op_pow = _arith(_pow)


@_broadcast
def op_concat(a, b):
    return _err(a, b) or to_text(a) + to_text(b)


@_map
def op_neg(a):
    a = to_num(a)
    return a if isinstance(a, XlError) else -a


@_map
def op_pos(a):
    return a


@_map
def op_pct(a):
    a = to_num(a)
    return a if isinstance(a, XlError) else a / 100.0


def _rank(v) -> int:
    return 2 if isinstance(v, bool) else 1 if isinstance(v, str) else 0


def fold(s: str) -> str:
    """Upper-case per character, leaving characters whose upper case is longer
    (German ß -> SS) unchanged: Excel never treats "ß" as equal to "ss"."""
    up = s.upper()
    if len(up) == len(s):
        return up
    return "".join(c if len(c.upper()) != 1 else c.upper() for c in s)


def compare(a, b) -> int:
    """Excel ordering: numbers < text < logicals; text is case-insensitive."""
    if a is None and b is None:
        return 0
    if a is None:
        a = "" if isinstance(b, str) else False if isinstance(b, bool) else 0.0
    if b is None:
        b = "" if isinstance(a, str) else False if isinstance(a, bool) else 0.0
    ra, rb = _rank(a), _rank(b)
    if ra != rb:
        return -1 if ra < rb else 1
    if ra == 1:
        a, b = fold(a), fold(b)
    return -1 if a < b else 1 if a > b else 0


_RE_PLAIN_TEXT = re.compile(r"[A-Za-z0-9 ]*")


def text_order_check(*texts) -> None:
    """Excel orders text by locale collation, which differs from character order
    for punctuation and accents. Only plain letters, digits and spaces are
    ordered here; anything else is reported instead of guessed."""
    for t in texts:
        if isinstance(t, str) and not _RE_PLAIN_TEXT.fullmatch(t):
            raise Unsupported(f"ordering text that contains punctuation or accents ({t[:20]!r})")


def order(a, b) -> int:
    """compare() for <, >, <=, >= and sorted lookups."""
    if isinstance(a, str) and isinstance(b, str):
        text_order_check(a, b)
    return compare(a, b)


def _cmp(test, ordering=False):
    cmp = order if ordering else compare

    @_broadcast
    def op(a, b):
        return _err(a, b) or test(cmp(a, b))
    return op


op_eq = _cmp(lambda c: c == 0)
op_ne = _cmp(lambda c: c != 0)
op_lt = _cmp(lambda c: c < 0, True)
op_gt = _cmp(lambda c: c > 0, True)
op_le = _cmp(lambda c: c <= 0, True)
op_ge = _cmp(lambda c: c >= 0, True)


# ---------------------------------------------------------------- aggregation

def _numbers(args):
    """Numbers the way SUM sees them, or the first error.

    Inside ranges only real numbers count. Arguments typed in directly are
    coerced (TRUE -> 1, "5" -> 5) and text that is not a number is #VALUE!.
    """
    out = []
    for a in args:
        if isinstance(a, Arr):
            for v in a.flat():
                if isinstance(v, XlError):
                    return v
                if isinstance(v, float):
                    out.append(v)
        else:
            n = to_num(a)
            if isinstance(n, XlError):
                return n
            out.append(n)
    return out


def _agg(f):
    def fn(*args):
        nums = _numbers(args)
        return nums if isinstance(nums, XlError) else f(nums)
    return fn


def _sum(nums):
    total = 0.0
    for n in nums:
        total += n
    return _check(total)


def _product(nums):
    total = 1.0
    for n in nums:
        total *= n
    return _check(total) if nums else 0.0


def _median(nums):
    if not nums:
        return NUM
    s = sorted(nums)
    mid = len(s) // 2
    return s[mid] if len(s) % 2 else (s[mid - 1] + s[mid]) / 2


fn_SUM = _agg(_sum)
fn_AVERAGE = _agg(lambda nums: _check(_sum(nums) / len(nums)) if nums else DIV0)
fn_MIN = _agg(lambda nums: min(nums) if nums else 0.0)
fn_MAX = _agg(lambda nums: max(nums) if nums else 0.0)
fn_PRODUCT = _agg(_product)
fn_MEDIAN = _agg(_median)


def fn_COUNT(*args):
    n = 0
    for a in args:
        if isinstance(a, Arr):
            n += sum(1 for v in a.flat() if isinstance(v, float))
        elif a is not None and not isinstance(to_num(a), XlError):
            n += 1
    return float(n)


def fn_COUNTA(*args):
    return float(sum(sum(1 for v in a.flat() if v is not None) if isinstance(a, Arr) else 1 for a in args))


def fn_COUNTBLANK(rng):
    arr = _arr(rng)
    if arr.clamped:
        raise Unsupported("COUNTBLANK over a very large range (beyond the used area)")
    return float(sum(1 for v in arr.flat() if v is None or v == ""))


def _zero(v) -> bool:
    """Would SUMPRODUCT treat this value as contributing nothing?"""
    return not isinstance(v, float) or v == 0.0


def fn_SUMPRODUCT(*args):
    for a in args:
        if not isinstance(a, Arr) or a.single:
            # Excel treats a single value or one-cell range as a scalar: text there is
            # #VALUE! (seen in Excel-saved test data); blanks and TRUE/FALSE are unconfirmed.
            v = sc(a)
            if isinstance(v, str):
                return VALUE
            if v is None or isinstance(v, bool):
                raise Unsupported("SUMPRODUCT with an empty or TRUE/FALSE single-cell argument")
    arrays = [_arr(a) for a in args]
    for a in arrays:  # Excel reports an error inside a range before a size mismatch
        e = _err(*(v for row in a.rows for v in row))
        if e:
            return e
    if any(a.clamped for a in arrays):
        if not all(a.clamped for a in arrays) or len({a.shape for a in arrays}) != 1:
            raise Unsupported("SUMPRODUCT over very large ranges of different sizes")
        if isinstance(_err(*(a.pad for a in arrays)), XlError) or not any(_zero(a.pad) for a in arrays):
            raise Unsupported("SUMPRODUCT where the unused part of a very large range is not zero")
    elif len({a.shape for a in arrays}) != 1:
        return VALUE
    total = 0.0
    for cells in zip(*([x for row in a.rows for x in row] for a in arrays)):
        e = _err(*cells)
        if e:
            return e
        p = 1.0
        for v in cells:
            p *= v if isinstance(v, float) else 0.0
        total += p
    return _check(total)


def _arr(v) -> Arr:
    return v if isinstance(v, Arr) else Arr([[v]])


# -------------------------------------------------------------------- logical

def fn_IF(cond, then=True, otherwise=False):
    c = to_bool(sc(cond))
    if isinstance(c, XlError):
        return c
    return then if c else otherwise


def fn_IFS(*args):
    if len(args) % 2:
        raise Unsupported("IFS with an odd number of arguments")
    for cond, value in zip(args[::2], args[1::2]):
        c = to_bool(sc(cond))
        if isinstance(c, XlError):
            return c
        if c:
            return value
    return NA


def fn_SWITCH(expr, *rest):
    expr = sc(expr)
    if isinstance(expr, XlError):
        return expr
    pairs, default = (rest[:-1], rest[-1]) if len(rest) % 2 else (rest, NA)
    for case, value in zip(pairs[::2], pairs[1::2]):
        case = sc(case)
        if not isinstance(case, XlError) and _rank(case) == _rank(expr) and compare(case, expr) == 0:
            return value
    return default


def _logical(args):
    """Logical values for AND/OR. Text inside ranges is ignored. Text typed
    directly as an argument (other than "TRUE"/"FALSE") is declined: two
    Excel-saved test files disagree on it (AND(TRUE,"0") = TRUE but
    AND("1",1) = #VALUE!)."""
    out = []
    for a in args:
        if isinstance(a, Arr) and not a.single:
            for v in a.flat():
                if isinstance(v, XlError):
                    return v
                if isinstance(v, (bool, float)):
                    out.append(bool(v))
        else:
            ref_like = isinstance(a, Arr)
            v = sc(a)
            if v is None or (isinstance(v, str) and ref_like):
                continue
            if isinstance(v, str) and v.upper() not in ("TRUE", "FALSE"):
                raise Unsupported("AND/OR with a text argument other than TRUE/FALSE")
            b = to_bool(v)
            if isinstance(b, XlError):
                return b
            out.append(b)
    return out if out else VALUE


def fn_AND(*args):
    vals = _logical(args)
    return vals if isinstance(vals, XlError) else all(vals)


def fn_OR(*args):
    vals = _logical(args)
    return vals if isinstance(vals, XlError) else any(vals)


def fn_NOT(v):
    b = to_bool(sc(v))
    return b if isinstance(b, XlError) else not b


def fn_TRUE():
    return True


def fn_FALSE():
    return False


def fn_IFERROR(v, alt=None):
    v = sc(v) if isinstance(v, Arr) and v.shape == (1, 1) else v
    return alt if isinstance(v, XlError) else v


def fn_IFNA(v, alt=None):
    v = sc(v) if isinstance(v, Arr) and v.shape == (1, 1) else v
    return alt if v == NA else v


def fn_ISERROR(v):
    return isinstance(sc(v), XlError)


def fn_ISERR(v):
    v = sc(v)
    return isinstance(v, XlError) and v != NA


def fn_ISNA(v):
    return sc(v) == NA


def fn_ISBLANK(v):
    return sc(v) is None


def fn_ISNUMBER(v):
    return isinstance(sc(v), float)


def fn_ISTEXT(v):
    return isinstance(sc(v), str)


def fn_ISLOGICAL(v):
    return isinstance(sc(v), bool)


def fn_NA():
    return NA


# ----------------------------------------------------------------------- math

def _numeric(f):
    """Wrap a function of floats: scalar context, coercion, error propagation."""
    def fn(*args):
        nums = [to_num(sc(a)) for a in args]
        e = _err(*nums)
        if e:
            return e
        try:
            return _check(f(*nums))
        except OverflowError:
            return NUM
    return fn


def _quantize(x: float, digits: float, mode) -> float:
    # Excel rounds the 15-significant-digit decimal value, not the binary one.
    q = Decimal(1).scaleb(-int(digits))
    try:
        return float(Decimal("%.15g" % x).quantize(q, rounding=mode))
    except InvalidOperation:  # too many digits: x is already a whole number at this precision
        if digits >= 0:
            return x
        raise Unsupported("rounding a number too large to represent") from None


fn_ABS = _numeric(abs)
fn_ROUND = _numeric(lambda x, d: _quantize(x, d, ROUND_HALF_UP))
fn_ROUNDUP = _numeric(lambda x, d: _quantize(x, d, ROUND_UP))
fn_ROUNDDOWN = _numeric(lambda x, d: _quantize(x, d, ROUND_DOWN))
fn_TRUNC = _numeric(lambda x, d=0.0: _quantize(x, d, ROUND_DOWN))
fn_INT = _numeric(lambda x: float(math.floor(x)))
fn_SIGN = _numeric(lambda x: float((x > 0) - (x < 0)))
fn_EXP = _numeric(math.exp)
fn_PI = _numeric(lambda: math.pi)


def _domain(f, ok, bad=NUM):
    def fn(*args):
        nums = [to_num(sc(a)) for a in args]
        e = _err(*nums)
        if e:
            return e
        if not ok(*nums):
            return bad
        try:
            return _check(f(*nums))
        except (OverflowError, ValueError, ZeroDivisionError):
            return NUM
    return fn


fn_SQRT = _domain(math.sqrt, lambda x: x >= 0)
fn_LN = _domain(math.log, lambda x: x > 0)
fn_LOG10 = _domain(math.log10, lambda x: x > 0)
fn_LOG = _domain(lambda x, base=10.0: math.log(x) / math.log(base), lambda x, base=10.0: x > 0 and base > 0 and base != 1)
fn_MOD = _domain(lambda n, d: n - d * math.floor(n / d), lambda n, d: d != 0, DIV0)


def fn_POWER(a, b):
    return op_pow(sc(a), sc(b))


# ------------------------------------------------------------------- criteria

def _wildcard(pattern: str) -> re.Pattern:
    out, i = [], 0
    while i < len(pattern):
        ch = pattern[i]
        if ch == "~" and i + 1 < len(pattern) and pattern[i + 1] in "*?~":
            out.append(re.escape(pattern[i + 1]))
            i += 2
            continue
        out.append(".*" if ch == "*" else "." if ch == "?" else re.escape(ch))
        i += 1
    return re.compile("".join(out), re.I | re.S)


_MONTHS = "jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec"
# Criterion text Excel might parse as a number, date or time: starts with a
# digit, sign, '$', '.', '(' or a month name followed by a digit somewhere.
_RE_DATELIKE = re.compile(rf"\s*(?:[\d+\-$.(]|(?:{_MONTHS})[a-z]*\.?\s*\d)", re.I)


def _criterion(crit):
    """Compile a SUMIF-style criterion into a predicate over cell values."""
    crit = sc(crit)
    if isinstance(crit, XlError):
        return lambda v: v == crit
    op, explicit = "=", False
    if isinstance(crit, str):
        m = re.match(r"<=|>=|<>|<|>|=", crit)
        if m:
            op, crit, explicit = m.group(), crit[m.end():], True
        if crit.upper() in _ERRORS:
            err = _ERRORS[crit.upper()]
            if op == "=":
                return lambda v: v == err
            if op == "<>":
                return lambda v: v != err
            raise Unsupported(f"criterion compares an error with {op}")
        if _RE_NUM.fullmatch(crit.strip()):
            crit = float(crit.strip())
        elif crit.upper() in ("TRUE", "FALSE"):
            crit = crit.upper() == "TRUE"
        elif _RE_DATELIKE.match(crit):
            m_text = crit
            try:
                crit = text_to_num(crit)
            except Unsupported:
                raise Unsupported(f"criterion {crit!r} may be a date, time or formatted number") from None
            if isinstance(crit, XlError):  # no digits: an ordinary text criterion
                crit = m_text
    elif crit is None:
        crit = 0.0  # an empty criterion cell means 0

    if isinstance(crit, float):
        def as_number(v):
            if isinstance(v, float):
                return v
            if isinstance(v, str) and op in ("=", "<>") and _RE_NUM.fullmatch(v.strip()):
                return float(v)
            return None
        tests = {"=": lambda n: n == crit, "<>": lambda n: n != crit, "<": lambda n: n < crit,
                 ">": lambda n: n > crit, "<=": lambda n: n <= crit, ">=": lambda n: n >= crit}
        test = tests[op]
        if op == "<>":
            return lambda v: as_number(v) is None or test(as_number(v))
        return lambda v: as_number(v) is not None and test(as_number(v))
    if isinstance(crit, bool):
        if op == "<>":
            return lambda v: not (isinstance(v, bool) and v == crit)
        bool_tests = {"=": lambda c: c == 0, "<": lambda c: c < 0, ">": lambda c: c > 0,
                      "<=": lambda c: c <= 0, ">=": lambda c: c >= 0}
        bool_test = bool_tests[op]
        return lambda v: isinstance(v, bool) and bool_test(compare(v, crit))
    if op in ("=", "<>"):
        if crit == "":
            if op == "<>":
                return lambda v: v is not None  # "<>": anything not truly empty
            if explicit:
                return lambda v: v is None  # "=": truly empty cells only
            return lambda v: v is None or v == ""  # "": empty cells and empty text
        pattern = _wildcard(crit)
        if op == "=":
            return lambda v: isinstance(v, str) and pattern.fullmatch(v) is not None
        return lambda v: not (isinstance(v, str) and pattern.fullmatch(v) is not None)
    text_order_check(crit)
    key = fold(crit)
    tests = {"<": lambda s: s < key, ">": lambda s: s > key, "<=": lambda s: s <= key, ">=": lambda s: s >= key}
    test = tests[op]

    def check(v):
        if not isinstance(v, str):
            return False
        text_order_check(v)
        return test(fold(v))
    return check


def _matching(pairs, target=None):
    """Values of `target` (or of the first range) where every criterion holds."""
    ranges = [(_arr(r), _criterion(c)) for r, c in pairs]
    shape = ranges[0][0].shape
    target = _arr(target) if target is not None else ranges[0][0]
    clamped = any(r.clamped for r, _ in ranges) or target.clamped
    if clamped:
        if not all(r.clamped for r, _ in ranges) or any(r.shape != shape for r, _ in ranges):
            raise Unsupported("criteria over very large ranges of different sizes")
        if all(test(r.pad) for r, test in ranges):
            raise Unsupported("criteria that match empty cells, over a very large range")
    if any(r.shape != shape for r, _ in ranges):
        return VALUE
    if target.shape != shape:
        raise Unsupported("criteria range and value range have different shapes")
    columns = [list(r.flat()) for r, _ in ranges]
    out = []
    for i, v in enumerate(target.flat()):
        if all(test(col[i]) for col, (_, test) in zip(columns, ranges)):
            out.append(v)
    return out


def _if_family(reduce, numbers_only=True):
    def fn(vals):
        if isinstance(vals, XlError):
            return vals
        if numbers_only:
            e = _err(*vals)
            if e:
                return e
            vals = [v for v in vals if isinstance(v, float)]
        return reduce(vals)
    return fn


_sum_matches = _if_family(_sum)
_avg_matches = _if_family(lambda nums: _check(_sum(nums) / len(nums)) if nums else DIV0)
_max_matches = _if_family(lambda nums: max(nums) if nums else 0.0)
_min_matches = _if_family(lambda nums: min(nums) if nums else 0.0)
_count_matches = _if_family(lambda vals: float(len(vals)), numbers_only=False)


def _pairs(args):
    if len(args) % 2:
        raise Unsupported("criteria arguments must come in range/criterion pairs")
    return list(zip(args[::2], args[1::2]))


def fn_SUMIF(rng, crit, sum_range=None):
    return _sum_matches(_matching([(rng, crit)], sum_range))


def fn_AVERAGEIF(rng, crit, avg_range=None):
    return _avg_matches(_matching([(rng, crit)], avg_range))


def fn_COUNTIF(rng, crit):
    return _count_matches(_matching([(rng, crit)]))


def fn_SUMIFS(sum_range, *args):
    return _sum_matches(_matching(_pairs(args), sum_range))


def fn_AVERAGEIFS(avg_range, *args):
    return _avg_matches(_matching(_pairs(args), avg_range))


def fn_COUNTIFS(*args):
    return _count_matches(_matching(_pairs(args)))


def fn_MAXIFS(max_range, *args):
    return _max_matches(_matching(_pairs(args), max_range))


def fn_MINIFS(min_range, *args):
    return _min_matches(_matching(_pairs(args), min_range))


# --------------------------------------------------------------------- lookup

def _same(a, b) -> bool:
    if a is None or b is None or isinstance(a, XlError) or isinstance(b, XlError):
        return False
    return _rank(a) == _rank(b) and compare(a, b) == 0


def _find_exact(needle, values) -> int | None:
    if isinstance(needle, str) and any(ch in needle for ch in "*?~"):
        pattern = _wildcard(needle)
        return next((i for i, v in enumerate(values) if isinstance(v, str) and pattern.fullmatch(v)), None)
    return next((i for i, v in enumerate(values) if _same(needle, v)), None)


def _find_sorted(needle, values, descending=False) -> int | None:
    """Approximate match, reproducing Excel's search so unsorted data gives
    Excel's answer too.

    Ascending (MATCH type 1, VLOOKUP/HLOOKUP TRUE): binary search that returns
    at the first exact hit and, when the middle value is of another type,
    uses the next comparable value to its right. Descending (MATCH type -1):
    a scan from the top that stops at the first smaller value. Both rules were
    fitted to, and agree with, all 531 approximate lookups in the Excel-saved
    files of the test corpus (scripts/fetch_corpus.sh).
    """
    def comparable(v):
        return not (v is None or isinstance(v, XlError) or _rank(v) != _rank(needle))

    if descending:
        best = None
        for i, v in enumerate(values):
            if not comparable(v):
                continue
            c = order(v, needle)
            if c < 0:
                break
            best = i
            if c == 0:
                break
        return best
    lo, hi, best = 0, len(values) - 1, None
    while lo <= hi:
        mid = (lo + hi) // 2
        j = mid
        while j <= hi and not comparable(values[j]):
            j += 1
        if j > hi:
            hi = mid - 1
            continue
        c = order(values[j], needle)
        if c == 0:
            return j
        if c < 0:
            best = j
            lo = j + 1
        else:
            hi = mid - 1
    return best


def _vector(arr: Arr):
    rows, cols = arr.shape
    if rows == 1:
        return arr.rows[0]
    if cols == 1:
        return [r[0] for r in arr.rows]
    return None


def _index_arg(v):
    n = to_num(sc(v))
    return n if isinstance(n, XlError) else int(n)


def _lookup(needle, table, index, approx, vertical):
    needle = sc(needle)
    table = _arr(table)
    index = _index_arg(index)
    approx = to_bool(sc(approx))
    e = _err(needle, index, approx)
    if e:
        return e
    rows, cols = table.shape
    if index < 1:
        return VALUE
    if index > (cols if vertical else rows):
        if table.clamped:
            raise Unsupported("lookup index points beyond the used area of a very large range")
        return REF
    keys = [r[0] for r in table.rows] if vertical else table.rows[0]
    if needle is None:
        needle = 0.0
    pos = _find_sorted(needle, keys) if approx else _find_exact(needle, keys)
    if pos is None:
        return NA
    return table.rows[pos][index - 1] if vertical else table.rows[index - 1][pos]


def fn_VLOOKUP(needle, table, index, approx=True):
    return _lookup(needle, table, index, approx, True)


def fn_HLOOKUP(needle, table, index, approx=True):
    return _lookup(needle, table, index, approx, False)


def fn_MATCH(needle, rng, match_type=1.0):
    needle = sc(needle)
    mode = to_num(sc(match_type))
    e = _err(needle, mode)
    if e:
        return e
    values = _vector(_arr(rng))
    if values is None:
        return NA
    if needle is None:
        needle = 0.0
    pos = _find_exact(needle, values) if mode == 0 else _find_sorted(needle, values, descending=mode < 0)
    return NA if pos is None else float(pos + 1)


def fn_INDEX(rng, row=None, col=None):
    arr = _arr(rng)
    rows, cols = arr.shape
    if col is None and rows == 1 and cols > 1:
        row, col = 1.0, row
    r = 0 if row is None else _index_arg(row)
    c = 0 if col is None else _index_arg(col)
    e = _err(r, c)
    if e:
        return e
    if r < 0 or c < 0:
        return VALUE
    if r > rows or c > cols:
        if arr.clamped:
            raise Unsupported("INDEX points beyond the used area of a very large range")
        return REF
    if r == 0 and rows == 1:
        r = 1
    if c == 0 and cols == 1:
        c = 1
    if r and c:
        return arr.rows[r - 1][c - 1]
    if r:
        return Arr([arr.rows[r - 1]], arr.clamped, arr.pad)
    if c:
        return Arr([[row_[c - 1]] for row_ in arr.rows], arr.clamped, arr.pad)
    return arr


def fn_CHOOSE(index, *values):
    i = _index_arg(index)
    if isinstance(i, XlError):
        return i
    return values[i - 1] if 1 <= i <= len(values) else VALUE


def fn_XLOOKUP(needle, lookup, result, if_not_found=NA, match_mode=0.0, search_mode=1.0):
    if to_num(sc(match_mode)) != 0 or to_num(sc(search_mode)) != 1:
        raise Unsupported("XLOOKUP match/search modes other than the defaults")
    needle = sc(needle)
    if isinstance(needle, XlError):
        return needle
    keys = _vector(_arr(lookup))
    result = _arr(result)
    if keys is None:
        return VALUE
    if needle is None and _arr(lookup).clamped:
        raise Unsupported("XLOOKUP of an empty value in a very large range")
    pos = next((i for i, v in enumerate(keys) if _same(needle, v) or (needle is None and v is None)), None)
    if pos is None:
        return if_not_found
    values = _vector(result)
    if values is None or len(values) != len(keys):
        raise Unsupported("XLOOKUP returning a whole row or column")
    return values[pos]


# ----------------------------------------------------------------------- text

def _textual(f):
    def fn(*args):
        vals = [sc(a) for a in args]
        return _err(*vals) or f(*vals)
    return fn


class _Omitted:
    """Default for an optional argument the formula leaves out. An empty cell
    passed explicitly arrives as None and counts as 0, as in Excel."""
    def __repr__(self):
        return "OMITTED"


OMITTED = _Omitted()


def _count_arg(n, default=1.0):
    n = default if n is OMITTED else to_num(n)
    return n if isinstance(n, XlError) else int(n)


def _left(t, n=OMITTED):
    n = _count_arg(n)
    return n if isinstance(n, XlError) else VALUE if n < 0 else to_text(t)[:n]


def _right(t, n=OMITTED):
    n = _count_arg(n)
    if isinstance(n, XlError):
        return n
    t = to_text(t)
    return VALUE if n < 0 else t[len(t) - n:] if n < len(t) else t


def _mid(t, start, n):
    start, n = _count_arg(start), _count_arg(n)
    e = _err(start, n)
    if e:
        return e
    return VALUE if start < 1 or n < 0 else to_text(t)[start - 1:start - 1 + n]


def _find(needle, hay, start=OMITTED, fold=False):
    start = _count_arg(start)
    if isinstance(start, XlError):
        return start
    needle, hay = to_text(needle), to_text(hay)
    if fold:
        if any(ch in needle for ch in "*?~"):
            raise Unsupported("SEARCH with wildcards")
        needle, hay = needle.upper(), hay.upper()
    if start < 1 or start > len(hay) + 1:
        return VALUE
    pos = hay.find(needle, start - 1)
    return VALUE if pos < 0 else float(pos + 1)


def _substitute(t, old, new, instance=OMITTED):
    t, old, new = to_text(t), to_text(old), to_text(new)
    if not old:
        return t
    if instance is OMITTED:
        return t.replace(old, new)
    n = _count_arg(instance)
    if isinstance(n, XlError):
        return n
    if n < 1:
        return VALUE
    pos = -1
    for _ in range(n):
        pos = t.find(old, pos + 1)
        if pos < 0:
            return t
    return t[:pos] + new + t[pos + len(old):]


def _value(t):
    return VALUE if isinstance(t, bool) else to_num(t)


def _rept(t, n):
    n = _count_arg(n)
    return n if isinstance(n, XlError) else VALUE if n < 0 else to_text(t) * n


fn_LEN = _textual(lambda t: float(len(to_text(t))))
fn_LEFT = _textual(_left)
fn_RIGHT = _textual(_right)
fn_MID = _textual(_mid)
fn_UPPER = _textual(lambda t: to_text(t).upper())
fn_LOWER = _textual(lambda t: to_text(t).lower())
fn_TRIM = _textual(lambda t: re.sub(" +", " ", to_text(t)).strip(" "))
fn_EXACT = _textual(lambda a, b: to_text(a) == to_text(b))
fn_FIND = _textual(_find)
fn_SEARCH = _textual(lambda n, h, s=OMITTED: _find(n, h, s, fold=True))
fn_SUBSTITUTE = _textual(_substitute)
fn_VALUE = _textual(_value)
fn_REPT = _textual(_rept)
fn_CONCATENATE = _textual(lambda *parts: "".join(to_text(p) for p in parts))


def fn_CONCAT(*args):
    out = []
    for a in args:
        for v in (a.flat() if isinstance(a, Arr) else (a,)):
            if isinstance(v, XlError):
                return v
            out.append(to_text(v))
    return "".join(out)


# ---------------------------------------------------------------------- dates

_EPOCH = date(1899, 12, 30)


def _serial(d: date) -> float:
    n = (d - _EPOCH).days
    return float(n if n > 60 else n - 1)  # Excel counts a 29 Feb 1900 that never existed


def _from_serial(n: float):
    n = math.floor(n)
    if n < 0 or n > 2958465:
        return None
    if n == 60:
        return (1900, 2, 29)
    if n == 0:
        return (1900, 1, 0)
    d = _EPOCH + timedelta(days=n if n > 60 else n + 1)
    return (d.year, d.month, d.day)


def _date(y, m, d):
    y, m, d = int(y), int(m), int(d)
    if 0 <= y < 1900:
        y += 1900
    y += (m - 1) // 12
    m = (m - 1) % 12 + 1
    if not 1900 <= y <= 9999:
        return NUM
    if (y, m) < (1900, 3):
        raise Unsupported("dates before March 1900")
    return _serial(date(y, m, 1) + timedelta(days=d - 1))


def _part(i):
    def f(n):
        parts = _from_serial(n)
        return NUM if parts is None else float(parts[i])
    return f


fn_DATE = _numeric(_date)
fn_YEAR = _numeric(_part(0))
fn_MONTH = _numeric(_part(1))
fn_DAY = _numeric(_part(2))


# ------------------------------------------------------------------ financial

def _pmt(rate, nper, pv, fv=0.0, due=0.0):
    if nper == 0:
        return NUM
    if rate == 0:
        return -(pv + fv) / nper
    g = (1 + rate) ** nper
    return -(rate * (pv * g + fv)) / ((1 + rate * (1 if due else 0)) * (g - 1))


def _fv(rate, nper, pmt, pv=0.0, due=0.0):
    if rate == 0:
        return -(pv + pmt * nper)
    g = (1 + rate) ** nper
    return -(pv * g + pmt * (1 + rate * (1 if due else 0)) * (g - 1) / rate)


def _pv(rate, nper, pmt, fv=0.0, due=0.0):
    if rate == 0:
        return -(fv + pmt * nper)
    g = (1 + rate) ** nper
    return -(fv + pmt * (1 + rate * (1 if due else 0)) * (g - 1) / rate) / g


def _annuity(f):
    """PMT/FV/PV: a 'type' other than 0 or 1 is declined (Excel's handling is unconfirmed)."""
    def fn(rate, nper, x, y=0.0, due=0.0):
        if not isinstance(due, float) or due not in (0.0, 1.0):
            raise Unsupported("payment type other than 0 or 1")
        return f(rate, nper, x, y, due)
    return _domain(fn, lambda *a: True)


fn_PMT = _annuity(_pmt)
fn_FV = _annuity(_fv)
fn_PV = _annuity(_pv)


def fn_NPV(rate, *values):
    rate = to_num(sc(rate))
    if isinstance(rate, XlError):
        return rate
    nums = _numbers(values)
    if isinstance(nums, XlError):
        return nums
    if rate == -1:
        return DIV0
    total = 0.0
    for i, n in enumerate(nums, 1):
        total += n / (1 + rate) ** i
    return _check(total)


FUNCTIONS = {name[3:]: obj for name, obj in list(globals().items()) if name.startswith("fn_")}

# Present in Excel but deliberately not compiled: the result depends on when or
# where the workbook is calculated, so there is nothing fixed to verify.
VOLATILE = {"NOW", "TODAY", "RAND", "RANDBETWEEN", "RANDARRAY", "OFFSET", "INDIRECT", "CELL", "INFO"}


# ------------------------------------------------------------------ evaluator

@lru_cache(maxsize=None)
def _parse_range(key: str):
    sheet, ref = key.rsplit("!", 1)
    first, _, last = ref.partition(":")
    r1, c1 = split_a1(first)
    r2, c2 = split_a1(last or first)
    return sheet, min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)


class Context:
    def __init__(self, inputs: dict, formulas: dict, unsupported: dict):
        self.values = dict(inputs)
        self.formulas = formulas
        self.failed = dict(unsupported)
        self.busy: set[str] = set()
        self.ranges: dict[str, Arr] = {}

    def cell(self, key: str):
        try:
            return self.values[key]
        except KeyError:
            pass
        f = self.formulas.get(key)
        if key in self.failed:
            raise Unsupported(_root_cause(key, self.failed[key]))
        if f is None:
            return None
        if key in self.busy:
            raise Unsupported("circular reference")
        self.busy.add(key)
        try:
            result = _finish(f(self))
        except Unsupported as exc:
            self.failed[key] = str(exc)
            raise Unsupported(_root_cause(key, str(exc))) from None
        except Exception as exc:  # includes RecursionError on an extreme chain  # a gap in the runtime must never become a wrong value
            self.failed[key] = f"runtime could not evaluate this formula ({type(exc).__name__}: {exc})"
            raise Unsupported(_root_cause(key, self.failed[key])) from None
        finally:
            self.busy.discard(key)
        self.values[key] = result
        return result

    def ref(self, key: str) -> Arr:
        """A single-cell reference passed to a function keeps its 'range-ness'."""
        return Arr([[self.cell(key)]])

    def rng(self, key: str, clamped: bool = False) -> Arr:
        arr = self.ranges.get((key, clamped))
        if arr is None:
            sheet, r1, c1, r2, c2 = _parse_range(key)
            if (r2 - r1 + 1) * (c2 - c1 + 1) > MAX_RANGE_CELLS:
                raise Unsupported(f"range {key} has more than {MAX_RANGE_CELLS:,} cells")
            cell = self.cell
            arr = Arr([[cell(f"{sheet}!{a1(r, c)}") for c in range(c1, c2 + 1)] for r in range(r1, r2 + 1)], clamped)
            self.ranges[(key, clamped)] = arr
        return arr


def _root_cause(key: str, reason: str) -> str:
    return reason if reason.startswith("depends on ") else f"depends on {key}, which is unverified: {reason}"


def _finish(v):
    if isinstance(v, Arr):
        if not v.single:
            raise Unsupported("formula returns an array (spill)")
        v = v.rows[0][0]
    if v is None:
        return 0.0
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        v = float(v)
        return v if math.isfinite(v) else NUM
    return v


MAX_RANGE_CELLS = 10_000_000


def dependency_graph(keys, deps: dict) -> dict[str, list[str]]:
    """For each cell in `keys`: the cells of `keys` it reads, ranges expanded."""
    keys = set(keys)
    by_column: dict[tuple[str, int], list[tuple[int, str]]] = {}
    for key in keys:
        sheet, ref = key.rsplit("!", 1)
        row, col = split_a1(ref)
        by_column.setdefault((sheet, col), []).append((row, key))
    columns_by_sheet: dict[str, list[int]] = {}
    for (sheet, col), cells in by_column.items():
        cells.sort()
        columns_by_sheet.setdefault(sheet, []).append(col)

    graph = {}
    for key in keys:
        found = []
        for dep in deps.get(key, ()):
            if ":" not in dep:
                if dep in keys:
                    found.append(dep)
                continue
            sheet, r1, c1, r2, c2 = _parse_range(dep)
            for col in columns_by_sheet.get(sheet, ()):
                if c1 <= col <= c2:
                    cells = by_column[(sheet, col)]
                    found.extend(k for _, k in cells[bisect_left(cells, (r1, "")):bisect_right(cells, (r2, "\uffff"))])
        graph[key] = found
    return graph


def _order(formulas: dict, deps: dict) -> list[str]:
    """Formula cells, each after the formula cells it reads (cycles left as found).

    Evaluating in this order keeps recursion shallow however long a dependency
    chain is; without it, a chain whose first cell depends on its last would
    recurse once per cell.
    """
    graph = dependency_graph(formulas, deps)
    done: set[str] = set()
    out: list[str] = []
    for start in formulas:
        if start in done:
            continue
        done.add(start)
        stack = [(start, iter(graph[start]))]
        while stack:
            key, pending = stack[-1]
            for dep in pending:
                if dep not in done:
                    done.add(dep)
                    stack.append((dep, iter(graph[dep])))
                    break
            else:
                stack.pop()
                out.append(key)
    return out


def circular_cells(deps: dict) -> set[str]:
    """Formula cells on a circular reference, or reading (directly or not) from one.

    `deps` covers every formula cell, compiled or not. Excel leaves unresolved
    circular cells, and cells computed from them, holding 0 or a partial
    iteration, so their stored values are not answers to check against.
    """
    graph = dependency_graph(deps, deps)
    index: dict[str, int] = {}
    low: dict[str, int] = {}
    on_stack: set[str] = set()
    stack: list[str] = []
    cyclic: set[str] = set()
    counter = 0
    for root_key in graph:
        if root_key in index:
            continue
        work = [(root_key, iter(graph[root_key]))]
        index[root_key] = low[root_key] = counter
        counter += 1
        stack.append(root_key)
        on_stack.add(root_key)
        while work:
            key, children = work[-1]
            advanced = False
            for child in children:
                if child not in index:
                    index[child] = low[child] = counter
                    counter += 1
                    stack.append(child)
                    on_stack.add(child)
                    work.append((child, iter(graph[child])))
                    advanced = True
                    break
                if child in on_stack:
                    low[key] = min(low[key], index[child])
            if advanced:
                continue
            work.pop()
            if work:
                parent = work[-1][0]
                low[parent] = min(low[parent], low[key])
            if low[key] == index[key]:
                component = []
                while True:
                    k = stack.pop()
                    on_stack.discard(k)
                    component.append(k)
                    if k == key:
                        break
                if len(component) > 1 or key in graph[key]:
                    cyclic.update(component)
    readers: dict[str, list[str]] = {}
    for key, children in graph.items():
        for child in children:
            readers.setdefault(child, []).append(key)
    todo = list(cyclic)
    while todo:
        for reader in readers.get(todo.pop(), ()):
            if reader not in cyclic:
                cyclic.add(reader)
                todo.append(reader)
    return cyclic


def evaluate(inputs: dict, formulas: dict, unsupported: dict, overrides: dict | None = None, deps: dict | None = None):
    """Recalculate a compiled workbook.

    Returns (values, failed): `values` maps each computed formula cell to its
    result, `failed` maps every formula cell that could not be computed to the
    reason. `deps` (cell -> cells and ranges it reads) lets cells be computed
    in dependency order.
    """
    merged = dict(inputs)
    for key, value in (overrides or {}).items():
        if key in formulas or key in unsupported:
            raise ValueError(f"{key} is a formula cell and cannot be overridden")
        merged[key] = float(value) if isinstance(value, int) and not isinstance(value, bool) else value
    ctx = Context(merged, formulas, unsupported)
    limit = sys.getrecursionlimit()
    sys.setrecursionlimit(max(limit, 20_000))
    try:
        for key in (_order(formulas, deps) if deps else formulas):
            try:
                ctx.cell(key)
            except Unsupported:
                pass
    finally:
        sys.setrecursionlimit(limit)
    return {k: ctx.values[k] for k in formulas if k in ctx.values}, ctx.failed
