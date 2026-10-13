"""MCP server: gives Claude (and any MCP client) verified access to Excel workbooks.

    provenance-mcp                         # stdio, for Claude Code / Claude Desktop; reads local paths
    provenance-mcp --http [--port 8000]    # remote connector; files arrive by URL or base64 only
    provenance-mcp --config                # print the client config snippet

Every tool works on one workbook, given as exactly one of `path` (local mode
only), `url` (https; Google Sheets share links are exported as .xlsx) or
`content_base64`. Results are the same numbers the CLI certifies: a formula
whose behaviour Provenance does not fully model is reported as unverified,
never guessed.

In HTTP mode the server never touches its own disk on a caller's behalf, refuses
URLs that resolve to private addresses, and runs each conversion in a separate
process with a time and memory limit.
"""

from __future__ import annotations

import argparse
import base64
import inspect
import binascii
import ipaddress
import json
import os
import re
import secrets
import shutil
import socket
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from collections import OrderedDict
from pathlib import Path
from typing import Annotated, Any, Literal

from . import __version__
from .canon import ConversionError, sha256

MAX_BYTES = 25 * 1024 * 1024
FETCH_TIMEOUT = 30
WORKER_TIMEOUT = int(os.environ.get("PROVENANCE_MCP_TIMEOUT", "120"))
WORKER_MEMORY = int(os.environ.get("PROVENANCE_MCP_MEMORY_MB", "2048")) << 20
WORKER_SLOTS = threading.BoundedSemaphore(int(os.environ.get("PROVENANCE_MCP_WORKERS", "2")))
MAX_CELLS_RETURNED = 500
SUFFIXES = (".xlsx", ".xlsm", ".docx")
CELL_KEY = re.compile(r"^(?:'(?:[^']|'')+'|[^!'\s][^!]*)!\$?[A-Za-z]{1,3}\$?[0-9]{1,7}$")

INSTRUCTIONS = """\
Provenance reads Excel workbooks (.xlsx/.xlsm) exactly and recalculates them in Python with
Excel's semantics. Start with inspect_workbook (or find_cells in a big workbook) to learn the
layout, then: explain_cells / trace_cells for how a number is calculated; recalculate_workbook
for what-if questions; goal_seek to hit a target; sensitivity_table to see how outputs respond
to an input; compare_workbooks for what changed between versions; convert_workbook for the
certified JSON + Python conversion. Cell keys look like "Sheet1!B2". A cell whose status is
unverified, mismatch or disputed must not be presented as Excel's answer; say it is unverified."""

UPLOAD_HINT = """
If the user has a spreadsheet in this conversation but no public link, ask them to upload it at
{url}/upload and paste back the file_id it shows (valid for 1 hour). Never guess file contents."""


class ToolError(Exception):
    """A problem the caller can fix; reported to the model as a tool error."""


# ------------------------------------------------------------------ sources

def _google_export(url: str) -> str:
    m = re.match(r"https://docs\.google\.com/spreadsheets/d/([A-Za-z0-9_-]+)", url)
    return f"https://docs.google.com/spreadsheets/d/{m.group(1)}/export?format=xlsx" if m else url


def _check_host(host: str, remote: bool) -> None:
    if not remote:
        return
    try:
        infos = socket.getaddrinfo(host, 443, proto=socket.IPPROTO_TCP)
    except socket.gaierror as exc:
        raise ToolError(f"cannot resolve {host}: {exc}") from None
    for info in infos:
        ip = ipaddress.ip_address(info[4][0])
        if not ip.is_global:
            raise ToolError(f"{host} resolves to a non-public address; only public URLs can be fetched")


class _NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, *args, **kwargs):
        return None


def _fetch(url: str, remote: bool) -> tuple[bytes, str]:
    url = _google_export(url.strip())
    opener = urllib.request.build_opener(_NoRedirect)
    for _ in range(5):
        parts = urllib.parse.urlsplit(url)
        if parts.scheme != "https" and (remote or parts.scheme != "http"):
            raise ToolError("url must start with https://")
        _check_host(parts.hostname or "", remote)
        request = urllib.request.Request(url, headers={"User-Agent": f"provenance-mcp/{__version__}"})
        try:
            with opener.open(request, timeout=FETCH_TIMEOUT) as response:
                data = response.read(MAX_BYTES + 1)
                if len(data) > MAX_BYTES:
                    raise ToolError(f"file is larger than {MAX_BYTES >> 20} MB")
                name = Path(urllib.parse.unquote(parts.path)).name
                disposition = response.headers.get("Content-Disposition", "")
                if m := re.search(r'filename="?([^";]+)', disposition):
                    name = m.group(1)
                return data, name or "workbook.xlsx"
        except urllib.error.HTTPError as exc:
            if exc.code in (301, 302, 303, 307, 308) and exc.headers.get("Location"):
                url = urllib.parse.urljoin(url, exc.headers["Location"])
                continue
            raise ToolError(f"fetching the url failed: HTTP {exc.code}") from None
        except urllib.error.URLError as exc:
            raise ToolError(f"fetching the url failed: {exc.reason}") from None
    raise ToolError("too many redirects")


# Upload store (hosted mode): a person drops a file on /upload and gives Claude the file_id.
# Files are never served back, only read by tool calls, and are deleted after UPLOAD_TTL.
UPLOADS = Path(os.environ.get("PROVENANCE_UPLOAD_DIR", Path(tempfile.gettempdir()) / "provenance-uploads"))
UPLOAD_TTL = 3600
UPLOAD_STORE_MAX = int(os.environ.get("PROVENANCE_UPLOAD_STORE_MB", "300")) << 20
FILE_ID = re.compile(r"^pv_[A-Za-z0-9_-]{24,48}$")
PUBLIC_URL = os.environ.get("PROVENANCE_PUBLIC_URL") or (
    f"https://{os.environ['RAILWAY_PUBLIC_DOMAIN']}" if os.environ.get("RAILWAY_PUBLIC_DOMAIN") else "")
_UPLOAD_LOCK = threading.Lock()
UPLOADS_PER_HOUR = int(os.environ.get("PROVENANCE_UPLOADS_PER_HOUR", "30"))
_UPLOAD_TIMES: dict[str, list[float]] = {}


def _sweep_uploads() -> int:
    """Delete expired uploads; return the bytes still stored."""
    total, cutoff = 0, time.time() - UPLOAD_TTL
    for p in UPLOADS.glob("pv_*") if UPLOADS.is_dir() else ():
        try:
            if p.stat().st_mtime < cutoff:
                p.unlink()
            else:
                total += p.stat().st_size
        except FileNotFoundError:
            pass
    return total


def store_upload(data: bytes, name: str) -> str:
    if not data.startswith(b"PK\x03\x04"):
        raise ToolError("that is not an Excel workbook (.xlsx / .xlsm)")
    if len(data) > MAX_BYTES:
        raise ToolError(f"file is larger than {MAX_BYTES >> 20} MB")
    with _UPLOAD_LOCK:
        UPLOADS.mkdir(parents=True, exist_ok=True, mode=0o700)
        if _sweep_uploads() + len(data) > UPLOAD_STORE_MAX:
            raise ToolError("the upload store is full right now; try again in a few minutes")
        file_id = "pv_" + secrets.token_urlsafe(24)
        (UPLOADS / file_id).write_bytes(data)
        (UPLOADS / f"{file_id}.name").write_text(Path(name).name[:200] or "workbook.xlsx", encoding="utf-8")
    return file_id


def load_upload(file_id: str) -> tuple[bytes, str]:
    path = UPLOADS / file_id.strip()
    if not FILE_ID.match(file_id.strip()) or not path.is_file() or path.stat().st_mtime < time.time() - UPLOAD_TTL:
        where = f" at {PUBLIC_URL}/upload" if PUBLIC_URL else ""
        raise ToolError(f"file_id {file_id!r} is unknown or has expired (uploads last 1 hour); "
                        f"ask the user to upload the file again{where}")
    name_file = path.with_name(f"{path.name}.name")
    return path.read_bytes(), name_file.read_text(encoding="utf-8") if name_file.exists() else "workbook.xlsx"


SOURCE_KEYS = ("path", "file_id", "url", "content_base64")


def load_source(src: dict, remote: bool) -> tuple[bytes, str]:
    """Return (bytes, file name) for the one source in src (path / file_id / url / content_base64)."""
    path, file_id, url, content_base64 = (src.get(k) for k in SOURCE_KEYS)
    filename = src.get("filename")
    given = [s for s in (path, file_id, url, content_base64) if s]
    if len(given) != 1:
        options = "file_id, url or content_base64" if remote else "path, url or content_base64"
        raise ToolError(f"give exactly one of {options}")
    if file_id:
        if not remote:
            raise ToolError("file_id is for the hosted server; locally, give the file's path")
        return load_upload(file_id)
    if path:
        if remote:
            raise ToolError("this is a hosted server: send the file as url or content_base64, not a path")
        p = Path(path).expanduser()
        if not p.is_file():
            raise ToolError(f"no such file: {p}")
        if p.stat().st_size > MAX_BYTES:
            raise ToolError(f"file is larger than {MAX_BYTES >> 20} MB")
        return p.read_bytes(), p.name
    if url:
        data, name = _fetch(url, remote)
        return data, filename or name
    try:
        data = base64.b64decode(content_base64, validate=True)
    except (binascii.Error, ValueError):
        raise ToolError("content_base64 is not valid base64") from None
    if len(data) > MAX_BYTES:
        raise ToolError(f"file is larger than {MAX_BYTES >> 20} MB")
    return data, filename or "workbook.xlsx"


# ---------------------------------------------------------------- analysis

def _plain(value):
    """JSON-safe form of a computed value (Excel errors become their code, e.g. '#DIV/0!')."""
    if isinstance(value, (bool, int, float, str)) or value is None:
        return value
    return repr(value)


def _stored(cell: dict | None):
    if not cell:
        return None
    return cell["e"] if "e" in cell else cell.get("v")


class Analysis:
    """One converted workbook: model, certificate, compiled logic."""

    def __init__(self, data: bytes, name: str, fuzz: int, seed: int):
        from .cli import convert
        from .compiler import load_module

        safe = re.sub(r"[^0-9A-Za-z._-]", "_", Path(name).name) or "workbook.xlsx"
        if not safe.lower().endswith(SUFFIXES):
            safe += ".xlsx"
        work = Path(tempfile.mkdtemp(prefix="provenance-mcp-"))
        try:
            src = work / safe
            src.write_bytes(data)
            self.certificate = convert(src, work / "out", fuzz, seed, "auto" if fuzz else "none")
            out = work / "out"
            stem = src.stem
            model = out / f"{stem}.workbook.json"
            self.model = json.loads(model.read_text(encoding="utf-8")) if model.exists() else None
            md = out / f"{stem}.md"
            self.markdown = md.read_text(encoding="utf-8") if md.exists() else None
            logic = list(out.glob("*_logic.py"))
            self.logic = logic[0].read_text(encoding="utf-8") if logic else None
            self.files = {p.name: p.read_bytes() for p in out.iterdir()}
        finally:
            shutil.rmtree(work, ignore_errors=True)
        self.name = name
        self._namespace = None
        self._load_module = load_module

    @property
    def namespace(self):
        if self._namespace is None:
            self._namespace = self._load_module(self.logic)
        return self._namespace

    def cell(self, key: str) -> dict | None:
        sheet_name, ref = key.rsplit("!", 1)
        sheet_name = sheet_name.strip("'").replace("''", "'")
        for sheet in self.model["sheets"]:
            if sheet["name"].lower() == sheet_name.lower():
                return sheet["cells"].get(ref.replace("$", "").upper())
        raise ToolError(f"no sheet named {sheet_name!r}; sheets are {[s['name'] for s in self.model['sheets']]}")

    def canonical(self, key: str) -> str:
        if not CELL_KEY.match(key.strip()):
            raise ToolError(f"{key!r} is not a cell key like 'Sheet1!B2'")
        sheet_name, ref = key.strip().rsplit("!", 1)
        sheet_name = sheet_name.strip("'").replace("''", "'")
        for sheet in self.model["sheets"]:
            if sheet["name"].lower() == sheet_name.lower():
                return f"{sheet['name']}!{ref.replace('$', '').upper()}"
        raise ToolError(f"no sheet named {sheet_name!r}; sheets are {[s['name'] for s in self.model['sheets']]}")

    def status(self, key: str) -> str | None:
        return self.certificate["layers"].get("formulas", {}).get("cells", {}).get(key)

    def reason(self, key: str) -> str | None:
        return self.certificate["layers"].get("formulas", {}).get("unverified", {}).get(key)

    def require_workbook(self) -> None:
        if self.model is None:
            raise ToolError("this file is not an Excel workbook; only convert_workbook applies to it")


_CACHE: OrderedDict[tuple, Analysis] = OrderedDict()


def analyse(data: bytes, name: str, fuzz: int = 0, seed: int = 0) -> Analysis:
    key = (sha256(data), fuzz, seed)
    if key in _CACHE:
        _CACHE.move_to_end(key)
        return _CACHE[key]
    try:
        result = Analysis(data, name, fuzz, seed)
    except ConversionError as exc:
        raise ToolError(f"cannot read this file: {exc}") from None
    _CACHE[key] = result
    while len(_CACHE) > 4:
        _CACHE.popitem(last=False)
    return result


def _verdict(cert: dict) -> dict:
    formulas = cert["layers"].get("formulas", {})
    reasons: dict[str, int] = {}
    for reason in formulas.get("unverified", {}).values():
        reason = re.sub(r"depends on \S+, which is unverified: ", "", reason)
        reasons[reason] = reasons.get(reason, 0) + 1
    out = {
        "verdict": cert["verdict"]["level"],
        "summary": cert["verdict"]["summary"],
        "structure": cert["layers"]["structure"]["status"],
        "formulas": {k: formulas[k] for k in ("status", "total", "counts") if k in formulas},
        "unverified_reasons": [{"reason": r, "cells": n} for r, n in sorted(reasons.items(), key=lambda kv: -kv[1])[:15]],
        "notes": cert["notes"][:15],
    }
    if "macros" in cert["layers"]:
        out["macros"] = cert["layers"]["macros"].get("status")
    fz = formulas.get("fuzz")
    if fz and fz.get("status") == "ran":
        out["fuzz"] = {"cases": fz["cases"], "engine": fz["engine"], "counts": fz["counts"]}
    mismatches = formulas.get("replay", {}).get("mismatches", [])
    if mismatches:
        out["mismatches"] = mismatches[:20]
    return out


# ---------------------------------------------------------------- operations
# Each takes (data, name, params) and returns a JSON-able dict, so it can run
# in this process (stdio) or in a sandboxed worker process (HTTP).

def op_inspect(data: bytes, name: str, p: dict) -> dict:
    a = analyse(data, name)
    a.require_workbook()
    from .model import to_markdown
    model = a.model
    if p.get("sheet"):
        wanted = [s for s in model["sheets"] if s["name"].lower() == p["sheet"].lower()]
        if not wanted:
            raise ToolError(f"no sheet named {p['sheet']!r}; sheets are {[s['name'] for s in model['sheets']]}")
        model = {**model, "sheets": wanted}
    md = to_markdown(model, name, sha256(data), max_rows=p.get("max_rows", 100), max_cols=p.get("max_cols", 26))
    return {
        "file": name,
        "sha256": sha256(data),
        "sheets": [{"name": s["name"], "state": s["state"], "rows": s["max_row"], "columns": s["max_col"],
                    "formulas": sum(1 for c in s["cells"].values() if "f" in c)} for s in a.model["sheets"]],
        "defined_names": a.model["defined_names"][:50],
        "has_macros": bool(a.model["macro_parts"]),
        "markdown": md,
    }


def op_convert(data: bytes, name: str, p: dict) -> dict:
    a = analyse(data, name, p.get("fuzz_cases", 25), p.get("seed", 0))
    out = {"file": name, "sha256": sha256(data), **_verdict(a.certificate)}
    if p.get("output_dir"):
        folder = Path(p["output_dir"]).expanduser()
        folder.mkdir(parents=True, exist_ok=True)
        for fname, blob in a.files.items():
            (folder / fname).write_bytes(blob)
        out["written"] = sorted(str(folder / f) for f in a.files)
    if p.get("include_certificate"):
        out["certificate"] = a.certificate
    if p.get("include_logic") and a.logic:
        out["logic_py"] = a.logic
    return out


def _coerce(value):
    if isinstance(value, str) and value.upper() in ("TRUE", "FALSE"):
        return value.upper() == "TRUE"
    if isinstance(value, str):
        try:
            return float(value.replace(",", "")) if re.fullmatch(r"\s*-?[\d,]*\.?\d+(?:[eE][-+]?\d+)?\s*", value) else value
        except ValueError:
            return value
    if isinstance(value, bool) or value is None:
        return value
    if isinstance(value, (int, float)):
        return float(value)
    raise ToolError(f"override values must be numbers, text or booleans, not {type(value).__name__}")


def op_recalculate(data: bytes, name: str, p: dict) -> dict:
    a = analyse(data, name)
    a.require_workbook()
    if not a.logic:
        raise ToolError("the workbook has no formulas to recalculate")
    overrides = {}
    for key, value in (p.get("overrides") or {}).items():
        k = a.canonical(key)
        cell = a.cell(k)
        if cell and "f" in cell:
            raise ToolError(f"{k} holds a formula; override the input cells it reads instead")
        overrides[k] = _coerce(value)
    base, _ = a.namespace["compute"]()
    values, failed = a.namespace["compute"](overrides)
    if p.get("cells"):
        keys = [a.canonical(k) for k in p["cells"]]
    else:
        keys = [k for k in values if _plain(values[k]) != _plain(base.get(k))]
    results, truncated = [], len(keys) > MAX_CELLS_RETURNED
    for k in keys[:MAX_CELLS_RETURNED]:
        cell = a.cell(k)
        entry: dict[str, Any] = {"cell": k}
        if k in overrides:
            entry.update(value=overrides[k], kind="input (overridden)")
        elif cell and "f" in cell:
            entry.update(formula="=" + cell["f"], value=_plain(values.get(k)), before=_plain(base.get(k)),
                         status=a.status(k))
            if k in failed or a.status(k) in ("unverified", "mismatch", "disputed"):
                entry["warning"] = a.reason(k) or failed.get(k) or "not verified: do not treat as Excel's answer"
        else:
            entry.update(value=_stored(cell), kind="input")
        results.append(entry)
    return {
        "file": name,
        "overrides": overrides,
        "changed_cells": sum(1 for k in values if _plain(values[k]) != _plain(base.get(k))),
        "results": results,
        "truncated": truncated,
        "verdict": a.certificate["verdict"]["level"],
    }


def op_explain(data: bytes, name: str, p: dict) -> dict:
    a = analyse(data, name)
    a.require_workbook()
    keys = [a.canonical(k) for k in p.get("cells") or []]
    if not keys:
        raise ToolError("give one or more cells, like ['Sheet1!D6']")
    values = a.namespace["compute"]()[0] if a.logic else {}
    out = []
    for k in keys[:100]:
        cell = a.cell(k)
        entry: dict[str, Any] = {"cell": k, "stored_value": _stored(cell)}
        if cell and "f" in cell:
            reads = a.namespace.get("DEPS", {}).get(k, ()) if a.logic else ()
            entry.update(formula="=" + cell["f"], recomputed=_plain(values.get(k)), status=a.status(k),
                         reads=list(reads))
            if a.reason(k):
                entry["unverified_because"] = a.reason(k)
        else:
            entry["kind"] = "input" if cell else "empty"
        out.append(entry)
    return {"file": name, "cells": out}


def _labels(sheet: dict, ref: str) -> dict:
    """The nearest text to the left (row label) and above (column header) of a cell."""
    from .formula import a1, split_a1
    row, col = split_a1(ref)
    out = {}
    for c in range(col - 1, max(0, col - 30), -1):
        v = sheet["cells"].get(a1(row, c), {}).get("v")
        if isinstance(v, str) and v.strip():
            out["row_label"] = v
            break
    for r in range(max(1, row - 200), row):  # the topmost text above: the column's header
        v = sheet["cells"].get(a1(r, col), {}).get("v")
        if isinstance(v, str) and v.strip():
            out["column_label"] = v
            break
    return out


def op_find(data: bytes, name: str, p: dict) -> dict:
    a = analyse(data, name)
    a.require_workbook()
    from .formula import split_a1
    query, mode, where = p["query"], p.get("match", "contains"), p.get("search_in", "both")
    if mode == "regex":
        try:
            pattern = re.compile(query, re.IGNORECASE)
        except re.error as exc:
            raise ToolError(f"bad regex: {exc}") from None
    number = None
    try:
        number = float(str(query).replace(",", ""))
    except ValueError:
        pass

    def hit(text) -> bool:
        if text is None:
            return False
        if isinstance(text, (int, float)) and not isinstance(text, bool) and number is not None and mode != "regex":
            return abs(text - number) <= 1e-9 * max(1.0, abs(number))
        s = str(text)
        if mode == "regex":
            return bool(pattern.search(s))
        if mode == "exact":
            return s.strip().lower() == str(query).strip().lower()
        return str(query).lower() in s.lower()

    found, total = [], 0
    for sheet in a.model["sheets"]:
        if p.get("sheet") and sheet["name"].lower() != p["sheet"].lower():
            continue
        for ref in sorted(sheet["cells"], key=split_a1):
            cell = sheet["cells"][ref]
            in_value = where in ("values", "both") and hit(_stored(cell))
            in_formula = where in ("formulas", "both") and "f" in cell and hit(cell["f"])
            if not (in_value or in_formula):
                continue
            total += 1
            if len(found) < p.get("limit", 50):
                k = f"{sheet['name']}!{ref}"
                entry = {"cell": k, "value": _stored(cell), **_labels(sheet, ref)}
                if "f" in cell:
                    entry.update(formula="=" + cell["f"], status=a.status(k))
                found.append(entry)
    return {"file": name, "query": query, "matches": total, "results": found, "truncated": total > len(found)}


def _parse_range(key: str):
    """'Sheet!A1:C9' or 'Sheet!B2' -> (sheet, r1, c1, r2, c2)."""
    from .formula import split_a1
    sheet, ref = key.rsplit("!", 1)
    first, _, last = ref.partition(":")
    r1, c1 = split_a1(first)
    r2, c2 = split_a1(last) if last else (r1, c1)
    return sheet, min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)


class _Graph:
    """Precedents and dependents of formula cells, from the compiled DEPS."""

    def __init__(self, a: Analysis):
        self.deps = a.namespace.get("DEPS", {}) if a.logic else {}
        self.single: dict[str, set] = {}
        self.ranges: list[tuple] = []
        for cell, reads in self.deps.items():
            for r in reads:
                if ":" in r:
                    self.ranges.append((_parse_range(r), cell))
                else:
                    self.single.setdefault(r, set()).add(cell)

    def dependents(self, key: str) -> list[str]:
        sheet, r, c, _, _ = _parse_range(key)
        out = set(self.single.get(key, ()))
        for (s, r1, c1, r2, c2), cell in self.ranges:
            if s == sheet and r1 <= r <= r2 and c1 <= c <= c2:
                out.add(cell)
        return sorted(out)

    def formula_cells_in(self, rng: str) -> list[str]:
        if ":" not in rng:
            return [rng] if rng in self.deps else []
        s, r1, c1, r2, c2 = _parse_range(rng)
        return [k for k in self.deps if k.rsplit("!", 1)[0] == s and
                (lambda t: r1 <= t[1] <= r2 and c1 <= t[2] <= c2)(_parse_range(k))]


def op_trace(data: bytes, name: str, p: dict) -> dict:
    a = analyse(data, name)
    a.require_workbook()
    g = _Graph(a)
    depth, direction, budget = max(1, min(p.get("depth", 3), 10)), p.get("direction", "both"), [400]

    def node(k: str) -> dict:
        cell = a.cell(k) if ":" not in k else None
        out: dict[str, Any] = {"cell": k}
        if cell and "f" in cell:
            out.update(formula="=" + cell["f"], value=_stored(cell), status=a.status(k))
        elif ":" not in k:
            out.update(value=_stored(cell), kind="input" if cell else "empty")
        return out

    def up(k: str, level: int, seen: set) -> dict:
        out = node(k)
        if level >= depth or budget[0] <= 0:
            return out
        children = []
        for r in g.deps.get(k, ()):
            budget[0] -= 1
            if r in seen:
                continue
            seen.add(r)
            child = up(r, level + 1, seen) if ":" not in r else node(r)
            if ":" in r and level + 1 < depth:
                child["formula_cells_inside"] = [up(f, level + 2, seen) for f in g.formula_cells_in(r)[:20]]
            children.append(child)
        if children:
            out["reads"] = children
        return out

    def down(k: str, level: int, seen: set) -> dict:
        out = node(k)
        if level >= depth or budget[0] <= 0:
            return out
        children = []
        for d in g.dependents(k):
            budget[0] -= 1
            if d not in seen:
                seen.add(d)
                children.append(down(d, level + 1, seen))
        if children:
            out["used_by"] = children
        return out

    def driving_inputs(k: str) -> list[str]:
        stack, seen, inputs = [k], {k}, set()
        while stack and len(seen) < 20000:
            for r in g.deps.get(stack.pop(), ()):
                if r in seen:
                    continue
                seen.add(r)
                inner = g.formula_cells_in(r)
                stack.extend(inner)
                if ":" in r:
                    _, r1, c1, r2, c2 = _parse_range(r)
                    if len(inner) < (r2 - r1 + 1) * (c2 - c1 + 1):  # holds inputs, not only formulas
                        inputs.add(r)
                elif r not in g.deps:
                    inputs.add(r)
        return sorted(inputs)[:200]

    results = []
    for k in [a.canonical(c) for c in p["cells"]][:10]:
        entry: dict[str, Any] = {"cell": k}
        if direction in ("precedents", "both"):
            entry["precedents"] = up(k, 0, {k})
            entry["driving_inputs"] = driving_inputs(k)
        if direction in ("dependents", "both"):
            entry["dependents"] = down(k, 0, {k})
        results.append(entry)
    return {"file": name, "depth": depth, "results": results, "truncated": budget[0] <= 0}


def _number(v, what: str) -> float:
    if isinstance(v, bool) or not isinstance(v, (int, float)):
        raise ToolError(f"{what} is {_plain(v)!r}, not a number")
    return float(v)


def _recalc_numeric(a: Analysis, overrides: dict, target: str) -> float:
    values, failed = a.namespace["compute"](overrides)
    if target in failed:
        raise ToolError(f"{target} cannot be computed: {failed[target]}")
    return _number(values.get(target), f"{target} (with {overrides})")


def _formula_target(a: Analysis, key: str) -> str:
    k = a.canonical(key)
    cell = a.cell(k)
    if not cell or "f" not in cell:
        raise ToolError(f"{k} is not a formula cell; pick the result cell")
    return k


def _input_cell(a: Analysis, key: str) -> str:
    k = a.canonical(key)
    cell = a.cell(k)
    if cell and "f" in cell:
        raise ToolError(f"{k} holds a formula; choose an input cell it depends on")
    return k


def _trust(a: Analysis, k: str) -> dict:
    status = a.status(k)
    out = {"status": status}
    if status in ("unverified", "mismatch", "disputed"):
        out["warning"] = a.reason(k) or "not verified: do not present this as Excel's answer"
    return out


def op_goal_seek(data: bytes, name: str, p: dict) -> dict:
    a = analyse(data, name)
    a.require_workbook()
    if not a.logic:
        raise ToolError("the workbook has no formulas")
    target, changing = _formula_target(a, p["target_cell"]), _input_cell(a, p["changing_cell"])
    goal, tol = float(p["target_value"]), float(p.get("tolerance") or 1e-7)
    max_iter = max(10, min(int(p.get("max_iterations") or 200), 500))
    start = a.cell(changing)
    x0 = float(_stored(start)) if isinstance(_stored(start), (int, float)) and not isinstance(_stored(start), bool) else 0.0
    calls = [0]

    def f(x: float) -> float:
        calls[0] += 1
        return _recalc_numeric(a, {changing: x}, target) - goal

    def scale(v):  # tolerance relative to the size of the goal
        return tol * max(1.0, abs(goal))

    lo, hi = p.get("lower"), p.get("upper")
    solution, fx = None, None
    if lo is None and hi is None:
        # Secant from the current value: fast for the smooth models most workbooks are.
        start_x = x0
        f0 = f(x0)
        try:
            x1 = x0 * 1.01 + (0.01 if x0 == 0 else 0)
            f1 = f(x1)
            for _ in range(60):
                if abs(f1) <= scale(f1):
                    solution, fx = x1, f1
                    break
                if f1 == f0:
                    break
                x0, f0, x1 = x1, f1, x1 - f1 * (x1 - x0) / (f1 - f0)
                f1 = f(x1)
        except (ToolError, OverflowError, ZeroDivisionError):
            pass  # the secant stepped somewhere the formula fails; fall back to bracketing
        if solution is None:
            x0, f0 = start_x, f(start_x)
            # Look outward from the start for a sign change, then bisect.
            step, base, fb = max(1.0, abs(x0)), x0, f0
            for _ in range(60):
                for cand in (base - step, base + step):
                    try:
                        fc = f(cand)
                    except ToolError:
                        continue
                    if (fc <= 0) != (fb <= 0):
                        lo, hi = sorted((base, cand))
                        break
                if lo is not None:
                    break
                step *= 2
    if solution is None:
        if lo is None or hi is None:
            raise ToolError(f"no value of {changing} found that makes {target} = {goal:g}; try giving lower and upper "
                            "bounds that bracket the answer")
        lo, hi = float(lo), float(hi)
        flo, fhi = f(lo), f(hi)
        if (flo <= 0) == (fhi <= 0):
            raise ToolError(f"{target} - {goal:g} has the same sign at {changing} = {lo:g} and {hi:g} "
                            f"({flo + goal:g} and {fhi + goal:g}); widen the bounds")
        for _ in range(max_iter):
            mid = (lo + hi) / 2
            fm = f(mid)
            if abs(fm) <= scale(fm) or hi - lo <= 1e-12 * max(1.0, abs(mid)):
                solution, fx = mid, fm
                break
            if (fm <= 0) == (flo <= 0):
                lo, flo = mid, fm
            else:
                hi = mid
        else:
            solution, fx = mid, fm
    return {
        "file": name,
        "changing_cell": changing,
        "solution": solution,
        "original_value": _stored(start),
        "target_cell": target,
        "target_value": goal,
        "achieved_value": fx + goal,
        "evaluations": calls[0],
        **_trust(a, target),
    }


def _value_list(p: dict, prefix: str) -> list:
    values = p.get(f"{prefix}values")
    if values:
        return [_coerce(v) for v in values][:100]
    start, stop, steps = p.get(f"{prefix}start"), p.get(f"{prefix}stop"), p.get(f"{prefix}steps") or 11
    if start is None or stop is None:
        raise ToolError(f"give {prefix}values, or {prefix}start and {prefix}stop")
    steps = max(2, min(int(steps), 100))
    return [round(float(start) + (float(stop) - float(start)) * i / (steps - 1), 12) for i in range(steps)]


def op_sensitivity(data: bytes, name: str, p: dict) -> dict:
    a = analyse(data, name)
    a.require_workbook()
    if not a.logic:
        raise ToolError("the workbook has no formulas")
    x = _input_cell(a, p["input_cell"])
    xs = _value_list(p, "")
    outputs = [_formula_target(a, c) for c in p["output_cells"]][:10]
    compute = a.namespace["compute"]
    base, _ = compute()
    if p.get("second_input_cell"):
        y = _input_cell(a, p["second_input_cell"])
        ys = _value_list(p, "second_")
        if len(xs) * len(ys) > 400:
            raise ToolError("a two-way table is limited to 400 cells; use fewer values")
        out = outputs[0]
        grid = []
        for xv in xs:
            row = []
            for yv in ys:
                values, failed = compute({x: xv, y: yv})
                row.append(_plain(values.get(out)) if out not in failed else f"error: {failed[out]}")
            grid.append(row)
        return {"file": name, "output_cell": out, "rows": {"cell": x, "values": xs},
                "columns": {"cell": y, "values": ys}, "table": grid, "base_value": _plain(base.get(out)),
                **_trust(a, out)}
    rows = []
    for xv in xs:
        values, failed = compute({x: xv})
        rows.append({x: xv, **{o: (_plain(values.get(o)) if o not in failed else f"error: {failed[o]}") for o in outputs}})
    return {"file": name, "input_cell": x, "original_input": _stored(a.cell(x)),
            "base_values": {o: _plain(base.get(o)) for o in outputs}, "rows": rows,
            "trust": {o: _trust(a, o) for o in outputs}}


def op_compare(data: bytes, name: str, p: dict) -> dict:
    other = p["_other"]
    a = analyse(data, name)
    b = analyse(base64.b64decode(other["b64"]), other["name"])
    a.require_workbook()
    b.require_workbook()
    from .formula import split_a1
    sheets_a = {s["name"]: s for s in a.model["sheets"]}
    sheets_b = {s["name"]: s for s in b.model["sheets"]}
    changes, counts = [], {"value_changed": 0, "formula_changed": 0, "added": 0, "removed": 0}
    limit = max(1, min(p.get("limit", 200), MAX_CELLS_RETURNED))
    for sname in [s for s in sheets_a if s in sheets_b]:
        ca, cb = sheets_a[sname]["cells"], sheets_b[sname]["cells"]
        for ref in sorted(set(ca) | set(cb), key=split_a1):
            x, y = ca.get(ref), cb.get(ref)
            vx, vy = _stored(x), _stored(y)
            fx, fy = (x or {}).get("f"), (y or {}).get("f")
            if x is None or (vx is None and fx is None):
                kind = "added" if (vy is not None or fy) else None
            elif y is None or (vy is None and fy is None):
                kind = "removed"
            elif fx != fy:
                kind = "formula_changed"
            elif vx != vy and not (isinstance(vx, float) and isinstance(vy, float)
                                   and abs(vx - vy) <= 1e-12 * max(1.0, abs(vx))):
                kind = "value_changed"
            else:
                kind = None
            if not kind:
                continue
            counts[kind] += 1
            if len(changes) < limit:
                entry = {"cell": f"{sname}!{ref}", "change": kind, "before": vx, "after": vy,
                         **_labels(sheets_b[sname] if y else sheets_a[sname], ref)}
                if fx or fy:
                    entry.update(formula_before=fx and "=" + fx, formula_after=fy and "=" + fy)
                changes.append(entry)
    return {
        "before": name, "after": other["name"],
        "sheets_added": [s for s in sheets_b if s not in sheets_a],
        "sheets_removed": [s for s in sheets_a if s not in sheets_b],
        "counts": counts,
        "changes": changes,
        "truncated": sum(counts.values()) > len(changes),
        "verdicts": {"before": a.certificate["verdict"]["level"], "after": b.certificate["verdict"]["level"]},
    }


OPS = {"inspect": op_inspect, "convert": op_convert, "recalculate": op_recalculate, "explain": op_explain,
       "find": op_find, "trace": op_trace, "goal_seek": op_goal_seek, "sensitivity": op_sensitivity,
       "compare": op_compare}


def _limit_memory():  # runs in the worker before exec
    try:
        import resource
        resource.setrlimit(resource.RLIMIT_AS, (WORKER_MEMORY, WORKER_MEMORY))
    except (ImportError, ValueError, OSError):
        pass


def run_op(op: str, data: bytes, name: str, params: dict, isolate: bool) -> dict:
    if not isolate:
        return OPS[op](data, name, params)
    request = json.dumps({"op": op, "name": name, "params": params, "data": base64.b64encode(data).decode()})
    # One math thread: per-core thread pools reserve address space and would trip the memory limit on big hosts.
    env = {**os.environ, "OPENBLAS_NUM_THREADS": "1", "OMP_NUM_THREADS": "1", "MKL_NUM_THREADS": "1",
           "MALLOC_ARENA_MAX": "2"}
    try:
        with WORKER_SLOTS:  # bounds total memory: at most N conversions at once, the rest queue
            proc = subprocess.run([sys.executable, "-m", "provenance.mcp_server", "--worker"], input=request.encode(),
                                  capture_output=True, timeout=WORKER_TIMEOUT, env=env,
                                  preexec_fn=_limit_memory if os.name == "posix" else None)
    except subprocess.TimeoutExpired:
        raise ToolError(f"the workbook took longer than {WORKER_TIMEOUT}s to process") from None
    try:
        reply = json.loads(proc.stdout)
    except json.JSONDecodeError:
        if proc.returncode < 0 or proc.returncode == 137:
            raise ToolError("the workbook used too much memory on the hosted server; run Provenance locally "
                            "for large files") from None
        last = proc.stderr.decode(errors="replace").strip().splitlines()[-1:] or [""]
        print(f"worker {op} exited {proc.returncode}: {last[0][:300]}", file=sys.stderr)  # no file contents
        raise ToolError("the workbook could not be processed (it may be too large or complex)") from None
    if "error" in reply:
        raise ToolError(reply["error"])
    return reply["result"]


def _worker() -> int:
    import logging
    import warnings
    logging.disable(logging.CRITICAL)
    warnings.simplefilter("ignore")
    import contextlib
    req = json.loads(sys.stdin.buffer.read())
    try:
        with contextlib.redirect_stdout(sys.stderr):  # stdout carries only the reply
            result = OPS[req["op"]](base64.b64decode(req["data"]), req["name"], req["params"])
        reply = {"result": result}
    except ToolError as exc:
        reply = {"error": str(exc)}
    except MemoryError:
        reply = {"error": f"this workbook needs more than the hosted server's {WORKER_MEMORY >> 20} MB of memory to "
                          "convert; run Provenance locally (pip install \"provenance-proxy[mcp] @ git+https://github.com/slimboi34/provenance\") for large files"}
    except Exception as exc:  # noqa: BLE001 - reported to the caller, never a traceback
        reply = {"error": f"processing failed: {type(exc).__name__}: {exc}"}
    sys.stdout.write(json.dumps(reply, ensure_ascii=False, default=repr))
    return 0


# -------------------------------------------------------------------- server

REQUIRED = inspect.Parameter.empty


def _p(annotation, description: str, default=None):
    """A tool parameter: (annotation with its description, default)."""
    from pydantic import Field
    return Annotated[annotation, Field(description=description)], default


CELL = "a cell key like 'Sheet1!B2'"

# name -> (op, title, read_only_locally, description, {param: _p(...)}, prepare(kwargs, remote) -> params)
TOOLS: dict[str, tuple] = {
    "inspect_workbook": (
        "inspect", "Inspect workbook", True,
        "Show a workbook's sheets, values and formulas as a Markdown table (Excel's stored results), with sheet "
        "sizes, defined names and whether it has macros. Start here for a new workbook.",
        {"sheet": _p(str | None, "only this sheet"),
         "max_rows": _p(int, "rows per sheet to show", 100),
         "max_cols": _p(int, "columns per sheet to show", 26)},
        None),
    "find_cells": (
        "find", "Find cells", True,
        "Search a workbook for text, numbers or formula fragments, e.g. 'revenue', '0.2' or 'VLOOKUP'. Each match "
        "comes with its row label and column header, so you can locate the cells to use in other tools without "
        "reading the whole sheet.",
        {"query": _p(str, "text, number or regex to look for", REQUIRED),
         "match": _p(Literal["contains", "exact", "regex"], "how to match text", "contains"),
         "search_in": _p(Literal["values", "formulas", "both"], "what to search", "both"),
         "sheet": _p(str | None, "only this sheet"),
         "limit": _p(int, "maximum matches to return", 50)},
        None),
    "explain_cells": (
        "explain", "Explain cells", True,
        "For each cell: its formula, the value Excel stored, the value Provenance recomputes, the cells it reads, "
        "and whether the formula is verified (and why not, if not).",
        {"cells": _p(list[str], f"cells to explain, each {CELL}", REQUIRED)},
        None),
    "trace_cells": (
        "trace", "Trace precedents and dependents", True,
        "Follow a cell's formula chain: precedents (what it is calculated from, several levels deep), the input "
        "cells that ultimately drive it, and dependents (everything that changes when it changes).",
        {"cells": _p(list[str], f"cells to trace (up to 10), each {CELL}", REQUIRED),
         "direction": _p(Literal["precedents", "dependents", "both"], "which way to trace", "both"),
         "depth": _p(int, "levels to follow (1-10)", 3)},
        None),
    "recalculate_workbook": (
        "recalculate", "Recalculate with new inputs", True,
        "What-if analysis: set input cells to new values and recalculate every formula with Excel's semantics, "
        'without Excel. Example overrides: {"Budget!B2": 40, "Rates!B1": 0.25}. Returns the requested cells (or '
        "every cell that changed) with old and new values; cells whose formulas are not verified carry a warning. "
        "The file is not modified.",
        {"overrides": _p(dict[str, Any], "input cell -> new value (number, text or true/false)", REQUIRED),
         "cells": _p(list[str] | None, "cells to report; default: every cell that changed")},
        None),
    "goal_seek": (
        "goal_seek", "Goal seek", True,
        "Find the value of one input cell that makes a formula cell reach a target, like Excel's Goal Seek: e.g. "
        "what unit price makes Summary!F12 (profit) equal 10000. Give lower/upper if you know a range that "
        "contains the answer.",
        {"target_cell": _p(str, f"the formula cell to drive, {CELL}", REQUIRED),
         "target_value": _p(float, "the value it should reach", REQUIRED),
         "changing_cell": _p(str, f"the input cell to change, {CELL}", REQUIRED),
         "lower": _p(float | None, "lowest value to consider for the changing cell"),
         "upper": _p(float | None, "highest value to consider for the changing cell"),
         "tolerance": _p(float, "relative accuracy", 1e-7)},
        None),
    "sensitivity_table": (
        "sensitivity", "Sensitivity table", True,
        "Vary one input over a list or range of values and tabulate how output cells respond (a data table). "
        "Add second_input_cell for a two-way table of the first output cell.",
        {"input_cell": _p(str, f"the input to vary, {CELL}", REQUIRED),
         "output_cells": _p(list[str], "formula cells to report (up to 10)", REQUIRED),
         "values": _p(list[Any] | None, "values to try; or use start/stop/steps"),
         "start": _p(float | None, "first value of an even range"),
         "stop": _p(float | None, "last value of an even range"),
         "steps": _p(int, "number of values in the range (2-100)", 11),
         "second_input_cell": _p(str | None, "a second input for a two-way table"),
         "second_values": _p(list[Any] | None, "values for the second input"),
         "second_start": _p(float | None, "first value for the second input"),
         "second_stop": _p(float | None, "last value for the second input"),
         "second_steps": _p(int, "number of values for the second input", 11)},
        None),
    "compare_workbooks": (
        "compare", "Compare two workbooks", True,
        "Compare two versions of a workbook: sheets added or removed, and every cell whose value or formula "
        "changed, with labels. The first workbook is 'before'; give the second ('after') with the other_ "
        "parameters.",
        {"other_file_id": _p(str | None, "'after' workbook: file_id from the upload page (hosted server)"),
         "other_url": _p(str | None, "'after' workbook: https link"),
         "other_content_base64": _p(str | None, "'after' workbook: file bytes, base64"),
         "other_path": _p(str | None, "'after' workbook: local path (local server only)"),
         "other_filename": _p(str | None, "'after' workbook's file name"),
         "limit": _p(int, "maximum changed cells to list", 200)},
        None),
    "convert_workbook": (
        "convert", "Convert and certify workbook", False,
        "Convert a workbook to canonical JSON plus a Python module of its formulas, and certify the conversion: "
        "structure is proven by rebuilding the file; each formula is replayed against Excel's stored result and "
        "fuzz-tested against an independent engine. Returns the verdict (verified / partial / failed), counts, "
        "and why any cells are unverified. include_certificate / include_logic return the full certificate and "
        "the generated Python.",
        {"fuzz_cases": _p(int, "random input cases to test against a second engine", 10),
         "seed": _p(int, "random seed for the fuzz cases", 0),
         "include_certificate": _p(bool, "return the full certificate JSON", False),
         "include_logic": _p(bool, "return the generated Python module", False),
         "output_dir": _p(str | None, "local server only: folder for the artifacts (default <file>.provenance/)")},
        None),
}


def _prepare(tool: str, kw: dict, src: dict, remote: bool) -> dict:
    if tool == "convert_workbook":
        kw["fuzz_cases"] = max(0, min(kw.get("fuzz_cases", 10), 25 if remote else 1000))
        if remote:
            kw.pop("output_dir", None)
        elif kw.get("output_dir") is None and src.get("path"):
            kw["output_dir"] = str(Path(src["path"]).expanduser()) + ".provenance"
    if tool == "compare_workbooks":
        other = {k: kw.pop(f"other_{k}", None) for k in (*SOURCE_KEYS, "filename")}
        if remote:
            other.pop("path", None)
        if not any(other.get(k) for k in SOURCE_KEYS):
            raise ToolError("give the second workbook as other_file_id, other_url, other_content_base64"
                            + ("" if remote else " or other_path"))
        data, name = load_source(other, remote)
        kw["_other"] = {"b64": base64.b64encode(data).decode(), "name": name}
    return kw


def build_server(remote: bool = False):
    import anyio
    from mcp.server.mcpserver import MCPServer
    from mcp.server.mcpserver.exceptions import ToolError as McpToolError
    from mcp.types import Icon, ToolAnnotations

    server = MCPServer(name="provenance", title="Provenance: verified Excel", version=__version__,
                       instructions=INSTRUCTIONS + (UPLOAD_HINT.format(url=PUBLIC_URL or "this server") if remote else ""),
                       website_url=REPO, icons=[Icon(src=ICON_URL, mime_type="image/png", sizes=["512x512"])])
    if remote:
        src_doc = ("Give the workbook as exactly one of: file_id (from the upload page" +
                   (f", {PUBLIC_URL}/upload" if PUBLIC_URL else "") + "), url (https; Google Sheets share links work) "
                   "or content_base64 (the file's bytes, small files only).")
        sources = {"file_id": _p(str | None, "id from the upload page, e.g. pv_..."),
                   "url": _p(str | None, "https link to the workbook"),
                   "content_base64": _p(str | None, "the file's bytes, base64 (small files)"),
                   "filename": _p(str | None, "the file's name, if sending content_base64")}
    else:
        src_doc = "Give the workbook as exactly one of: path (a local file), url (https) or content_base64."
        sources = {"path": _p(str | None, "path to a local .xlsx / .xlsm file"),
                   "url": _p(str | None, "https link to the workbook"),
                   "content_base64": _p(str | None, "the file's bytes, base64"),
                   "filename": _p(str | None, "the file's name, if sending content_base64")}

    def register(tool: str, spec: tuple) -> None:
        op, title, read_only_locally, description, own, _ = spec
        if remote:
            own = {k: v for k, v in own.items() if k not in ("output_dir", "other_path")}
        else:
            own = {k: v for k, v in own.items() if k != "other_file_id"}
        if tool == "convert_workbook" and not remote:
            own["fuzz_cases"] = _p(int, "random input cases to test against a second engine", 25)
        ordered = {**{k: v for k, v in own.items() if v[1] is REQUIRED}, **sources,
                   **{k: v for k, v in own.items() if v[1] is not REQUIRED}}
        params = [inspect.Parameter(k, inspect.Parameter.KEYWORD_ONLY, annotation=ann, default=default)
                  for k, (ann, default) in ordered.items()]

        async def handler(**kw):
            try:
                src = {k: kw.pop(k, None) for k in (*SOURCE_KEYS, "filename")}
                prepared = await anyio.to_thread.run_sync(_prepare, tool, kw, src, remote)
                data, name = await anyio.to_thread.run_sync(load_source, src, remote)
                return await anyio.to_thread.run_sync(run_op, op, data, name, prepared, remote)
            except ToolError as exc:
                raise McpToolError(str(exc)) from None

        handler.__signature__ = inspect.Signature(params, return_annotation=dict)
        handler.__annotations__ = {k: ann for k, (ann, _) in ordered.items()} | {"return": dict}
        handler.__name__ = tool
        read_only = remote or read_only_locally
        server.add_tool(handler, name=tool, title=title, description=f"{description} {src_doc}",
                        annotations=ToolAnnotations(read_only_hint=read_only, destructive_hint=False,
                                                    idempotent_hint=True, open_world_hint=True))

    for tool, spec in TOOLS.items():
        register(tool, spec)
    return server


REPO = "https://github.com/slimboi34/provenance"
CONTACT = os.environ.get("PROVENANCE_CONTACT", f"{REPO}/issues")
ICON_URL = "https://raw.githubusercontent.com/slimboi34/provenance/main/docs/icon.png"

PRIVACY = f"""\
<h1>Provenance MCP server: privacy policy</h1>
<p><em>Effective 2026-10-10. Applies to the hosted server at this address and to the local
<code>provenance-mcp</code> program.</em></p>
<h2>What is collected</h2>
<p>Only what a tool call contains: the workbook you send (as file contents or a URL to fetch) and the tool's
parameters, such as cell values for a recalculation. No account, name, email address or cookie is involved.</p>
<h2>How it is used and stored</h2>
<p>Each workbook is used only to answer the calls that use it. On the hosted server it is written to a temporary
folder, processed in a separate short-lived process, and deleted when the call finishes, normally within seconds.
A file you put on the upload page is kept for at most one hour so that Claude can use it by its file id, then deleted;
it is never shown or sent back to anyone, and the file id (a random 32-character code) is the only way to use it.
Nothing else from a file is kept: no copies, no database, no cache across calls. Request logs record time, status and
duration, never file contents, URLs you send or cell values. The local program runs entirely on your computer and
only writes files to the output folder you choose.</p>
<h2>Third parties</h2>
<p>Workbooks and results are not sold, shared or used to train models. If you give a URL, the server fetches that
URL from its owner (for example Google Sheets), which sees an ordinary download request. The hosted server runs on
Railway (railway.com), which provides the machine and network.</p>
<h2>Retention</h2>
<p>Workbook data: deleted at the end of each call; uploaded files, within one hour. Request logs without content: kept by the hosting provider for
up to 30 days.</p>
<h2>Contact</h2>
<p>Questions or deletion requests: <a href="{CONTACT}">{CONTACT}</a>.</p>
"""

HOME = f"""\
<img src="/icon.png" width=72 height=72 alt="">
<h1>Provenance: verified Excel for Claude</h1>
<p>An MCP server that lets Claude read Excel workbooks exactly, recalculate them with new inputs, and say which
answers are verified. A formula Provenance does not fully model is reported as unverified, never guessed.</p>
<h2>Connect</h2>
<p>In Claude: Settings &rarr; Connectors &rarr; Add custom connector, URL <code>{{url}}</code>. No sign-in.</p>
<p>Locally (reads files on your computer): <code>pip install "provenance-proxy[mcp] @ git+https://github.com/slimboi34/provenance"</code>, then
<code>claude mcp add provenance -- provenance-mcp</code>.</p>
<h2>Tools</h2>
<ul>
<li><b>inspect_workbook</b>: sheets, values and formulas as a table.</li>
<li><b>find_cells</b>: find text, numbers or formulas, with row and column labels.</li>
<li><b>explain_cells</b>: a cell's formula, stored and recomputed values, inputs, and whether it is verified.</li>
<li><b>trace_cells</b>: precedents, dependents, and the inputs that ultimately drive a result.</li>
<li><b>recalculate_workbook</b>: what-if: change input cells, get every affected result.</li>
<li><b>goal_seek</b>: find the input value that makes a result hit a target.</li>
<li><b>sensitivity_table</b>: one- or two-way data tables over input ranges.</li>
<li><b>compare_workbooks</b>: every changed value, formula and sheet between two versions.</li>
<li><b>convert_workbook</b>: canonical JSON + Python for every formula, with a certificate (verified / partial /
failed) of what was proven.</li>
</ul>
<h2>Giving Claude a workbook</h2>
<p><b><a href="/upload">Upload it here</a></b> and paste the file id into your chat, or share a public https link
(Google Sheets share links work). The hosted server handles typical workbooks (up to about 1 MB compressed); use
the local version for larger files.</p>
<h2>Examples</h2>
<ol>
<li>"Here's our pricing model: https://&hellip;/pricing.xlsx. What does the total come to if units go to 500?"</li>
<li>"In file pv_&hellip;, what price per unit gets profit to 10,000?"</li>
<li>"Show me how margin changes as the discount goes from 0% to 30%."</li>
<li>"What changed between last month's forecast and this one?"</li>
<li>"Explain how Summary!F12 is calculated, and is it verified?"</li>
</ol>
<p><a href="/privacy">Privacy policy</a> &middot; <a href="{REPO}">Source and documentation</a> &middot;
<a href="{CONTACT}">Support</a></p>
"""


UPLOAD_PAGE = """\
<img src="/icon.png" width=56 height=56 alt="">
<h1>Give Claude a workbook</h1>
<p>Drop an .xlsx or .xlsm file here. You get a file id to paste into your Claude chat; the file is deleted after one
hour and is never shown to anyone. <a href="/privacy">Privacy</a></p>
<label id=drop style="display:block;border:2px dashed #9cc7b2;border-radius:14px;padding:2.5rem;text-align:center;
cursor:pointer;background:#f6fbf8">Drop a workbook, or click to choose one
<input id=file type=file accept=".xlsx,.xlsm" hidden></label>
<div id=out style="margin-top:1.5rem"></div>
<script>
const drop = document.getElementById('drop'), input = document.getElementById('file'), out = document.getElementById('out');
const esc = s => s.replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
async function send(f) {
  out.textContent = 'Uploading ' + f.name + '...';
  const r = await fetch('/upload?name=' + encodeURIComponent(f.name), {method: 'POST', body: f});
  const j = await r.json();
  if (!r.ok) { out.innerHTML = '<p style="color:#b42318">' + esc(j.error || 'Upload failed') + '</p>'; return; }
  const line = 'Use Provenance on my workbook ' + j.name + ', file_id ' + j.file_id;
  out.innerHTML = '<p>Uploaded <b>' + esc(j.name) + '</b>. Paste this into Claude (valid for 1 hour):</p>' +
    '<p><code id=line style="display:block;padding:.8rem;word-break:break-all">' + esc(line) + '</code></p>' +
    '<button id=copy style="padding:.5rem 1rem;border-radius:8px;border:0;background:#0f5132;color:#fff;cursor:pointer">Copy</button>';
  document.getElementById('copy').onclick = () => { navigator.clipboard.writeText(line); document.getElementById('copy').textContent = 'Copied'; };
}
input.onchange = () => input.files[0] && send(input.files[0]);
drop.ondragover = e => { e.preventDefault(); drop.style.background = '#e3f3ea'; };
drop.ondragleave = () => drop.style.background = '#f6fbf8';
drop.ondrop = e => { e.preventDefault(); drop.style.background = '#f6fbf8'; e.dataTransfer.files[0] && send(e.dataTransfer.files[0]); };
</script>
"""


def _page(body: str) -> str:
    return ("<!doctype html><html lang=en><meta charset=utf-8><meta name=viewport content='width=device-width'>"
            "<title>Provenance MCP</title><link rel=icon type=image/png href=/icon.png><style>body{font:16px/1.55 system-ui,sans-serif;max-width:46rem;"
            "margin:3rem auto;padding:0 1.2rem;color:#1d1d1f}code{background:#f2f2f4;padding:.1em .3em;"
            "border-radius:4px}h2{margin-top:2rem}</style><body>" + body + "</body></html>")


def _add_pages(server) -> None:
    import anyio
    from starlette.responses import HTMLResponse, JSONResponse, Response

    @server.custom_route("/", methods=["GET"])
    async def home(request):
        url = str(request.base_url).rstrip("/").replace("http://", "https://", 1) + "/mcp"
        return HTMLResponse(_page(HOME.replace("{url}", url)))

    @server.custom_route("/privacy", methods=["GET"])
    async def privacy(request):
        return HTMLResponse(_page(PRIVACY))

    @server.custom_route("/icon.png", methods=["GET"])
    @server.custom_route("/favicon.ico", methods=["GET"])
    async def icon(request):
        return Response((Path(__file__).parent / "static" / "icon.png").read_bytes(), media_type="image/png",
                        headers={"Cache-Control": "public, max-age=86400"})

    @server.custom_route("/upload", methods=["GET"])
    async def upload_page(request):
        return HTMLResponse(_page(UPLOAD_PAGE))

    @server.custom_route("/upload", methods=["POST", "PUT"])
    async def upload(request):
        client = (request.headers.get("x-forwarded-for") or (request.client.host if request.client else "")).split(",")[0]
        now = time.time()
        recent = [t for t in _UPLOAD_TIMES.get(client, []) if t > now - 3600]
        if len(recent) >= UPLOADS_PER_HOUR:
            return JSONResponse({"error": "too many uploads from your address; try again later"}, status_code=429)
        _UPLOAD_TIMES[client] = recent + [now]
        if len(_UPLOAD_TIMES) > 10000:
            _UPLOAD_TIMES.clear()
        if int(request.headers.get("content-length") or 0) > MAX_BYTES:
            return JSONResponse({"error": f"file is larger than {MAX_BYTES >> 20} MB"}, status_code=413)
        chunks, size = [], 0
        async for chunk in request.stream():
            size += len(chunk)
            if size > MAX_BYTES:
                return JSONResponse({"error": f"file is larger than {MAX_BYTES >> 20} MB"}, status_code=413)
            chunks.append(chunk)
        name = request.query_params.get("name") or "workbook.xlsx"
        try:
            file_id = await anyio.to_thread.run_sync(store_upload, b"".join(chunks), name)
        except ToolError as exc:
            return JSONResponse({"error": str(exc)}, status_code=400)
        return JSONResponse({"file_id": file_id, "name": Path(name).name, "expires_in_seconds": UPLOAD_TTL})

    @server.custom_route("/health", methods=["GET"])
    async def health(request):
        return JSONResponse({"ok": True, "version": __version__})


def _config() -> str:
    exe = shutil.which("provenance-mcp") or str(Path(sys.executable).with_name("provenance-mcp"))
    return json.dumps({"mcpServers": {"provenance": {"command": exe}}}, indent=2) + \
        f"\n\nClaude Code:  claude mcp add provenance -- {exe}\n"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="provenance-mcp", description="Provenance MCP server")
    parser.add_argument("--http", action="store_true", help="serve streamable HTTP (remote connector mode)")
    parser.add_argument("--host", default=os.environ.get("HOST", "0.0.0.0"))
    parser.add_argument("--port", type=int, default=int(os.environ.get("PORT", "8000")))
    parser.add_argument("--config", action="store_true", help="print MCP client configuration and exit")
    parser.add_argument("--worker", action="store_true", help=argparse.SUPPRESS)
    args = parser.parse_args(argv)
    if args.worker:
        return _worker()
    if args.config:
        print(_config())
        return 0
    server = build_server(remote=args.http)
    if not args.http:
        server.run(transport="stdio")
        return 0
    import uvicorn
    from mcp.server.transport_security import TransportSecuritySettings
    _add_pages(server)
    app = server.streamable_http_app(
        stateless_http=True, json_response=True, max_request_body_size=MAX_BYTES * 4 // 3 + 65536,
        transport_security=TransportSecuritySettings(enable_dns_rebinding_protection=False), host=args.host)
    uvicorn.run(_friendly_get(app), host=args.host, port=args.port, log_level="info", proxy_headers=True,
                forwarded_allow_ips="*")
    return 0


MCP_PAGE = """\
<img src="/icon.png" width=56 height=56 alt="">
<h1>This is an MCP connector address</h1>
<p>It is working. It is meant to be added to Claude, not opened in a browser.</p>
<p>In Claude: <b>Settings &rarr; Connectors &rarr; Add custom connector</b>, and paste
<code>{url}</code>. No sign-in is needed.</p>
<p><a href="/">About Provenance</a> &middot; <a href="/upload">Upload a workbook</a> &middot;
<a href="/privacy">Privacy</a></p>
"""


def _friendly_get(app):
    """GET /mcp: a page for browsers, an immediate 405 for anything else.

    The server is stateless and sends no server-initiated messages, so it offers no
    SSE stream on GET (the spec allows 405); without this a plain GET held the
    connection open forever and a browser saw a bare 406.
    """
    async def wrapped(scope, receive, send):
        if scope["type"] == "http" and scope["method"] == "GET" and scope["path"].rstrip("/") == "/mcp":
            from starlette.requests import Request
            from starlette.responses import HTMLResponse, PlainTextResponse
            request = Request(scope)
            if "text/html" in request.headers.get("accept", ""):
                url = str(request.url.replace(query="")).replace("http://", "https://", 1)
                response = HTMLResponse(_page(MCP_PAGE.replace("{url}", url)))
            else:
                response = PlainTextResponse("This MCP server does not offer a GET stream; send JSON-RPC with POST.\n",
                                             status_code=405, headers={"Allow": "POST"})
            await response(scope, receive, send)
            return
        await app(scope, receive, send)
    return wrapped


if __name__ == "__main__":
    sys.exit(main())
