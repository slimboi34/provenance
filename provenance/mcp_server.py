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
import binascii
import ipaddress
import json
import os
import re
import shutil
import socket
import subprocess
import sys
import tempfile
import urllib.error
import urllib.parse
import urllib.request
from collections import OrderedDict
from pathlib import Path
from typing import Any

from . import __version__
from .canon import ConversionError, sha256

MAX_BYTES = 25 * 1024 * 1024
FETCH_TIMEOUT = 30
WORKER_TIMEOUT = int(os.environ.get("PROVENANCE_MCP_TIMEOUT", "120"))
WORKER_MEMORY = int(os.environ.get("PROVENANCE_MCP_MEMORY_MB", "2048")) << 20
MAX_CELLS_RETURNED = 500
SUFFIXES = (".xlsx", ".xlsm", ".docx")
CELL_KEY = re.compile(r"^(?:'(?:[^']|'')+'|[^!'\s][^!]*)!\$?[A-Za-z]{1,3}\$?[0-9]{1,7}$")

INSTRUCTIONS = """\
Provenance reads Excel workbooks (.xlsx/.xlsm) exactly and recalculates them in Python.
Use inspect_workbook to see sheets, values and formulas; convert_workbook for the full
verified conversion and its certificate; recalculate_workbook for what-if questions
("what is the total if B2 is 40?"); explain_cells to trace a cell's formula and whether
it was verified. Cell keys look like "Sheet1!B2". Values of cells marked unverified must
not be presented as Excel's answer."""


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


def load_source(path: str | None, url: str | None, content_base64: str | None, filename: str | None,
                remote: bool) -> tuple[bytes, str]:
    """Return (bytes, file name) for the one source the caller gave."""
    given = [s for s in (path, url, content_base64) if s]
    if len(given) != 1:
        options = "url or content_base64" if remote else "path, url or content_base64"
        raise ToolError(f"give exactly one of {options}")
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


OPS = {"inspect": op_inspect, "convert": op_convert, "recalculate": op_recalculate, "explain": op_explain}


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
    try:
        proc = subprocess.run([sys.executable, "-m", "provenance.mcp_server", "--worker"], input=request.encode(),
                              capture_output=True, timeout=WORKER_TIMEOUT,
                              preexec_fn=_limit_memory if os.name == "posix" else None)
    except subprocess.TimeoutExpired:
        raise ToolError(f"the workbook took longer than {WORKER_TIMEOUT}s to process") from None
    try:
        reply = json.loads(proc.stdout)
    except json.JSONDecodeError:
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
    except Exception as exc:  # noqa: BLE001 - reported to the caller, never a traceback
        reply = {"error": f"processing failed: {type(exc).__name__}: {exc}"}
    sys.stdout.write(json.dumps(reply, ensure_ascii=False, default=repr))
    return 0


# -------------------------------------------------------------------- server

def build_server(remote: bool = False):
    import anyio
    from mcp.server.mcpserver import MCPServer
    from mcp.server.mcpserver.exceptions import ToolError as McpToolError
    from mcp.types import ToolAnnotations

    server = MCPServer(name="provenance", title="Provenance: verified Excel", version=__version__,
                       instructions=INSTRUCTIONS, website_url="https://github.com/slimboi34/provenance")
    src_doc = ("The workbook is given as exactly one of: url (https; public link, Google Sheets share links work), "
               "content_base64 (the file's bytes)" + ("" if remote else ", or path (a local file path)") + ".")

    async def call(op, path, url, content_base64, filename, params):
        try:
            data, name = await anyio.to_thread.run_sync(load_source, path, url, content_base64, filename, remote)
            return await anyio.to_thread.run_sync(run_op, op, data, name, params, remote)
        except ToolError as exc:
            raise McpToolError(str(exc)) from None

    if remote:
        async def inspect_workbook(url: str | None = None, content_base64: str | None = None, filename: str | None = None,
                                   sheet: str | None = None, max_rows: int = 100, max_cols: int = 26) -> dict:
            return await call("inspect", None, url, content_base64, filename,
                              {"sheet": sheet, "max_rows": max_rows, "max_cols": max_cols})

        async def convert_workbook(url: str | None = None, content_base64: str | None = None, filename: str | None = None,
                                   fuzz_cases: int = 10, seed: int = 0, include_certificate: bool = False,
                                   include_logic: bool = False) -> dict:
            return await call("convert", None, url, content_base64, filename,
                              {"fuzz_cases": max(0, min(fuzz_cases, 25)), "seed": seed,
                               "include_certificate": include_certificate, "include_logic": include_logic})

        async def recalculate_workbook(overrides: dict[str, Any], url: str | None = None,
                                       content_base64: str | None = None, filename: str | None = None,
                                       cells: list[str] | None = None) -> dict:
            return await call("recalculate", None, url, content_base64, filename, {"overrides": overrides, "cells": cells})

        async def explain_cells(cells: list[str], url: str | None = None, content_base64: str | None = None,
                                filename: str | None = None) -> dict:
            return await call("explain", None, url, content_base64, filename, {"cells": cells})
    else:
        async def inspect_workbook(path: str | None = None, url: str | None = None, content_base64: str | None = None,
                                   filename: str | None = None, sheet: str | None = None, max_rows: int = 100,
                                   max_cols: int = 26) -> dict:
            return await call("inspect", path, url, content_base64, filename,
                              {"sheet": sheet, "max_rows": max_rows, "max_cols": max_cols})

        async def convert_workbook(path: str | None = None, url: str | None = None, content_base64: str | None = None,
                                   filename: str | None = None, output_dir: str | None = None, fuzz_cases: int = 25,
                                   seed: int = 0, include_certificate: bool = False, include_logic: bool = False) -> dict:
            if output_dir is None and path:
                output_dir = str(Path(path).expanduser()) + ".provenance"
            return await call("convert", path, url, content_base64, filename,
                              {"fuzz_cases": max(0, fuzz_cases), "seed": seed, "output_dir": output_dir,
                               "include_certificate": include_certificate, "include_logic": include_logic})

        async def recalculate_workbook(overrides: dict[str, Any], path: str | None = None, url: str | None = None,
                                       content_base64: str | None = None, filename: str | None = None,
                                       cells: list[str] | None = None) -> dict:
            return await call("recalculate", path, url, content_base64, filename, {"overrides": overrides, "cells": cells})

        async def explain_cells(cells: list[str], path: str | None = None, url: str | None = None,
                                content_base64: str | None = None, filename: str | None = None) -> dict:
            return await call("explain", path, url, content_base64, filename, {"cells": cells})

    read_only = ToolAnnotations(read_only_hint=True, destructive_hint=False, idempotent_hint=True, open_world_hint=True)
    server.add_tool(inspect_workbook, name="inspect_workbook", title="Inspect workbook", annotations=read_only,
                    description="Show a workbook's sheets, values and formulas as a Markdown table (Excel's stored "
                                "results), with sheet sizes, defined names and whether it has macros. Optional: sheet "
                                "to show one sheet; max_rows/max_cols to bound the table. " + src_doc)
    server.add_tool(convert_workbook, name="convert_workbook", title="Convert and certify workbook",
                    annotations=read_only if remote else ToolAnnotations(
                        read_only_hint=False, destructive_hint=False, idempotent_hint=True, open_world_hint=True),
                    description="Convert a workbook to canonical JSON plus a Python module of its formulas, and "
                                "certify the conversion: structure is proven by rebuilding the file; each formula is "
                                "replayed against Excel's stored result and fuzz-tested against an independent engine. "
                                "Returns the verdict (verified / partial / failed), counts, and why any cells are "
                                "unverified. include_certificate / include_logic return the full certificate and the "
                                "generated Python." + ("" if remote else " With a path, artifacts are written to "
                                "output_dir (default: <file>.provenance/ beside it).") + " " + src_doc)
    server.add_tool(recalculate_workbook, name="recalculate_workbook", title="Recalculate with new inputs",
                    annotations=read_only,
                    description="What-if analysis: set input cells to new values and recalculate every formula with "
                                "Excel's semantics, without Excel. overrides maps cell keys to values, e.g. "
                                '{"Budget!B2": 40, "Rates!B1": 0.25}. Returns the requested cells (or every cell that '
                                "changed) with old and new values; cells whose formulas are not verified carry a "
                                "warning. The original file is not modified. " + src_doc)
    server.add_tool(explain_cells, name="explain_cells", title="Explain cells", annotations=read_only,
                    description="For each cell key (e.g. 'Budget!D6'): its formula, the value Excel stored, the value "
                                "Provenance recomputes, the cells it reads, and whether the formula is verified (and "
                                "why not, if not). " + src_doc)
    return server


REPO = "https://github.com/slimboi34/provenance"
CONTACT = os.environ.get("PROVENANCE_CONTACT", f"{REPO}/issues")

PRIVACY = f"""\
<h1>Provenance MCP server: privacy policy</h1>
<p><em>Effective 2026-10-09. Applies to the hosted server at this address and to the local
<code>provenance-mcp</code> program.</em></p>
<h2>What is collected</h2>
<p>Only what a tool call contains: the workbook you send (as file contents or a URL to fetch) and the tool's
parameters, such as cell values for a recalculation. No account, name, email address or cookie is involved.</p>
<h2>How it is used and stored</h2>
<p>Each workbook is used only to answer the call that sent it. On the hosted server it is written to a temporary
folder, processed in a separate short-lived process, and deleted when the call finishes, normally within seconds.
Nothing from the file is kept: no copies, no database, no cache across calls. Request logs record time, status and
duration, never file contents, URLs you send or cell values. The local program runs entirely on your computer and
only writes files to the output folder you choose.</p>
<h2>Third parties</h2>
<p>Workbooks and results are not sold, shared or used to train models. If you give a URL, the server fetches that
URL from its owner (for example Google Sheets), which sees an ordinary download request. The hosted server runs on
Railway (railway.com), which provides the machine and network.</p>
<h2>Retention</h2>
<p>Workbook data: deleted at the end of each call. Request logs without content: kept by the hosting provider for
up to 30 days.</p>
<h2>Contact</h2>
<p>Questions or deletion requests: <a href="{CONTACT}">{CONTACT}</a>.</p>
"""

HOME = f"""\
<h1>Provenance: verified Excel for Claude</h1>
<p>An MCP server that lets Claude read Excel workbooks exactly, recalculate them with new inputs, and say which
answers are verified. A formula Provenance does not fully model is reported as unverified, never guessed.</p>
<h2>Connect</h2>
<p>In Claude: Settings &rarr; Connectors &rarr; Add custom connector, URL <code>{{url}}</code>. No sign-in.</p>
<p>Locally (reads files on your computer): <code>pip install "provenance-proxy[mcp]"</code>, then
<code>claude mcp add provenance -- provenance-mcp</code>.</p>
<h2>Tools</h2>
<ul>
<li><b>inspect_workbook</b>: sheets, values and formulas as a table.</li>
<li><b>convert_workbook</b>: canonical JSON + Python for every formula, with a certificate (verified / partial /
failed) of what was proven.</li>
<li><b>recalculate_workbook</b>: what-if: change input cells, get every affected result.</li>
<li><b>explain_cells</b>: a cell's formula, stored and recomputed values, the cells it reads, and its status.</li>
</ul>
<p>Send the workbook as a public https link (Google Sheets share links work) or as file contents. Files up to
{MAX_BYTES >> 20} MB.</p>
<h2>Examples</h2>
<ol>
<li>"Here's our pricing model: https://&hellip;/pricing.xlsx. What does the total come to if units go to 500?"</li>
<li>"Explain how cell Summary!F12 in this workbook is calculated, and is it verified?"</li>
<li>"Convert this spreadsheet and tell me whether every formula reproduces Excel's results."</li>
</ol>
<p><a href="/privacy">Privacy policy</a> &middot; <a href="{REPO}">Source and documentation</a> &middot;
<a href="{CONTACT}">Support</a></p>
"""


def _page(body: str) -> str:
    return ("<!doctype html><html lang=en><meta charset=utf-8><meta name=viewport content='width=device-width'>"
            "<title>Provenance MCP</title><style>body{font:16px/1.55 system-ui,sans-serif;max-width:46rem;"
            "margin:3rem auto;padding:0 1.2rem;color:#1d1d1f}code{background:#f2f2f4;padding:.1em .3em;"
            "border-radius:4px}h2{margin-top:2rem}</style><body>" + body + "</body></html>")


def _add_pages(server) -> None:
    from starlette.responses import HTMLResponse, JSONResponse

    @server.custom_route("/", methods=["GET"])
    async def home(request):
        url = str(request.base_url).rstrip("/").replace("http://", "https://", 1) + "/mcp"
        return HTMLResponse(_page(HOME.replace("{url}", url)))

    @server.custom_route("/privacy", methods=["GET"])
    async def privacy(request):
        return HTMLResponse(_page(PRIVACY))

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
    import anyio
    from mcp.server.transport_security import TransportSecuritySettings
    _add_pages(server)
    anyio.run(lambda: server.run_streamable_http_async(
        host=args.host, port=args.port, stateless_http=True, json_response=True,
        max_request_body_size=MAX_BYTES * 4 // 3 + 65536,
        transport_security=TransportSecuritySettings(enable_dns_rebinding_protection=False)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
