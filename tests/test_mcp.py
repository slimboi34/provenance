"""MCP server: tools over the real protocol, stdio and HTTP, plus the hosted-mode safety rules."""

import base64
import json
import os
import socket
import subprocess
import sys
import time
import urllib.error
import urllib.request

import anyio
import pytest

pytest.importorskip("mcp")

from mcp.client import Client  # noqa: E402
from mcp.client.stdio import StdioServerParameters  # noqa: E402

from provenance import mcp_server as m  # noqa: E402

TOOLS = {"inspect_workbook", "convert_workbook", "recalculate_workbook", "explain_cells", "find_cells",
         "trace_cells", "goal_seek", "sensitivity_table", "compare_workbooks"}


def _payload(result):
    assert not result.is_error, result.content
    return result.structured_content or json.loads(result.content[0].text)


async def _exercise(client, source: dict, remote: bool):
    tools = {t.name: t for t in (await client.list_tools()).tools}
    assert set(tools) == TOOLS
    for tool in tools.values():
        assert tool.title and tool.annotations is not None
        assert tool.annotations.read_only_hint is not None or tool.annotations.destructive_hint is not None
    for tool in tools.values():
        props = tool.input_schema["properties"]
        assert ("path" in props) is not remote and ("file_id" in props) is remote
        assert all(p.get("description") for p in props.values()), tool.name

    inspected = _payload(await client.call_tool("inspect_workbook", source))
    assert [s["name"] for s in inspected["sheets"]] == ["Budget", "Rates"]
    assert "| **6** | Sum |" in inspected["markdown"]

    recalc = _payload(await client.call_tool("recalculate_workbook", {**source, "overrides": {"Budget!B2": 40},
                                                                       "cells": ["Budget!D6", "Budget!D12"]}))
    assert [(r["cell"], r["value"], r["before"]) for r in recalc["results"]] == [("Budget!D6", 112.0, 22.0),
                                                                                  ("Budget!D12", 22.4, 4.4)]

    explained = _payload(await client.call_tool("explain_cells", {**source, "cells": ["Budget!D12"]}))
    assert explained["cells"][0]["reads"] == ["Budget!D6", "Rates!B1"]

    converted = _payload(await client.call_tool("convert_workbook", {**source, "fuzz_cases": 0}))
    assert converted["verdict"] == "verified" and converted["structure"] == "proven"

    found = _payload(await client.call_tool("find_cells", {**source, "query": "sum"}))
    first = found["results"][0]
    assert (first["cell"], first["value"]) == ("Budget!A6", "Sum")

    traced = _payload(await client.call_tool("trace_cells", {**source, "cells": ["Budget!D12"], "depth": 5}))
    assert traced["results"][0]["driving_inputs"] == ["Budget!B2", "Budget!B3", "Budget!B4", "Budget!C2", "Budget!C3",
                                                      "Budget!C4", "Rates!B1"]  # not D2:D4: formulas only

    sought = _payload(await client.call_tool("goal_seek", {**source, "target_cell": "Budget!D6", "target_value": 100,
                                                            "changing_cell": "Budget!B2"}))
    assert sought["solution"] == pytest.approx(35.2) and sought["achieved_value"] == pytest.approx(100)

    table = _payload(await client.call_tool("sensitivity_table", {**source, "input_cell": "Rates!B1", "start": 0,
                                                                   "stop": 0.3, "steps": 4,
                                                                   "output_cells": ["Budget!D12"]}))
    assert [r["Budget!D12"] for r in table["rows"]] == [0.0, 2.2, 4.4, 6.6]

    other = "other_path" if "path" in source else "other_content_base64"
    same = source.get("path") or source.get("content_base64")
    compared = _payload(await client.call_tool("compare_workbooks", {**source, other: same}))
    assert sum(compared["counts"].values()) == 0

    bad = await client.call_tool("recalculate_workbook", {**source, "overrides": {"Budget!D6": 1}})
    assert bad.is_error and "holds a formula" in bad.content[0].text
    return converted


def test_stdio_server(tmp_path, budget):
    book = tmp_path / "budget.xlsx"
    book.write_bytes(budget)

    async def run():
        params = StdioServerParameters(command=sys.executable, args=["-m", "provenance.mcp_server"])
        async with Client(params) as client:
            converted = await _exercise(client, {"path": str(book)}, remote=False)
            assert any(p.endswith("budget.certificate.json") for p in converted["written"])
            assert (tmp_path / "budget.xlsx.provenance" / "budget_logic.py").exists()

    anyio.run(run)


def _free_port():
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def test_http_server(budget, tmp_path):
    port = _free_port()
    env = {**os.environ, "PROVENANCE_UPLOAD_DIR": str(tmp_path / "uploads")}
    proc = subprocess.Popen([sys.executable, "-m", "provenance.mcp_server", "--http", "--host", "127.0.0.1",
                             "--port", str(port)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=env)
    try:
        for _ in range(100):
            try:
                socket.create_connection(("127.0.0.1", port), timeout=0.2).close()
                break
            except OSError:
                time.sleep(0.1)

        async def run():
            async with Client(f"http://127.0.0.1:{port}/mcp") as client:
                await _exercise(client, {"content_base64": base64.b64encode(budget).decode(),
                                         "filename": "budget.xlsx"}, remote=True)
                # GET /mcp: a page for browsers, an immediate 405 for other clients (never a hanging stream).
                page = urllib.request.urlopen(urllib.request.Request(f"http://127.0.0.1:{port}/mcp",
                                                                     headers={"Accept": "text/html"}), timeout=5)
                assert b"MCP connector address" in page.read()
                with pytest.raises(urllib.error.HTTPError) as err:
                    urllib.request.urlopen(f"http://127.0.0.1:{port}/mcp", timeout=5)
                assert err.value.code == 405
                # The upload page flow: POST the file, then use its file_id.
                request = urllib.request.Request(f"http://127.0.0.1:{port}/upload?name=budget.xlsx", data=budget,
                                                 method="POST")
                file_id = json.loads(urllib.request.urlopen(request).read())["file_id"]
                inspected = _payload(await client.call_tool("inspect_workbook", {"file_id": file_id}))
                assert inspected["file"] == "budget.xlsx"
                gone = await client.call_tool("inspect_workbook", {"file_id": "pv_" + "x" * 32})
                assert gone.is_error and "expired" in gone.content[0].text

        anyio.run(run)
    finally:
        proc.terminate()
        proc.wait(10)


def test_hosted_mode_refuses_local_files_and_private_urls():
    for kwargs, message in [
        ({"path": "/etc/passwd"}, "not a path"),
        ({"url": "http://example.com/a.xlsx"}, "https://"),
        ({"url": "https://localhost/a.xlsx"}, "non-public"),
        ({"url": "https://127.0.0.1/a.xlsx"}, "non-public"),
        ({"content_base64": "not base64!"}, "not valid base64"),
        ({}, "exactly one"),
    ]:
        with pytest.raises(m.ToolError, match=message):
            m.load_source(kwargs, remote=True)


def test_upload_store(tmp_path, monkeypatch, budget):
    monkeypatch.setattr(m, "UPLOADS", tmp_path)
    file_id = m.store_upload(budget, "../../etc/q3.xlsx")
    assert m.FILE_ID.match(file_id)
    assert m.load_upload(file_id) == (budget, "q3.xlsx")
    with pytest.raises(m.ToolError, match="not an Excel workbook"):
        m.store_upload(b"hello", "x.xlsx")
    for bad in ("../" + file_id, "pv_short", file_id + "/x"):
        with pytest.raises(m.ToolError, match="unknown or has expired"):
            m.load_upload(bad)
    old = time.time() - m.UPLOAD_TTL - 5
    os.utime(tmp_path / file_id, (old, old))
    with pytest.raises(m.ToolError, match="expired"):
        m.load_upload(file_id)
    m._sweep_uploads()
    assert not (tmp_path / file_id).exists()


def test_goal_seek_brackets_and_reports_no_solution(budget):
    with pytest.raises(m.ToolError, match="no value"):  # D13 is IFERROR(D6/B4,-1) and B4 = 0: C3 can't move it
        m.run_op("goal_seek", budget, "b.xlsx", {"target_cell": "Budget!D13", "target_value": -5,
                                                  "changing_cell": "Budget!C3", "lower": None, "upper": None}, False)
    r = m.run_op("goal_seek", budget, "b.xlsx", {"target_cell": "Budget!D9", "target_value": 2,
                                                  "changing_cell": "Budget!B4", "lower": 0, "upper": 10}, False)
    assert (22 + 99.99 * r["solution"]) / (14 + r["solution"]) == pytest.approx(2)


def test_google_sheets_links_export_as_xlsx():
    assert m._google_export("https://docs.google.com/spreadsheets/d/abc_D-1/edit#gid=0") == \
        "https://docs.google.com/spreadsheets/d/abc_D-1/export?format=xlsx"


def test_worker_isolation_reports_errors_cleanly():
    with pytest.raises(m.ToolError, match="cannot read this file"):
        m.run_op("inspect", b"not a zip", "x.xlsx", {}, isolate=True)
