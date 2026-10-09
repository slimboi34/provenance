"""MCP server: tools over the real protocol, stdio and HTTP, plus the hosted-mode safety rules."""

import base64
import json
import socket
import subprocess
import sys
import time

import anyio
import pytest

pytest.importorskip("mcp")

from mcp.client import Client  # noqa: E402
from mcp.client.stdio import StdioServerParameters  # noqa: E402

from provenance import mcp_server as m  # noqa: E402

TOOLS = {"inspect_workbook", "convert_workbook", "recalculate_workbook", "explain_cells"}


def _payload(result):
    assert not result.is_error, result.content
    return result.structured_content or json.loads(result.content[0].text)


async def _exercise(client, source: dict, remote: bool):
    tools = {t.name: t for t in (await client.list_tools()).tools}
    assert set(tools) == TOOLS
    for tool in tools.values():
        assert tool.title and tool.annotations is not None
        assert tool.annotations.read_only_hint is not None or tool.annotations.destructive_hint is not None
    assert ("path" in tools["inspect_workbook"].input_schema["properties"]) is not remote

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


def test_http_server(budget):
    port = _free_port()
    proc = subprocess.Popen([sys.executable, "-m", "provenance.mcp_server", "--http", "--host", "127.0.0.1",
                             "--port", str(port)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
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
            m.load_source(kwargs.get("path"), kwargs.get("url"), kwargs.get("content_base64"), None, remote=True)


def test_google_sheets_links_export_as_xlsx():
    assert m._google_export("https://docs.google.com/spreadsheets/d/abc_D-1/edit#gid=0") == \
        "https://docs.google.com/spreadsheets/d/abc_D-1/export?format=xlsx"


def test_worker_isolation_reports_errors_cleanly():
    with pytest.raises(m.ToolError, match="cannot read this file"):
        m.run_op("inspect", b"not a zip", "x.xlsx", {}, isolate=True)
