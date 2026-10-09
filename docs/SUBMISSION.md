# Connectors Directory submission: Provenance

Portal: https://claude.ai/directory/manage → Submit new → MCP connector.
Requirements: https://claude.com/docs/connectors/building/submission
Escalations: mcp-review@anthropic.com

## Checklist

- [x] Server deployed and reachable at `https://mcp-production-5fd2.up.railway.app/mcp` (Railway, `railway up` from this folder)
- [x] `https://mcp-production-5fd2.up.railway.app/` (docs) and `https://mcp-production-5fd2.up.railway.app/privacy` load
- [ ] Added in Claude as a custom connector (Settings → Connectors → Add custom connector) and every tool called once from a chat
- [x] Public repo `slimboi34/provenance` (the docs, privacy and support links point there)
- [x] Icon: 512×512 PNG at `docs/icon.png`
- [x] Every tool has a `title` and `readOnlyHint` (all four are read-only in hosted mode)
- [x] No authentication needed: the server only processes files the user sends; no account data
- [x] Privacy policy: README section + `/privacy`

## Listing (paste into the portal)

**Name** (≤100): Provenance: Verified Excel

**One-liner** (≤200): Read Excel workbooks exactly, recalculate them with new inputs, and know which answers are verified against Excel's own results.

**Description** (≤2,000):

> Provenance gives Claude exact access to Excel workbooks (.xlsx, .xlsm). Instead of eyeballing a spreadsheet, Claude reads every cell and formula, recalculates the workbook in Python with Excel's semantics, and reports which results are verified.
>
> - Inspect: sheets, values and formulas as a clean table.
> - Recalculate: "what if units were 500?" Change any input cells and get every affected result, before and after, without Excel.
> - Explain: how a cell is calculated, what it reads, and whether its formula was verified.
> - Convert and certify: lossless canonical JSON plus a Python module of every formula, with a certificate. Structure is proven by rebuilding the file, each formula is replayed against the result Excel stored, and the formulas are fuzz-tested against an independent engine.
>
> Formulas Provenance does not fully model are labelled unverified with the reason, never guessed, so Claude can tell you which numbers to trust.
>
> Send a workbook as a public https link (Google Sheets share links work) or as file contents. The hosted server handles typical workbooks (up to about 1 MB compressed); larger files run with the free local version. Files are processed in an isolated process and deleted when the call ends. Nothing is stored.

**Categories**: Data & Analytics; Productivity; Finance

**Documentation URL**: https://mcp-production-5fd2.up.railway.app/ (or the GitHub README)
**Privacy policy URL**: https://mcp-production-5fd2.up.railway.app/privacy
**Support contact**: https://github.com/slimboi34/provenance/issues

## Use cases

- Answer questions about a spreadsheet using exact cell values instead of a lossy read
- What-if analysis on financial, pricing and budget models
- Audit a workbook: which formulas reproduce Excel's results and which can't be verified
- Convert a workbook into JSON and Python for use in code

Prerequisites: none (no account). Reads data only; writes nothing.

## Authentication

None. The server holds no user accounts and acts on nothing but the file in each call.

## Data handling

Own software; no third-party API. Only fetches a URL the user supplies. No health data, no sponsored content.

## Test & launch (reviewer instructions)

No credentials needed. Add `https://mcp-production-5fd2.up.railway.app/mcp` as a custom connector, then:

1. "Inspect this workbook: https://raw.githubusercontent.com/slimboi34/provenance/main/docs/sample-budget.xlsx"
2. "In that workbook, set Budget!B2 to 40 and tell me the new Budget!D6"
3. "Explain how Budget!D12 is calculated and whether it's verified"
4. "Convert it and give me the verdict"

Sample workbook URL for reviewers: https://raw.githubusercontent.com/slimboi34/provenance/main/docs/sample-budget.xlsx
