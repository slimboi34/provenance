# Connectors Directory submission: Provenance

Portal: https://claude.ai/directory/manage → Submit new → MCP connector.
Requirements: https://claude.com/docs/connectors/building/submission
Escalations: mcp-review@anthropic.com

## Checklist

- [ ] Server deployed and reachable at `https://<domain>/mcp` (Railway, `railway up` from this folder)
- [ ] `https://<domain>/` (docs) and `https://<domain>/privacy` load
- [ ] Added in Claude as a custom connector (Settings → Connectors → Add custom connector) and every tool called once from a chat
- [ ] Public repo `slimboi34/provenance` (the docs, privacy and support links point there)
- [ ] Icon: 512×512 PNG
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
> Send a workbook as a public https link (Google Sheets share links work) or as file contents, up to 25 MB. Files are processed in an isolated process and deleted when the call ends. Nothing is stored.

**Categories**: Data & Analytics; Productivity; Finance

**Documentation URL**: https://<domain>/ (or the GitHub README)
**Privacy policy URL**: https://<domain>/privacy
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

No credentials needed. Add `https://<domain>/mcp` as a custom connector, then:

1. "Inspect this workbook: <public .xlsx URL>"
2. "In that workbook, set Budget!B2 to 40 and tell me the new Budget!D6"
3. "Explain how Budget!D12 is calculated and whether it's verified"
4. "Convert it and give me the verdict"

A sample workbook for reviewers: host `docs/sample-budget.xlsx` at a public URL (e.g. the GitHub raw link).
