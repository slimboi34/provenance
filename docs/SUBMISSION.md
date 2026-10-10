# Connectors Directory submission: Provenance

Portal: https://claude.ai/directory/manage → Submit new → MCP connector.
Requirements: https://claude.com/docs/connectors/building/submission
Escalations: mcp-review@anthropic.com

## Checklist

- [x] Server deployed and reachable at `https://mcp-production-5fd2.up.railway.app/mcp` (Railway, `railway up` from this folder)
- [x] `https://mcp-production-5fd2.up.railway.app/` (docs) and `https://mcp-production-5fd2.up.railway.app/privacy` load
- [x] Added in Claude as a custom connector and all nine tools called through it (2026-10-10, reviewer script steps 1–6 on the sample workbook: all `reproduced`, verdict `verified`)
- [x] Public repo `slimboi34/provenance` (the docs, privacy and support links point there)
- [x] Icon: 512×512 PNG at `docs/icon.png`
- [x] Every tool has a `title` and `readOnlyHint` (all nine are read-only in hosted mode)
- [x] No authentication needed: the server only processes files the user sends; no account data
- [x] Privacy policy: README section + `/privacy`

## Listing (paste into the portal)

**Name** (≤100): Provenance: Verified Excel

**One-liner** (≤200): Read Excel workbooks exactly, recalculate them with new inputs, and know which answers are verified against Excel's own results.

**Description** (≤2,000):

> Provenance gives Claude exact access to Excel workbooks (.xlsx, .xlsm). Instead of eyeballing a spreadsheet, Claude reads every cell and formula, recalculates the workbook in Python with Excel's semantics, and reports which results are verified.
>
> - Inspect and find: sheets, values and formulas as a table; search for any label, number or formula.
> - Explain and trace: how a cell is calculated, what feeds it, what depends on it, and whether it is verified.
> - What-if: change input cells and get every affected result, before and after, without Excel.
> - Goal seek: find the input that makes a result hit a target ("what price gets profit to 10,000?").
> - Sensitivity tables: one- or two-way data tables over ranges of inputs.
> - Compare: every changed value, formula and sheet between two versions of a workbook.
> - Convert and certify: lossless canonical JSON plus a Python module of every formula, with a certificate. Structure is proven by rebuilding the file, each formula is replayed against the result Excel stored, and the formulas are fuzz-tested against an independent engine.
>
> Formulas Provenance does not fully model are labelled unverified with the reason, never guessed, so Claude can tell you which numbers to trust.
>
> Give Claude a workbook by dropping it on the upload page (you get a file id to paste into your chat; the file is deleted after an hour) or by sharing a public https link (Google Sheets share links work). The hosted server handles typical workbooks (up to about 1 MB compressed); larger files run with the free local version. Nothing is kept after use.

**Categories**: Data & Analytics; Productivity; Finance

**Documentation URL**: https://mcp-production-5fd2.up.railway.app/ (or the GitHub README)
**Privacy policy URL**: https://mcp-production-5fd2.up.railway.app/privacy
**Support contact**: https://github.com/slimboi34/provenance/issues

## Use cases

- Answer questions about a spreadsheet using exact cell values instead of a lossy read
- What-if analysis, goal seeking and sensitivity tables on financial, pricing and budget models
- See what changed between two versions of a model
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
2. "Where is the tax figure, and how is it calculated? Trace it back to its inputs."
3. "Set Budget!B2 to 40 and tell me the new total."
4. "What quantity of widgets (B2) makes the total (D6) exactly 100?"
5. "Show how the tax (D12) changes as the VAT rate (Rates!B1) goes from 0% to 30%."
6. "Compare it with itself" (or with any edited copy) and "Convert it and give me the verdict."
7. Upload flow: open https://mcp-production-5fd2.up.railway.app/upload, drop the sample file, paste the line it gives into the chat, and ask anything above.

Sample workbook URL for reviewers: https://raw.githubusercontent.com/slimboi34/provenance/main/docs/sample-budget.xlsx
