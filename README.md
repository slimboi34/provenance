# provenance

Converts Excel workbooks (and the structure of Word documents) into formats
agents and code can work with, and issues a certificate stating exactly what
was proven about the conversion.

```
provenance convert budget.xlsx
```

writes, into `budget.xlsx.provenance/`:

| File | What it is |
|---|---|
| `budget.package.json` | The whole file as canonical JSON. Lossless: the original can be rebuilt from it. |
| `budget.workbook.json` | Readable model: every cell's value, formula and Excel's stored result. |
| `budget.md` | Markdown view for people and agents. Derived, not the source of truth. |
| `budget_logic.py` | Every formula as a Python function; `compute(overrides)` recalculates the workbook. |
| `budget.certificate.json` | What was proven, tested, and left unverified, with a checksum (and a signature if a key is set). |

## What the certificate claims

Three layers, each with its own kind of evidence.

**Structure and data: proven.** The file is rebuilt from the JSON and every
part is compared with the original: XML parts by Canonical XML digest, all
other parts (images, VBA binaries) byte for byte. The comparison uses libxml2
while the converter uses expat, so the checker shares no code with the thing
it checks. Not compared, because they carry no content: zip compression and
timestamps, the XML declaration, attribute order.

**Formulas: reproduced and tested.** A rule-based compiler (no AI) turns each
formula into Python. Two checks follow:

- *replay*: compute each formula from the stored values of the cells it reads
  and compare with the result saved in the file, cell by cell, so each formula
  is judged on its own logic; the workbook is also recalculated end to end and
  any floating-point drift is reported separately. The certificate records
  which program saved those results (`stored_by`): matching a file saved by
  LibreOffice shows agreement with LibreOffice, not Excel. Cells on or after a
  circular reference are not checked, because Excel does not store real
  answers for them;
- *fuzz*: change the inputs at random (numbers scaled or zeroed, TRUE/FALSE
  flipped, text swapped for other text in the same column) and compare against an
  independent engine (LibreOffice if it runs, otherwise the `formulas` Python
  library) recalculating the same changed workbook.

This is strong evidence, not a mathematical proof: a formula is certified as
matching on the stored inputs and on N random cases. A formula the compiler
does not fully model is never guessed; it is listed as unverified with the
reason, along with every cell that depends on it.

**Macros: preserved, not verified.** VBA is kept byte-for-byte in the package.
Its behaviour is not translated or checked, and the verdict says so.

Verdict levels and exit codes: `verified` (0), `failed` (1: round trip failed
or a formula differs), `partial` (2: something is unverified).

## Commands

```
provenance convert FILE [-o DIR] [--fuzz N] [--seed S] [--engine auto|libreoffice|formulas|none] [--json]
provenance rebuild FILE.package.json -o FILE.xlsx
provenance check FILE.certificate.json [--source FILE.xlsx]
```

`check` re-verifies a certificate later: its checksum, the package JSON hash,
and that the file rebuilt from the JSON still matches the certified content.

The checksum only catches accidental damage: anyone can edit an unsigned
certificate and recompute it. Set `PROVENANCE_SIGNING_KEY` to add an
HMAC-SHA256 signature, which `check` verifies with the same key. Treat
unsigned certificates as reports, not proof of origin.

## Using the compiled logic

```python
import budget_logic
values, failed = budget_logic.compute({"Budget!B2": 40})
values["Budget!D6"]
```

## Use with Claude (MCP server)

`provenance-mcp` is an MCP server, so Claude can read workbooks exactly,
recalculate them with new inputs, and tell you which answers are verified.

| Tool | Does | Changes anything? |
|---|---|---|
| `inspect_workbook` | Sheets, values and formulas as a Markdown table | No |
| `convert_workbook` | Full conversion + certificate (verified / partial / failed) | Locally, writes artifacts to `output_dir` |
| `recalculate_workbook` | What-if: `{"Budget!B2": 40}` → every affected cell, old and new | No |
| `explain_cells` | A cell's formula, stored vs recomputed value, inputs, status | No |

Each tool takes the workbook as one of `path` (local only), `url` (https;
Google Sheets share links are exported as .xlsx) or `content_base64`.

**Local** (Claude Code, Claude Desktop; reads files on your machine):

```
pip install "provenance-proxy[mcp]"
claude mcp add provenance -- provenance-mcp
provenance-mcp --config      # JSON snippet for claude_desktop_config.json
```

**Hosted** (claude.ai custom connector; no sign-in):
`provenance-mcp --http` serves MCP at `/mcp`, plus `/` (docs), `/privacy`
and `/health`. In this mode the server never reads its own disk for a caller,
refuses URLs that resolve to private addresses, caps files at 25 MB, and runs
each call in a separate process limited to 120 s and 2 GB, at most 2 at once
(`PROVENANCE_MCP_TIMEOUT`, `PROVENANCE_MCP_MEMORY_MB`, `PROVENANCE_MCP_WORKERS`).
Deploy with the `Dockerfile` (Railway: `railway up`).

Live: **https://mcp-production-5fd2.up.railway.app/mcp** (runs with 90 s and
1 GB per call). Conversion holds the whole file as a tree in memory, roughly
1 GB per MB of compressed workbook, so files over about 1 MB can exceed the
hosted limit; Claude is then told to use the local server, which has no cap.

Example prompts:

1. "Here's our pricing model: https://…/pricing.xlsx. What's the total if units go to 500?"
2. "How is Summary!F12 calculated in this workbook, and is it verified?"
3. "Convert this spreadsheet and tell me whether every formula reproduces Excel's results."

## Privacy policy

The MCP server only processes the workbook and parameters in each tool call.
On the hosted server the file is held in a temporary folder for the length of
the call and then deleted. Nothing is stored, cached across calls, sold,
shared, or used for training, and logs never contain file contents. If you
give a URL, the server downloads it from its owner. Hosting is Railway.
Request logs without content are kept for up to 30 days. Contact: GitHub
issues on this repository. The full policy is served at `/privacy`.

## Not supported yet (reported as unverified, never guessed)

Array and dynamic-array formulas, structured table references, links to other
workbooks, 3-D references, range intersection, data tables, volatile functions
(`NOW`, `RAND`, `OFFSET`, `INDIRECT`), `TEXT`, and any function not in
`provenance/runtime.py`. Also: text that Excel might read as a date, time or
formatted number when used in arithmetic or criteria (`"1/1/2024"+0`,
`">=1/1/2024"`); ordering comparisons on text containing punctuation or
accents (Excel uses locale collation); ranges over 10 million cells; and
blank-sensitive functions (`COUNTBLANK`, criteria matching empty cells) over
ranges that run past the used area. Approximate-match lookups assume sorted
data, as Excel requires. Word documents get the structure layer only.

If the second engine disagrees with a formula whose stored Excel result was
reproduced, the cell is marked *disputed* (verdict: partial), because the
second engine is sometimes the one that is wrong.

## Development

```
uv venv && uv pip install -e ".[dev]" formulas
.venv/bin/python -m pytest
```

Product checks (rerun after changing the converter or runtime):

```
scripts/fetch_corpus.sh ~/provenance-corpus            # ~600 public test files, mostly saved by Excel
.venv/bin/python scripts/corpus.py ~/provenance-corpus  # convert all; crashes, round trips, replay vs stored answers
.venv/bin/python scripts/corpus.py ~/provenance-corpus --fuzz 10   # plus random inputs vs the formulas library
.venv/bin/python scripts/sweep.py                       # every function x awkward inputs vs the formulas library
.venv/bin/python scripts/mutate.py FILE...              # the round-trip check must catch every content edit
.venv/bin/python scripts/readers.py FILE...             # openpyxl sees rebuilt == original; output is deterministic
```
