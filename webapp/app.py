"""Provenance web proof of concept.

    .venv/bin/python webapp/app.py            # http://127.0.0.1:8795

Upload a workbook, get the verdict, the certificate and every converted file.
Uploads are deleted as soon as they are converted; results are deleted after
RETENTION_MINUTES. Local proof of concept: one process, synchronous conversion.
"""

from __future__ import annotations

import io
import logging
import re
import secrets
import shutil
import sys
import tempfile
import threading
import time
import warnings
import zipfile
from collections import Counter
from pathlib import Path

from flask import Flask, abort, redirect, render_template_string, request, send_file, url_for

ROOT = Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT), str(ROOT / "tests")]

from provenance.canon import ConversionError  # noqa: E402
from provenance.cli import convert  # noqa: E402

logging.disable(logging.CRITICAL)
warnings.simplefilter("ignore")

RETENTION_MINUTES = 30
ALLOWED = {".xlsx", ".xlsm", ".docx"}
WORK = Path(tempfile.mkdtemp(prefix="provenance-web-"))
RESULTS: dict[str, dict] = {}
LOCK = threading.Lock()

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024


# ----------------------------------------------------------------- conversion

def _cleanup() -> None:
    cutoff = time.time() - RETENTION_MINUTES * 60
    with LOCK:
        for token in [t for t, r in RESULTS.items() if r["created"] < cutoff]:
            shutil.rmtree(RESULTS.pop(token)["dir"], ignore_errors=True)


def _run(name: str, data: bytes, fuzz: int) -> str:
    _cleanup()
    token = secrets.token_urlsafe(12)
    folder = WORK / token
    folder.mkdir()
    stem = re.sub(r"[^A-Za-z0-9._ -]+", "_", Path(name).stem).strip(" .") or "workbook"
    upload = folder / (stem[:80] + Path(name).suffix.lower())
    upload.write_bytes(data)
    out = folder / "out"
    started = time.time()
    try:
        cert = convert(upload, out, fuzz_cases=fuzz, engine="formulas" if fuzz else "none")
        error = None
    except ConversionError as exc:
        cert, error = None, str(exc)
    finally:
        upload.unlink(missing_ok=True)  # the original is never kept
    with LOCK:
        RESULTS[token] = {"dir": folder, "out": out, "name": name, "cert": cert, "error": error,
                          "seconds": time.time() - started, "created": time.time()}
    return token


def _sample() -> bytes:
    from conftest import make_xlsx
    return make_xlsx({
        "Budget": {
            "A1": "Item", "B1": "Qty", "C1": "Price", "D1": "Total",
            "A2": "Widget", "B2": 4, "C2": 2.5, "D2": ("B2*C2", 10),
            "A3": "Gadget", "B3": 10, "C3": 1.2, "D3": ("B3*C3", 12),
            "A4": "Gizmo", "B4": 0, "C4": 99.99, "D4": ("B4*C4", 0),
            "A6": "Sum", "D6": ("SUM(D2:D4)", 22),
            "A7": "Average price", "D7": ("AVERAGE(C2:C4)", 34.56333333333333),
            "A8": "Big order?", "D8": ('IF(D6>20,"yes","no")', "yes"),
            "A9": "Per unit", "D9": ("D6/SUM(B2:B4)", 1.5714285714285714),
            "A10": "Gadget price", "D10": ('VLOOKUP("Gadget",A2:C4,3,FALSE)', 1.2),
            "A11": "Label", "D11": ('A2&" x"&B2', "Widget x4"),
            "A12": "Tax", "D12": ("ROUND(D6*Rates!B1,2)", 4.4),
            "A13": "Safe", "D13": ("IFERROR(D6/B4,-1)", -1),
            "A14": "Margin text", "D14": ('TEXT(D6,"0.00")', "22.00"),
        },
        "Rates": {"A1": "VAT", "B1": 0.2},
    })


# ------------------------------------------------------------------- summary

STATUS_LABELS = [
    ("reproduced+tested", "Matched Excel + random tests", "ok"),
    ("reproduced", "Matched Excel's answer", "ok"),
    ("tested", "Passed random tests", "ok"),
    ("disputed", "Disputed by 2nd engine", "warn"),
    ("unverified", "Not verified", "muted"),
    ("mismatch", "Different from Excel", "bad"),
]


def _summary(r: dict) -> dict:
    cert = r["cert"]
    layers = cert["layers"]
    structure = layers["structure"]
    formulas = layers.get("formulas", {})
    counts = formulas.get("counts", {})
    reasons = Counter()
    for reason in formulas.get("unverified", {}).values():
        reason = re.sub(r"depends on \S+, which is unverified: ", "", reason)
        reasons[reason] += 1
    fuzz = formulas.get("fuzz") or {}
    files = sorted(p.name for p in r["out"].iterdir()) if r["out"].exists() else []
    preview = ""
    md = next((r["out"] / f for f in files if f.endswith(".md")), None)
    if md:
        preview = "\n".join(md.read_text(encoding="utf-8").splitlines()[:40])
    return {
        "verdict": cert["verdict"],
        "parts": len(structure["parts"]),
        "structure": structure["status"],
        "total": formulas.get("total", 0),
        "tiles": [(label, counts.get(key, 0), tone) for key, label, tone in STATUS_LABELS if counts.get(key)],
        "reasons": reasons.most_common(6),
        "mismatches": formulas.get("replay", {}).get("mismatches", [])[:10],
        "stored_by": formulas.get("replay", {}).get("stored_by"),
        "fuzz": fuzz,
        "macros": layers.get("macros", {}).get("status") == "unverified",
        "notes": [n for n in cert["notes"] if "LibreOffice" not in n][:5],
        "files": files,
        "preview": preview,
        "sha": cert["source"]["sha256"],
    }


# --------------------------------------------------------------------- routes

@app.get("/")
def index():
    return render_template_string(PAGE, view="upload", retention=RETENTION_MINUTES)


@app.post("/convert")
def upload():
    f = request.files.get("file")
    if not f or not f.filename:
        return render_template_string(PAGE, view="upload", retention=RETENTION_MINUTES, error="Choose a file first.")
    if Path(f.filename).suffix.lower() not in ALLOWED:
        return render_template_string(PAGE, view="upload", retention=RETENTION_MINUTES,
                                      error="Upload an .xlsx, .xlsm or .docx file.")
    fuzz = 10 if request.form.get("fuzz") else 0
    return redirect(url_for("result", token=_run(f.filename, f.read(), fuzz)))


@app.post("/sample")
def sample():
    return redirect(url_for("result", token=_run("sample-budget.xlsx", _sample(), 10)))


@app.get("/r/<token>")
def result(token):
    r = RESULTS.get(token) or abort(404)
    if r["error"]:
        return render_template_string(PAGE, view="upload", retention=RETENTION_MINUTES,
                                      error=f"{r['name']} could not be converted: {r['error']}")
    return render_template_string(PAGE, view="result", r=r, s=_summary(r), token=token, retention=RETENTION_MINUTES)


@app.get("/r/<token>/file/<name>")
def download(token, name):
    r = RESULTS.get(token) or abort(404)
    path = r["out"] / name
    if "/" in name or not path.is_file():
        abort(404)
    return send_file(path, as_attachment=True, download_name=name)


@app.get("/r/<token>/all.zip")
def download_all(token):
    r = RESULTS.get(token) or abort(404)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for p in sorted(r["out"].iterdir()):
            z.write(p, p.name)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=Path(r["name"]).stem + ".provenance.zip")


@app.errorhandler(413)
def too_big(_):
    return render_template_string(PAGE, view="upload", retention=RETENTION_MINUTES, error="Files up to 25 MB."), 413


PAGE = r"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Provenance · verified spreadsheet conversion</title>
<style>
  :root { --ink:#0f172a; --sub:#475569; --line:#e2e8f0; --bg:#f8fafc; --card:#fff; --brand:#0f766e;
          --ok:#15803d; --okbg:#ecfdf5; --warn:#b45309; --warnbg:#fffbeb; --bad:#b91c1c; --badbg:#fef2f2; --muted:#64748b; --mutedbg:#f1f5f9; }
  * { box-sizing:border-box } body { margin:0; font:15px/1.55 -apple-system,"Segoe UI",Helvetica,Arial,sans-serif; color:var(--ink); background:var(--bg) }
  header { background:#0b1220; color:#e2e8f0 } .bar { max-width:1040px; margin:auto; padding:18px 24px; display:flex; align-items:center; gap:12px }
  .logo { width:30px; height:30px; border-radius:8px; background:linear-gradient(135deg,#14b8a6,#0e7490); display:grid; place-items:center; font-weight:800; color:#fff }
  .bar b { font-size:17px; letter-spacing:.2px } .bar span { color:#94a3b8; font-size:13px; margin-left:auto }
  main { max-width:1040px; margin:auto; padding:32px 24px 64px }
  h1 { font-size:34px; line-height:1.15; margin:8px 0 10px } h2 { font-size:18px; margin:0 0 12px } .lead { color:var(--sub); font-size:17px; max-width:680px }
  .card { background:var(--card); border:1px solid var(--line); border-radius:14px; padding:22px; margin-top:20px }
  .drop { border:2px dashed #cbd5e1; border-radius:12px; padding:30px; text-align:center; background:#f8fafc }
  .drop input { margin:10px 0 } .row { display:flex; gap:12px; align-items:center; flex-wrap:wrap }
  button, .btn { font:inherit; border:0; border-radius:10px; padding:11px 18px; cursor:pointer; text-decoration:none; display:inline-block }
  .primary { background:var(--brand); color:#fff; font-weight:600 } .ghost { background:#fff; color:var(--ink); border:1px solid var(--line) }
  .steps { display:grid; grid-template-columns:repeat(4,1fr); gap:12px; margin-top:20px }
  .step { background:#fff; border:1px solid var(--line); border-radius:12px; padding:14px } .step b { display:block; margin-bottom:4px } .step small { color:var(--sub) }
  .err { background:var(--badbg); color:var(--bad); border:1px solid #fecaca; padding:12px 16px; border-radius:10px; margin-top:16px }
  .verdict { border-radius:14px; padding:22px; color:#fff } .verdict.verified { background:var(--ok) } .verdict.partial { background:#a16207 } .verdict.failed { background:var(--bad) }
  .verdict .lvl { font-size:13px; text-transform:uppercase; letter-spacing:1.5px; opacity:.85 } .verdict .big { font-size:26px; font-weight:700; margin:2px 0 6px }
  .tiles { display:grid; grid-template-columns:repeat(auto-fit,minmax(170px,1fr)); gap:12px }
  .tile { border-radius:12px; padding:14px; border:1px solid var(--line) } .tile .n { font-size:26px; font-weight:700 } .tile small { color:var(--sub) }
  .tile.ok { background:var(--okbg); border-color:#bbf7d0 } .tile.ok .n { color:var(--ok) } .tile.warn { background:var(--warnbg); border-color:#fde68a } .tile.warn .n { color:var(--warn) }
  .tile.bad { background:var(--badbg); border-color:#fecaca } .tile.bad .n { color:var(--bad) } .tile.muted { background:var(--mutedbg) } .tile.muted .n { color:var(--muted) }
  .tile.info { background:#eff6ff; border-color:#bfdbfe } .tile.info .n { color:#1d4ed8 }
  table { width:100%; border-collapse:collapse; font-size:14px } td, th { text-align:left; padding:8px 6px; border-bottom:1px solid var(--line); vertical-align:top } th { color:var(--sub); font-weight:600 }
  code, pre { font-family:ui-monospace,SFMono-Regular,Menlo,monospace; font-size:13px } pre { background:#0b1220; color:#e2e8f0; padding:16px; border-radius:10px; overflow:auto; max-height:420px }
  .files a { display:inline-block; margin:4px 8px 4px 0; padding:7px 12px; border:1px solid var(--line); border-radius:8px; color:var(--ink); text-decoration:none; background:#fff }
  .files a:hover, .ghost:hover { border-color:var(--brand) } .muted { color:var(--sub); font-size:13px } .grid2 { display:grid; grid-template-columns:1fr 1fr; gap:20px }
  @media (max-width:800px) { .steps, .grid2 { grid-template-columns:1fr } }
</style></head><body>
<header><div class="bar"><div class="logo">P</div><b>Provenance</b><span>proof of concept · runs locally</span></div></header>
<main>
{% if view == "upload" %}
  <h1>Prove your spreadsheet survived the move.</h1>
  <p class="lead">Convert an Excel workbook into JSON, readable text and Python, and get a certificate showing exactly what was proven, what was tested and what wasn't checked.</p>
  {% if error %}<div class="err">{{ error }}</div>{% endif %}
  <div class="card">
    <form action="{{ url_for('upload') }}" method="post" enctype="multipart/form-data">
      <div class="drop">
        <h2>Upload a workbook</h2>
        <div class="muted">.xlsx, .xlsm or .docx · up to 25 MB</div>
        <input type="file" name="file" accept=".xlsx,.xlsm,.docx" required><br>
        <label class="muted"><input type="checkbox" name="fuzz" checked> Also run 10 random-input tests against a second engine</label>
      </div>
      <div class="row" style="margin-top:14px">
        <button class="primary" type="submit">Convert and verify</button>
        <span class="muted">Your file is deleted right after conversion; results are deleted after {{ retention }} minutes.</span>
      </div>
    </form>
    <form action="{{ url_for('sample') }}" method="post" style="margin-top:10px"><button class="ghost" type="submit">No file handy? Try a sample workbook</button></form>
  </div>
  <div class="steps">
    <div class="step"><b>1 · Convert</b><small>JSON copy of the whole file, a text view, formulas as Python.</small></div>
    <div class="step"><b>2 · Prove nothing was lost</b><small>Rebuild the file from the JSON; every part must match exactly.</small></div>
    <div class="step"><b>3 · Check the answers</b><small>Each formula must give Excel's saved answer, then pass random tests.</small></div>
    <div class="step"><b>4 · Certify</b><small>Proven, tested or not verified, per cell. Never guessed.</small></div>
  </div>
{% else %}
  <div class="row" style="justify-content:space-between"><div class="muted">{{ r.name }} · converted in {{ '%.1f' % r.seconds }}s</div>
    <a class="btn ghost" href="{{ url_for('index') }}">Convert another file</a></div>
  <div class="verdict {{ s.verdict.level }}" style="margin-top:12px">
    <div class="lvl">Verdict</div><div class="big">{{ s.verdict.level|capitalize }}</div><div>{{ s.verdict.summary }}</div>
  </div>

  <div class="card"><h2>What was checked</h2>
    <div class="tiles">
      <div class="tile {{ 'ok' if s.structure == 'proven' else 'bad' }}"><div class="n">{{ s.parts }}</div><small>file parts · lossless round trip {{ s.structure }}</small></div>
      {% if s.total %}<div class="tile info"><div class="n">{{ s.total }}</div><small>formulas found</small></div>{% endif %}
      {% for label, n, tone in s.tiles %}<div class="tile {{ tone }}"><div class="n">{{ n }}</div><small>{{ label }}</small></div>{% endfor %}
      {% if s.fuzz.get('status') == 'ran' %}<div class="tile ok"><div class="n">{{ s.fuzz.counts.exact + s.fuzz.counts.close }}</div><small>random-test comparisons agreed ({{ s.fuzz.counts.mismatch }} disagreed)</small></div>{% endif %}
    </div>
    <p class="muted" style="margin-top:14px">Saved answers in this file were calculated by: <b>{{ s.stored_by or 'unknown' }}</b>{% if s.macros %} · contains macros (preserved, not verified){% endif %}</p>
  </div>

  {% if s.mismatches %}<div class="card"><h2>Formulas that differ from the saved answer</h2>
    <table><tr><th>Cell</th><th>Formula</th><th>Saved</th><th>Ours</th></tr>
    {% for m in s.mismatches %}<tr><td><code>{{ m.cell }}</code></td><td><code>={{ m.formula }}</code></td><td>{{ m.excel }}</td><td>{{ m.compiled }}</td></tr>{% endfor %}</table></div>{% endif %}

  <div class="grid2">
    {% if s.reasons %}<div class="card"><h2>Why some formulas weren't verified</h2>
      <table>{% for reason, n in s.reasons %}<tr><td>{{ reason }}</td><td style="text-align:right"><b>{{ n }}</b></td></tr>{% endfor %}</table>
      <p class="muted">Never guessed: these cells are reported instead of being given a value that might be wrong.</p></div>{% endif %}
    <div class="card"><h2>Download</h2>
      <div class="files">{% for f in s.files %}<a href="{{ url_for('download', token=token, name=f) }}">{{ f }}</a>{% endfor %}</div>
      <p style="margin-top:12px"><a class="btn primary" href="{{ url_for('download_all', token=token) }}">Download everything (.zip)</a></p>
      <p class="muted">Source SHA-256 <code>{{ s.sha[:16] }}…</code> · results deleted after {{ retention }} minutes.</p>
      {% for n in s.notes %}<p class="muted">Note: {{ n }}</p>{% endfor %}
    </div>
  </div>

  {% if s.preview %}<div class="card"><h2>Readable view (preview)</h2><pre>{{ s.preview }}</pre></div>{% endif %}
{% endif %}
</main></body></html>"""


if __name__ == "__main__":
    try:
        app.run(host="127.0.0.1", port=8795, debug=False)
    finally:
        shutil.rmtree(WORK, ignore_errors=True)
