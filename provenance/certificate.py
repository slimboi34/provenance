"""Conversion certificate: what was proven, what was tested, what was not."""

from __future__ import annotations

import hashlib
import hmac
import os
from datetime import datetime, timezone

from . import __version__
from .canon import dumps, sha256

FORMAT = "provenance.certificate/1"
KEY_ENV = "PROVENANCE_SIGNING_KEY"


def _formula_layer(model: dict, compile_report: dict, replay: dict, fuzz: dict | None) -> dict:
    total = len(compile_report["compiled"]) + len(compile_report["unsupported"])
    if total == 0:
        return {"status": "none", "total": 0}
    fuzz_cells = fuzz.get("cells", {}) if fuzz and fuzz.get("status") == "ran" else {}
    cells = {}
    unverified = dict(compile_report["unsupported"])
    unverified.update({k: v for k, v in replay["failed"].items() if k not in unverified and k in replay["cells"]})
    for cell_key, status in replay["cells"].items():
        if cell_key in unverified:
            cells[cell_key] = "unverified"
        elif status == "mismatch":
            cells[cell_key] = "mismatch"
        elif fuzz_cells.get(cell_key) == "mismatch":
            # Excel's stored result was reproduced, but the second engine disagrees on
            # changed inputs. Either side can be wrong, so this blocks "verified"
            # without claiming the conversion failed.
            cells[cell_key] = "disputed"
        elif status == "no_stored_value":
            cells[cell_key] = "tested" if cell_key in fuzz_cells else "unverified"
            if cell_key not in fuzz_cells:
                unverified[cell_key] = "the file holds no stored result to compare with"
        else:
            cells[cell_key] = "reproduced+tested" if cell_key in fuzz_cells else "reproduced"
    counts: dict[str, int] = {}
    for status in cells.values():
        counts[status] = counts.get(status, 0) + 1
    mismatched = counts.get("mismatch", 0)
    status = "mismatch" if mismatched else "partial" if counts.get("unverified") or counts.get("disputed") else "verified"
    layer = {
        "status": status,
        "total": total,
        "counts": counts,
        "replay": {"what": "each formula computed from the stored values of the cells it reads and compared with "
                           "the result stored in the file; also recalculated end to end from the inputs",
                   "stored_by": replay["stored_by"] or "unknown (docProps/app.xml missing)",
                   "counts": replay["counts"], "mismatches": replay["mismatches"],
                   "end_to_end_drift": {"count": replay["drift_count"], "cells": replay["drift"],
                                        "what": "cells correct on their own whose end-to-end result differs beyond "
                                                "1e-9 relative, from floating-point rounding in earlier cells"}},
        "unverified": unverified,
        "cells": cells,
    }
    if fuzz:
        layer["fuzz"] = {k: v for k, v in fuzz.items() if k != "cells"}
        layer["fuzz"]["what"] = "inputs changed at random; compiled results compared with an independent engine"
    return layer


def build(source_name: str, source: bytes, package_json: bytes, structure: dict,
          model: dict | None, compile_report: dict | None, replay: dict | None, fuzz: dict | None,
          logic_source: str | None, notes: list[str]) -> dict:
    layers: dict = {"structure": structure}
    if model is not None:
        layers["formulas"] = _formula_layer(model, compile_report, replay, fuzz)
        macros = model["macro_parts"]
        layers["macros"] = {
            "status": "unverified" if macros else "none",
            "parts": macros,
            **({"note": "VBA is preserved byte-for-byte in the package but its behaviour is not translated or checked."} if macros else {}),
        }
    cert = {
        "format": FORMAT,
        "tool": {"name": "provenance", "version": __version__},
        "created": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "source": {"name": source_name, "sha256": sha256(source), "bytes": len(source)},
        "artifacts": {
            "package_json_sha256": sha256(package_json),
            **({"logic_py_sha256": sha256(logic_source.encode())} if logic_source is not None else {}),
        },
        "layers": layers,
        "notes": notes,
    }
    cert["verdict"] = verdict(cert)
    return seal(cert)


def verdict(cert: dict) -> dict:
    layers = cert["layers"]
    structure_ok = layers["structure"]["status"] == "proven"
    formulas = layers.get("formulas", {"status": "none"})["status"]
    macros = layers.get("macros", {"status": "none"})["status"]
    if not structure_ok or formulas == "mismatch":
        level = "failed"
    elif formulas == "partial" or macros == "unverified":
        level = "partial"
    else:
        level = "verified"
    parts = ["Structure and data: " + ("proven lossless by round trip." if structure_ok else "ROUND TRIP FAILED.")]
    if formulas != "none":
        f = layers["formulas"]
        c = f["counts"]
        done = sum(c.get(k, 0) for k in ("reproduced", "reproduced+tested", "tested", "disputed"))
        parts.append(f"Formulas: {done} of {f['total']} reproduce Excel's results"
                     + (f", {c['mismatch']} DIFFER" if c.get("mismatch") else "")
                     + (f", {c['disputed']} disputed by the second engine" if c.get("disputed") else "")
                     + (f", {c['unverified']} not verified" if c.get("unverified") else "") + ".")
        stored_by = f.get("replay", {}).get("stored_by", "")
        if "Excel" not in stored_by:
            parts.append(f"Stored results were calculated by {stored_by}, not Microsoft Excel.")
        drift = f.get("replay", {}).get("end_to_end_drift", {}).get("count", 0)
        if drift:
            parts.append(f"{drift} cells drift by floating-point rounding when recalculated end to end.")
    if macros == "unverified":
        parts.append("Macros: present, preserved, behaviour not verified.")
    return {"level": level, "summary": " ".join(parts)}


def _digest(cert: dict) -> str:
    body = {k: v for k, v in cert.items() if k not in ("seal", "signature")}
    return sha256(dumps(body))


def seal(cert: dict) -> dict:
    cert["seal"] = {"alg": "sha256", "digest": _digest(cert),
                    "note": "Checksum of this certificate. It detects accidental changes only; anyone can recompute it. "
                            "Only the signature (HMAC-SHA256, present when a signing key was set) shows who issued it."}
    signing_key = os.environ.get(KEY_ENV)
    if signing_key:
        mac = hmac.new(signing_key.encode(), cert["seal"]["digest"].encode(), hashlib.sha256).hexdigest()
        cert["signature"] = {"alg": "HMAC-SHA256", "value": mac}
    return cert


def check_seal(cert: dict) -> list[str]:
    """Problems with the certificate's own integrity (empty list = intact)."""
    problems = []
    if cert.get("seal", {}).get("digest") != _digest(cert):
        problems.append("certificate was altered after it was sealed")
    signature = cert.get("signature")
    signing_key = os.environ.get(KEY_ENV)
    if signature and signing_key:
        mac = hmac.new(signing_key.encode(), cert["seal"]["digest"].encode(), hashlib.sha256).hexdigest()
        if not hmac.compare_digest(mac, signature.get("value", "")):
            problems.append("signature does not match the signing key")
    return problems
