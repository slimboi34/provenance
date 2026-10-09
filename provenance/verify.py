"""The round-trip gate: prove the canonical JSON lost nothing.

The check never looks at the JSON. It fingerprints the original file and the
file rebuilt from the JSON with libxml2's Canonical XML (C14N 1.0, comments
kept), which is a separate implementation from the expat-based converter in
`provenance.canon`. Two XML parts with equal C14N bytes are the same XML
document: same elements, attributes, namespaces, text, comments and
processing instructions. Non-XML parts must be byte-identical.

What is *not* compared: zip metadata (compression level, timestamps), the XML
declaration, attribute order and quoting style. None of these carry content.
A part C14N cannot process (relative namespace URI) is compared by a digest
of its full parsed infoset instead.
"""

from __future__ import annotations

import io
import json
import zipfile

from lxml import etree

from .canon import XML_SUFFIXES, build_package, read_package, sha256


def _infoset(doc) -> bytes:
    """Every item of the parsed document, in order, as an unambiguous byte string.

    Used for the rare part that Canonical XML refuses (it rejects relative
    namespace URIs). Covers the same information C14N does.
    """
    items: list = []

    def walk(el, inherited):
        if not isinstance(el.tag, str):  # comment or processing instruction
            items.append([type(el).__name__, getattr(el, "target", None), el.text, el.tail])
            return
        declared = sorted((k or "", v) for k, v in el.nsmap.items() if inherited.get(k) != v)
        items.append(["start", el.tag, el.prefix, declared, sorted(el.attrib.items()), el.text])
        for child in el:
            walk(child, el.nsmap)
        items.append(["end", el.tail])

    root_el = doc.getroot()
    for sibling in reversed(list(root_el.itersiblings(preceding=True))):
        walk(sibling, {})
    walk(root_el, {})
    for sibling in root_el.itersiblings():
        walk(sibling, {})
    items[-1][-1] = None  # whitespace after the root element is not content
    return json.dumps(items, ensure_ascii=False).encode("utf-8")


def _canonical(raw: bytes) -> tuple[str, bytes] | None:
    parser = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=True)
    try:
        doc = etree.fromstring(raw, parser).getroottree()
    except etree.XMLSyntaxError:
        return None
    if doc.docinfo.doctype:
        return None
    try:
        out = io.BytesIO()
        doc.write_c14n(out, with_comments=True)
        return "xml-c14n", out.getvalue()
    except etree.C14NError:
        return "xml-infoset", _infoset(doc)


def fingerprint(data: bytes) -> list[dict]:
    """One entry per part, in package order: name, comparison method, digest."""
    entries = []
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        for info in archive.infolist():
            if info.is_dir():
                continue
            raw = archive.read(info)
            canonical = _canonical(raw) if info.filename.lower().endswith(XML_SUFFIXES) else None
            if canonical is not None:
                entries.append({"name": info.filename, "method": canonical[0], "sha256": sha256(canonical[1])})
            else:
                entries.append({"name": info.filename, "method": "bytes", "sha256": sha256(raw)})
    return entries


def round_trip(source: bytes, pkg: dict | None = None) -> dict:
    """Run the gate. Returns a report whose `status` is 'proven' or 'failed'."""
    pkg = pkg if pkg is not None else read_package(source)
    rebuilt = build_package(pkg)
    before = fingerprint(source)
    after = fingerprint(rebuilt)

    after_by_name = {e["name"]: e for e in after}
    parts = []
    for entry in before:
        other = after_by_name.get(entry["name"])
        parts.append({
            "name": entry["name"],
            "method": entry["method"],
            "source_sha256": entry["sha256"],
            "rebuilt_sha256": other["sha256"] if other else None,
            "match": bool(other) and other["method"] == entry["method"] and other["sha256"] == entry["sha256"],
        })
    problems = [f"part differs after round trip: {p['name']}" for p in parts if not p["match"]]
    if [e["name"] for e in before] != [e["name"] for e in after]:
        problems.append("part list or order changed")
    if build_package(pkg) != rebuilt:
        problems.append("rebuild is not deterministic")
    if read_package(rebuilt) != pkg:
        problems.append("canonical JSON is not a fixed point of rebuild")

    return {
        "status": "failed" if problems else "proven",
        "method": "Every XML part compared by Canonical XML 1.0 digest (libxml2; full infoset digest where C14N "
                  "does not apply), every other part by raw bytes.",
        "problems": problems,
        "parts": parts,
        "rebuilt_sha256": sha256(rebuilt),
    }
