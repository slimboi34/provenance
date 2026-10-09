"""Lossless canonical JSON for OOXML packages (.xlsx / .xlsm / .docx).

An OOXML file is a zip of parts. Every XML part becomes a JSON tree that keeps
the literal element names, attribute order, namespace declarations, text,
comments and processing instructions; every other part is kept byte-for-byte
as base64. `build_package` turns that JSON back into a file.

Nothing here decides whether the conversion was faithful - that is the job of
`provenance.verify`, which uses a different XML implementation on purpose.
"""

from __future__ import annotations

import base64
import hashlib
import io
import json
import zipfile
import zlib
from xml.parsers import expat

FORMAT = "provenance.package/1"
XML_SUFFIXES = (".xml", ".rels", ".vml")
MAX_XML_DEPTH = 256           # deeper XML parts are kept as exact bytes instead of a JSON tree
MAX_UNCOMPRESSED = 2 << 30    # refuse packages that inflate past 2 GiB (zip bombs)
_ZIP_EPOCH = (1980, 1, 1, 0, 0, 0)


class ConversionError(Exception):
    """The file cannot be converted without losing information."""


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


# ---------------------------------------------------------------- XML <-> tree

def xml_to_tree(data: bytes) -> dict:
    """Parse XML into a JSON-able tree without resolving namespaces."""
    parser = expat.ParserCreate()
    parser.ordered_attributes = True
    max_depth = MAX_XML_DEPTH
    parser.buffer_text = True
    parser.buffer_size = 1 << 20
    doc: dict = {"decl": None, "nodes": []}
    stack: list[list] = [doc["nodes"]]

    def start(name, attrs):
        el: dict = {"tag": name}
        if attrs:
            el["attrs"] = [[attrs[i], attrs[i + 1]] for i in range(0, len(attrs), 2)]
        el["children"] = []
        stack[-1].append(el)
        stack.append(el["children"])
        if len(stack) > max_depth:
            raise ConversionError(f"XML nested more than {max_depth} levels deep")

    def end(_name):
        stack.pop()

    def chars(text):
        # Text outside the root element can only be whitespace and carries no
        # information (canonical XML drops it too).
        if len(stack) > 1:
            siblings = stack[-1]
            if siblings and isinstance(siblings[-1], str):
                siblings[-1] += text  # expat splits long text; one node keeps the JSON canonical
            else:
                siblings.append(text)

    def comment(text):
        stack[-1].append({"comment": text})

    def pi(target, body):
        stack[-1].append({"pi": [target, body]})

    def decl(version, _encoding, standalone):
        # The encoding is not kept: parts are always written back as UTF-8.
        doc["decl"] = [version, standalone]

    def doctype(*_args):
        raise ConversionError("XML part declares a DOCTYPE, which OOXML does not allow")

    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.CharacterDataHandler = chars
    parser.CommentHandler = comment
    parser.ProcessingInstructionHandler = pi
    parser.XmlDeclHandler = decl
    parser.StartDoctypeDeclHandler = doctype
    parser.Parse(data, True)
    return doc


def _esc_text(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\r", "&#xD;")


def _esc_attr(s: str) -> str:
    return (
        s.replace("&", "&amp;").replace("<", "&lt;").replace('"', "&quot;")
        .replace("\t", "&#x9;").replace("\n", "&#xA;").replace("\r", "&#xD;")
    )


def _write_node(node, out: list[str]) -> None:
    if isinstance(node, str):
        out.append(_esc_text(node))
    elif "tag" in node:
        out.append("<" + node["tag"])
        for name, value in node.get("attrs", ()):
            out.append(f' {name}="{_esc_attr(value)}"')
        if node["children"]:
            out.append(">")
            for child in node["children"]:
                _write_node(child, out)
            out.append(f"</{node['tag']}>")
        else:
            out.append("/>")
    elif "comment" in node:
        out.append(f"<!--{node['comment']}-->")
    else:
        target, body = node["pi"]
        out.append(f"<?{target} {body}?>" if body else f"<?{target}?>")


def tree_to_xml(doc: dict) -> bytes:
    out: list[str] = []
    if doc.get("decl"):
        version, standalone = doc["decl"]
        text = f'<?xml version="{version}" encoding="UTF-8"'
        if standalone in (0, 1):
            text += ' standalone="yes"' if standalone else ' standalone="no"'
        out.append(text + "?>\n")
    for node in doc["nodes"]:
        _write_node(node, out)
    return "".join(out).encode("utf-8")


# ------------------------------------------------------------- tree navigation

def local(tag: str) -> str:
    return tag.rsplit(":", 1)[-1]


def root(doc: dict) -> dict:
    for node in doc["nodes"]:
        if isinstance(node, dict) and "tag" in node:
            return node
    raise ConversionError("XML part has no root element")


def kids(el: dict, name: str | None = None):
    for child in el["children"]:
        if isinstance(child, dict) and "tag" in child and (name is None or local(child["tag"]) == name):
            yield child


def kid(el: dict, name: str) -> dict | None:
    return next(kids(el, name), None)


def attr(el: dict, name: str, default=None):
    """Attribute by local name (prefix ignored)."""
    for key, value in el.get("attrs", ()):
        if local(key) == name and not key.startswith("xmlns"):
            return value
    return default


def text_of(el: dict) -> str:
    return "".join(c for c in el["children"] if isinstance(c, str))


# ------------------------------------------------------------ package <-> JSON

def read_package(data: bytes) -> dict:
    """Convert the bytes of an OOXML file into the canonical package dict."""
    try:
        archive = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile as exc:
        raise ConversionError(f"not an OOXML (zip) file: {exc}") from exc
    parts = []
    seen = set()
    with archive:
        if sum(i.file_size for i in archive.infolist()) > MAX_UNCOMPRESSED:
            raise ConversionError(f"package inflates past {MAX_UNCOMPRESSED >> 30} GiB")
        for info in archive.infolist():
            if info.is_dir():
                continue
            if info.filename in seen:
                raise ConversionError(f"duplicate part name {info.filename!r}")
            seen.add(info.filename)
            try:
                raw = archive.read(info)
            except (zipfile.BadZipFile, zlib.error, EOFError, NotImplementedError, RuntimeError) as exc:
                raise ConversionError(f"part {info.filename!r} cannot be read: {exc}") from None
            part: dict = {"name": info.filename}
            tree = None
            if info.filename.lower().endswith(XML_SUFFIXES):
                try:
                    tree = xml_to_tree(raw)
                    root(tree)
                except (expat.ExpatError, ConversionError):
                    tree = None  # malformed or unusual: keep the exact bytes
            if tree is not None:
                part["kind"] = "xml"
                part["xml"] = tree
            else:
                part["kind"] = "binary"
                part["base64"] = base64.b64encode(raw).decode("ascii")
            parts.append(part)
    if "[Content_Types].xml" not in seen:
        raise ConversionError("not an OOXML package: [Content_Types].xml is missing")
    return {"format": FORMAT, "parts": parts}


def part_bytes(part: dict) -> bytes:
    if part["kind"] == "xml":
        return tree_to_xml(part["xml"])
    return base64.b64decode(part["base64"])


def build_package(pkg: dict) -> bytes:
    """Rebuild an OOXML file. The same package dict always gives the same bytes."""
    if pkg.get("format") != FORMAT:
        raise ConversionError(f"unknown package format {pkg.get('format')!r}")
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as archive:
        for part in pkg["parts"]:
            info = zipfile.ZipInfo(part["name"], date_time=_ZIP_EPOCH)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o644 << 16
            archive.writestr(info, part_bytes(part))
    return buf.getvalue()


def dumps(obj) -> bytes:
    """Canonical JSON bytes: sorted keys, no insignificant whitespace, UTF-8."""
    return json.dumps(obj, sort_keys=True, separators=(",", ":"), ensure_ascii=False).encode("utf-8")


def find_part(pkg: dict, name: str) -> dict | None:
    for part in pkg["parts"]:
        if part["name"] == name:
            return part
    return None
