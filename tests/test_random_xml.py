"""Randomised round trip: generated XML with awkward-but-legal content must
survive canon -> rebuild with an identical Canonical XML digest."""

import io
import random
import zipfile

from provenance.canon import read_package
from provenance.verify import round_trip

CHARS = ["a", "Z", "0", " ", "\t", "\n", "\r", "\r\n", "&", "<", ">", '"', "'", "]]>", "é", "中",
         "\U0001F600", "​", "﻿", "&#13;", "&#x9;", "&amp;", "--"]


def _text(rng, escape_entities=True):
    out = []
    for _ in range(rng.randint(0, 6)):
        c = rng.choice(CHARS)
        if c.startswith("&") and len(c) > 1:
            out.append(c)  # already an entity / char reference
        elif c in "&<>":
            out.append({"&": "&amp;", "<": "&lt;", ">": "&gt;"}[c])
        elif c == "]]>":
            out.append("]]&gt;")
        else:
            out.append(c)
    return "".join(out)


def _attr(rng):
    v = _text(rng).replace('"', "&quot;")
    return v


def _element(rng, depth, prefixes):
    prefix = rng.choice(prefixes)
    tag = (prefix + ":" if prefix else "") + rng.choice(["a", "b", "c-d", "e.f", "él"])
    attrs, names = [], set()
    if rng.random() < 0.2:
        new = f"p{depth}"
        attrs.append(f'xmlns:{new}="urn:{rng.choice(["x", "y", ""]) or "z"}"')
        prefixes = prefixes + [new]
    for _ in range(rng.randint(0, 3)):
        name = rng.choice(["x", "y", "z", "xml:space", "q"])
        if name not in names:
            names.add(name)
            value = rng.choice(["preserve", "default"]) if name == "xml:space" else _attr(rng)
            attrs.append(f'{name}="{value}"')
    body = []
    if depth < 4:
        for _ in range(rng.randint(0, 4)):
            roll = rng.random()
            if roll < 0.45:
                body.append(_text(rng))
            elif roll < 0.55:
                body.append(f"<![CDATA[{rng.choice(['x<y', '&amp;', ' ', 'a]b'])}]]>")
            elif roll < 0.62:
                body.append(f"<!--{rng.choice(['', ' note ', 'a-b', chr(0xe9)])}-->")
            elif roll < 0.68:
                body.append(f"<?pi{depth} {rng.choice(['', 'data', 'a=1 ?'])}?>")
            else:
                body.append(_element(rng, depth + 1, prefixes))
    open_tag = "<" + " ".join([tag] + attrs)
    return open_tag + ("/>" if not body and rng.random() < 0.5 else ">" + "".join(body) + f"</{tag}>")


def _document(rng):
    decl = rng.choice(['', '<?xml version="1.0"?>', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'])
    root = _element(rng, 0, ["", "r"]).replace("<", '<', 1)
    # declare the 'r' prefix and a default namespace on the root
    head, rest = root.split(">", 1) if not root.endswith("/>") or root.count("<") > 1 else (root[:-2], "")
    if root.endswith("/>") and root.count("<") == 1:
        root = head + ' xmlns:r="urn:r" xmlns="urn:default"/>'
    else:
        root = head.rstrip("/") + ' xmlns:r="urn:r" xmlns="urn:default">' + rest
    misc = rng.choice(["", "<!-- before -->", "<?top x?>\n"])
    return (decl + misc + root + rng.choice(["", "\n", "<!--after-->"])).encode("utf-8")


def _package(doc: bytes) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
        z.writestr("part.xml", doc)
    return buf.getvalue()


def test_random_documents_round_trip():
    rng = random.Random(1234)
    checked = 0
    for _ in range(1500):
        doc = _document(rng)
        data = _package(doc)
        pkg = read_package(data)
        part = next(p for p in pkg["parts"] if p["name"] == "part.xml")
        if part["kind"] != "xml":
            continue  # generator produced something not well-formed; kept as bytes
        report = round_trip(data, pkg)
        assert report["status"] == "proven", (doc, report["problems"])
        checked += 1
    assert checked > 1000
