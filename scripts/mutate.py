"""Mutation test of the round-trip checker.

    .venv/bin/python scripts/mutate.py FILE... [--per-file N] [--seed S]

Applies random single edits to each file's canonical JSON and runs the
round-trip gate on the result. Every edit that changes content must be caught;
edits that change only presentation (attribute order) must not be.
"""

from __future__ import annotations

import argparse
import copy
import random
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from provenance.canon import ConversionError, read_package  # noqa: E402
from provenance.verify import round_trip  # noqa: E402


def _elements(node, path=()):
    """Yield (element, parent_children_list, index) for every element in a tree."""
    for i, child in enumerate(node["children"] if "children" in node else node["nodes"]):
        if isinstance(child, dict) and "tag" in child:
            yield child, (node["children"] if "children" in node else node["nodes"]), i
            yield from _elements(child)


def _texts(node):
    for el, _, _ in _elements(node):
        for i, child in enumerate(el["children"]):
            if isinstance(child, str):
                yield el["children"], i


def _flip(text: str, rng: random.Random) -> str:
    if not text:
        return "x"
    i = rng.randrange(len(text))
    choices = [c for c in "aZ0 \t9.&<é" if c != text[i]]
    return text[:i] + rng.choice(choices) + text[i + 1:]


CONTENT = ["text", "attr_value", "attr_drop", "attr_add", "elem_drop", "elem_add", "elem_rename",
           "elem_swap", "ns_uri", "binary_byte", "part_drop", "part_rename", "text_add"]
BENIGN = ["attr_reorder"]


def mutate(pkg: dict, kind: str, rng: random.Random) -> dict | None:
    """Return a mutated deep copy, or None if this kind does not apply to the package."""
    pkg = copy.deepcopy(pkg)
    xml_parts = [p for p in pkg["parts"] if p["kind"] == "xml"]
    if kind in ("part_drop", "part_rename"):
        part = rng.choice(pkg["parts"])
        if kind == "part_drop":
            pkg["parts"].remove(part)
        else:
            part["name"] = part["name"] + "x"
        return pkg
    if kind == "binary_byte":
        parts = [p for p in pkg["parts"] if p["kind"] == "binary"]
        if not parts:
            return None
        import base64
        part = rng.choice(parts)
        raw = bytearray(base64.b64decode(part["base64"]))
        if not raw:
            return None
        i = rng.randrange(len(raw))
        raw[i] ^= 1 << rng.randrange(8)
        part["base64"] = base64.b64encode(bytes(raw)).decode()
        return pkg
    if not xml_parts:
        return None
    tree = rng.choice(xml_parts)["xml"]
    elements = list(_elements(tree))
    if kind in ("text", "text_add"):
        texts = list(_texts(tree))
        if kind == "text":
            if not texts:
                return None
            children, i = rng.choice(texts)
            children[i] = _flip(children[i], rng)
        else:
            el = rng.choice(elements)[0]
            el["children"].insert(rng.randint(0, len(el["children"])), "q")
        return pkg
    with_attrs = [e for e, _, _ in elements if any(not k.startswith("xmlns") for k, _ in e.get("attrs", []))]
    if kind == "attr_value":
        if not with_attrs:
            return None
        el = rng.choice(with_attrs)
        attr = rng.choice([a for a in el["attrs"] if not a[0].startswith("xmlns")])
        attr[1] = _flip(attr[1], rng)
        return pkg
    if kind == "attr_drop":
        if not with_attrs:
            return None
        el = rng.choice(with_attrs)
        el["attrs"].remove(rng.choice([a for a in el["attrs"] if not a[0].startswith("xmlns")]))
        return pkg
    if kind == "attr_add":
        el = rng.choice(elements)[0]
        if any(k == "provmut" for k, _ in el.get("attrs", [])):
            return None
        el.setdefault("attrs", []).append(["provmut", "1"])
        return pkg
    if kind == "attr_reorder":
        many = [e for e, _, _ in elements if len(e.get("attrs", [])) > 1]
        if not many:
            return None
        el = rng.choice(many)
        before = list(el["attrs"])
        while el["attrs"] == before:
            rng.shuffle(el["attrs"])
        return pkg
    if kind == "elem_drop":
        nested = [(e, sib, i) for e, sib, i in elements if sib is not tree["nodes"]]
        if not nested:
            return None
        _, siblings, i = rng.choice(nested)
        del siblings[i]
        return pkg
    if kind == "elem_add":
        el = rng.choice(elements)[0]
        el["children"].insert(rng.randint(0, len(el["children"])), {"tag": "provmut", "children": []})
        return pkg
    if kind == "elem_rename":
        el = rng.choice(elements)[0]
        prefix, _, local_name = el["tag"].rpartition(":")
        el["tag"] = (prefix + ":" if prefix else "") + local_name + "X"
        return pkg
    if kind == "elem_swap":
        pairs = []
        for e, _, _ in elements:
            kids = [i for i, c in enumerate(e["children"]) if isinstance(c, dict) and "tag" in c]
            pairs += [(e, a, b) for a, b in zip(kids, kids[1:]) if e["children"][a] != e["children"][b]]
        if not pairs:
            return None
        e, a, b = rng.choice(pairs)
        e["children"][a], e["children"][b] = e["children"][b], e["children"][a]
        return pkg
    if kind == "ns_uri":
        decls = [a for e, _, _ in elements for a in e.get("attrs", []) if a[0].startswith("xmlns") and a[1]]
        if not decls:
            return None
        rng.choice(decls)[1] += "/x"
        return pkg
    raise ValueError(kind)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("files", nargs="+", type=Path)
    parser.add_argument("--per-file", type=int, default=60)
    parser.add_argument("--seed", type=int, default=0)
    args = parser.parse_args()
    rng = random.Random(args.seed)

    caught, missed, benign_ok, benign_flagged = Counter(), Counter(), 0, 0
    misses = []
    for path in args.files:
        source = path.read_bytes()
        try:
            pkg = read_package(source)
        except ConversionError:
            continue
        for n in range(args.per_file):
            kind = (CONTENT + BENIGN)[n % (len(CONTENT) + len(BENIGN))]
            mutated = mutate(pkg, kind, rng)
            if mutated is None or mutated == pkg:
                continue
            try:
                failed = round_trip(source, mutated)["status"] == "failed"
            except Exception:  # a crash still means the change did not pass silently
                failed = True
            if kind in BENIGN:
                benign_ok += not failed
                benign_flagged += failed
            elif failed:
                caught[kind] += 1
            else:
                missed[kind] += 1
                misses.append((path.name, kind))
    total = sum(caught.values()) + sum(missed.values())
    print(f"content mutations: {sum(caught.values())}/{total} caught")
    for kind in CONTENT:
        print(f"  {kind:12} caught {caught[kind]:5}  missed {missed[kind]}")
    print(f"presentation-only mutations (attribute order): {benign_ok} accepted, {benign_flagged} wrongly flagged")
    for name, kind in misses[:20]:
        print(f"  MISSED {kind} in {name}")
    return 1 if missed or benign_flagged else 0


if __name__ == "__main__":
    sys.exit(main())
