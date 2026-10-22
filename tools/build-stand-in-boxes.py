#!/usr/bin/env python3
"""Generate apps/docs/src/renderer/fonts/stand-in-boxes.json: the hhea content
box (ascent / descent, em) Chromium lays each stand-in face out with.

"bundled" is every family declared in fonts.css, measured from the first
@font-face block (overrides win over the file's hhea; size-adjust scales both).
"installed" is the macOS system font snapshot the renderer consults when a
chain member is not bundled. The renderer's lead shift (--doc-lead-gap /
--doc-lead-exact) is computed from these boxes, never typed. Requires fontTools.

Usage: python3 tools/build-stand-in-boxes.py [--out <path>]
"""

import argparse
import glob
import json
import os
import re
import sys
import unicodedata

from fontTools.ttLib import TTCollection, TTFont

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
FONTS_CSS = os.path.join(ROOT, "apps/docs/src/renderer/fonts/fonts.css")
UI_FONTS = os.path.join(ROOT, "packages/ui/src/fonts")
SYSTEM_DIRS = ["/System/Library/Fonts", "/System/Library/Fonts/Supplemental"]
REGULARISH = ("regular", "normal", "book", "roman", "w3", "w4", "medium")


def key(name):
    return unicodedata.normalize("NFKC", name).lower().strip()


def r4(v):
    return round(v, 4)


def hhea_box(font):
    upm = font["head"].unitsPerEm
    hhea = font["hhea"]
    return hhea.ascent / upm, -hhea.descent / upm


def open_fonts(path):
    return TTCollection(path).fonts if path.lower().endswith(".ttc") else [TTFont(path, lazy=True)]


def family_names(font):
    out = set()
    for rec in font["name"].names:
        if rec.nameID in (1, 16):
            try:
                out.add(rec.toUnicode())
            except Exception:
                pass
    return out


def is_regular(font):
    name = font["name"]
    sub = (name.getDebugName(17) or name.getDebugName(2) or "Regular").lower()
    return any(t in sub for t in REGULARISH) and "italic" not in sub and "bold" not in sub


def collect_system():
    out = {}
    for d in SYSTEM_DIRS:
        for p in glob.glob(os.path.join(d, "*.[ot]t[cf]")):
            try:
                fonts = open_fonts(p)
            except Exception as e:
                print(f"skip {p}: {e}", file=sys.stderr)
                continue
            for f in fonts:
                try:
                    asc, desc = hhea_box(f)
                    regular = is_regular(f)
                    fams = family_names(f)
                except Exception as e:
                    print(f"skip {p}: {e}", file=sys.stderr)
                    continue
                for fam in fams:
                    k = key(fam)
                    prev = out.get(k)
                    if prev is None or (regular and not prev["_regular"]):
                        out[k] = {"ascent": r4(asc), "descent": r4(desc), "_regular": regular}
    return out


BLOCK_RE = re.compile(r"@font-face\s*\{([^}]*)\}", re.S)
DECL_RE = re.compile(r"([a-z-]+)\s*:\s*([^;]+);")


def pct(v, default=1.0):
    if v is None:
        return default
    m = re.match(r"\s*([0-9.]+)%", v)
    return float(m.group(1)) / 100 if m else default


def css_blocks():
    """first plain block per family: weight normal / style normal, no unicode-range"""
    blocks = {}
    for m in BLOCK_RE.finditer(open(FONTS_CSS, encoding="utf8").read()):
        d = {k: v.strip() for k, v in DECL_RE.findall(m.group(1))}
        fam = d.get("font-family", "").strip("'\"")
        if not fam:
            continue
        plain = (
            "unicode-range" not in d
            and d.get("font-weight", "normal") in ("normal", "400")
            and d.get("font-style", "normal") == "normal"
        )
        prev = blocks.get(fam)
        if prev is None or (plain and not prev[1]):
            blocks[fam] = (d, plain)
    return {fam: d for fam, (d, _) in blocks.items()}


def src_box(src, system):
    """box of the first resolvable src: a bundled file or a local() system face"""
    for part in src.split(","):
        part = part.strip()
        m = re.match(r"url\(\s*['\"]?([^'\")]+)['\"]?\s*\)", part)
        if m:
            url = m.group(1).split("?")[0]
            if url.startswith("@genoffice/ui/fonts/"):
                path = os.path.join(UI_FONTS, url[len("@genoffice/ui/fonts/"):])
            else:
                path = os.path.join(os.path.dirname(FONTS_CSS), url)
            if not os.path.isfile(path):
                continue
            try:
                return hhea_box(open_fonts(path)[0])
            except Exception as e:
                print(f"skip {path}: {e}", file=sys.stderr)
                continue
        m = re.match(r"local\(\s*['\"]?([^'\")]+)['\"]?\s*\)", part)
        if m:
            hit = system.get(key(m.group(1)))
            if hit:
                return hit["ascent"], hit["descent"]
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(ROOT, "apps/docs/src/renderer/fonts/stand-in-boxes.json"))
    args = ap.parse_args()

    system = collect_system()
    bundled = {}
    for fam, d in css_blocks().items():
        box = src_box(d.get("src", ""), system)
        if box is None:
            continue
        scale = pct(d.get("size-adjust"))
        asc = pct(d.get("ascent-override"), box[0]) * scale
        desc = pct(d.get("descent-override"), box[1]) * scale
        bundled[key(fam)] = {"ascent": r4(asc), "descent": r4(desc)}
    for v in system.values():
        v.pop("_regular", None)
    doc = {
        "_generated": "tools/build-stand-in-boxes.py (fontTools); do not hand-edit",
        "bundled": dict(sorted(bundled.items())),
        "installed": dict(sorted(system.items())),
    }
    with open(args.out, "w") as fh:
        json.dump(doc, fh, ensure_ascii=True, separators=(",", ":"))
        fh.write("\n")
    print(f"{len(bundled)} bundled, {len(system)} installed families -> {args.out}")


if __name__ == "__main__":
    main()
