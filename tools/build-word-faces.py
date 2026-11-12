#!/usr/bin/env python3
"""Generate packages/font-metrics/word-faces.json from Word for Mac's private
font stores (DFonts + CloudFonts cache). Word lays text out with these copies,
not with the macOS fonts of the same name, so the renderer needs their metrics.

Per family: hhea total/ascent/descent (em), OS/2 xAvgCharWidth (em), the
East Asian flag (OS/2 ulCodePageRange1 bits 17-21) and PANOSE. The first
Regular-ish cut wins. Requires fontTools.

Usage: python3 tools/build-word-faces.py [--legacy-seed <json>] [--out <path>]
When neither font store exists the JSON is seeded from --legacy-seed rows
tagged "source": "legacy-table" so the diff against real data stays visible.
"""

import argparse
import glob
import json
import os
import sys
import unicodedata

from fontTools.ttLib import TTFont, TTCollection

DFONTS = "/Applications/Microsoft Word.app/Contents/Resources/DFonts"
CLOUD = os.path.expanduser(
    "~/Library/Group Containers/UBF8T346G9.Office/FontCache/4/CloudFonts"
)
SYSTEM_DIRS = ["/System/Library/Fonts", "/System/Library/Fonts/Supplemental"]
# macOS on-demand font assets Word for Mac lays out with although they sit
# outside every font directory (probe 2026-10-22: NanumGothic 14.96 @10pt)
ASSET_GLOB = "/System/Library/AssetsV2/com_apple_MobileAsset_Font*/*/AssetData/Nanum*.ttc"
EA_BITS = range(17, 22)  # 932 JIS, 936 GB, 949 KR, 950 Big5, 1361 Johab
REGULARISH = ("regular", "normal", "book", "roman", "w3", "w4", "medium")


def key(name):
    return unicodedata.normalize("NFKC", name).lower().strip()


def face_rows(font, source):
    name = font["name"]
    families = set()
    for rec in name.names:
        if rec.nameID in (1, 16):
            try:
                families.add(rec.toUnicode())
            except Exception:
                pass
    sub = (name.getDebugName(17) or name.getDebugName(2) or "Regular").lower()
    upm = font["head"].unitsPerEm
    hhea = font["hhea"]
    os2 = font.get("OS/2")
    r4 = lambda v: round(v, 4)
    row = {
        "hheaTotal": r4((hhea.ascent - hhea.descent + hhea.lineGap) / upm),
        "ascent": r4(hhea.ascent / upm),
        "descent": r4(-hhea.descent / upm),
        "xAvg": r4(os2.xAvgCharWidth / upm) if os2 else None,
        "eastAsian": bool(
            os2 and any((os2.ulCodePageRange1 >> b) & 1 for b in EA_BITS)
        ),
        "panose": "".join(
            "%02x" % getattr(os2.panose, f) for f in os2.panose.__dict__ if f.startswith("b")
        )
        if os2
        else None,
        "source": source,
    }
    regular = any(t in sub for t in REGULARISH) and "italic" not in sub and "bold" not in sub
    return [(key(f), f, regular, row) for f in families]


def collect(paths, source):
    out = {}
    for p in paths:
        try:
            fonts = TTCollection(p).fonts if p.lower().endswith(".ttc") else [TTFont(p, lazy=True)]
        except Exception as e:
            print(f"skip {p}: {e}", file=sys.stderr)
            continue
        for f in fonts:
            try:
                rows = face_rows(f, source)
            except Exception as e:
                print(f"skip {p}: {e}", file=sys.stderr)
                continue
            for k, display, regular, row in rows:
                prev = out.get(k)
                if prev is None or (regular and not prev["_regular"]):
                    out[k] = {"family": display, **row, "_regular": regular}
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(os.path.dirname(__file__), "..", "packages/font-metrics/word-faces.json"))
    ap.add_argument("--legacy-seed")
    args = ap.parse_args()

    faces = collect(glob.glob(ASSET_GLOB), "asset")
    if os.path.isdir(CLOUD):
        faces.update(collect(glob.glob(os.path.join(CLOUD, "*", "*")), "cloud"))
    if os.path.isdir(DFONTS):
        # DFonts win over the cloud cache: they ship with Word itself
        faces.update(collect(glob.glob(os.path.join(DFONTS, "*")), "dfonts"))
    if all(v["source"] == "asset" for v in faces.values()) and args.legacy_seed:
        with open(args.legacy_seed) as fh:
            faces = {key(k): {"family": k, **v, "source": "legacy-table", "_regular": True} for k, v in json.load(fh).items()}
    if not faces:
        sys.exit("no Word font store found and no --legacy-seed given")

    # macOS copies Word for Mac lays out with when a name is installed (Times New
    # Roman: the system copy has lineGap 87, DFonts' has 0); only families the
    # Word store lacks or meters differently. Stands in for the main-process
    # font index where there is none (unit tests).
    system = {}
    for k, v in collect([p for d in SYSTEM_DIRS for p in glob.glob(os.path.join(d, "*.[ot]t[cf]"))], "system").items():
        w = faces.get(k)
        if w is None or w["hheaTotal"] != v["hheaTotal"]:
            v.pop("panose", None)
            system[k] = v
    for v in [*faces.values(), *system.values()]:
        v.pop("_regular", None)
    doc = {
        "_generated": "tools/build-word-faces.py (fontTools); do not hand-edit",
        "faces": dict(sorted(faces.items())),
        "system": dict(sorted(system.items())),
    }
    with open(args.out, "w") as fh:
        json.dump(doc, fh, ensure_ascii=True, indent=2)
        fh.write("\n")
    print(f"{len(faces)} word families, {len(system)} system families -> {args.out}")


if __name__ == "__main__":
    main()
