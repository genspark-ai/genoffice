#!/usr/bin/env python3
"""Rewrite bundled CFF-outline faces with TrueType (glyf) outlines.

Chromium's PDF backend (Skia, printToPDF) cannot embed CFF-flavoured OpenType
faces as fonts: every glyph drawn with one lands in the PDF as a Type3 path
procedure (no font file, so viewers cannot match the font, select/search
reliably or print it as text). glyf faces embed as Type0/CIDFontType2. This
pass converts the cubic CFF outlines to quadratic glyf contours (cu2qu,
max error 1 font unit) and changes nothing else that layout reads:

- glyph order, cmap, hmtx/vmtx advances, OS/2, hhea, vhea, name, GSUB, GPOS,
  GDEF and BASE are carried over as-is (asserted against the input);
- hmtx left side bearings follow the new outline bounds (the glyphs do not
  move); vmtx top side bearings keep each glyph's vertical origin, which the
  dropped VORG table held for the CFF face;
- maxp becomes version 1.0, post format 2 (format 3 for CID-keyed sources,
  whose cidNNNNN names carry nothing), DSIG (now invalid) is dropped, and a
  gasp table asks for symmetric smoothing at every size, as unhinted glyf
  faces usually do.

The container is kept: a woff2 input stays woff2 (glyf/loca untransformed, as
the other bundled subsets ship, so tests/helpers/woff2-metrics.ts can read it)
and a file keeps its name. Faces that already have glyf are skipped, so the
pass is idempotent. Run it last, after the subset/metric tools
(normalize-kr-sans-hmtx.py, normalize-kr-latin-metrics.py,
scale-kr-sans-latin-ink.py, build-kr-che-latin-font.py), which edit CFF
charstrings.

Licensing: changing the outline format makes a Modified Version under SIL OFL
1.1 §3, so a face may only be converted when its family name is not a
Reserved Font Name. The default list holds only such faces; Pretendard
(Reserved Font Name "Pretendard") is deliberately absent.

Usage: python tools/convert-cff-to-glyf.py [font ...] [--out-dir DIR]
"""

import argparse
import sys
from pathlib import Path

from fontTools.pens.boundsPen import BoundsPen
from fontTools.pens.cu2quPen import Cu2QuPen
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.ttLib import TTFont, newTable
from fontTools.ttLib.woff2 import WOFF2FlavorData

# Faces the docs renderer loads whose upstream Reserved Font Name ("Source"
# for Noto CJK) is not the family name they ship under.
DEFAULT_FONTS = [
    "apps/docs/src/renderer/fonts/GenOfficeSansKR-Regular-subset.woff2",
    "apps/docs/src/renderer/fonts/GenOfficeSerifKR-Regular-subset.woff2",
    "apps/docs/src/renderer/fonts/GenOfficeCheLatinKR.woff2",
    "apps/docs/src/renderer/fonts/NotoSansCJKsc-Regular-subset.woff2",
    "apps/docs/src/renderer/fonts/NotoSerifCJKsc-Regular-subset.woff2",
]

MAX_ERR = 1.0  # font units
KEPT_TABLES = ("cmap", "OS/2", "hhea", "vhea", "name", "GSUB", "GPOS", "GDEF", "BASE")
GASP_SMOOTH_ALL = 0x000F  # gridfit | grayscale | symmetric gridfit | symmetric smoothing


def table_bytes(font: TTFont, tag: str) -> bytes | None:
    return font.getTableData(tag) if tag in font else None


def metrics(font: TTFont, tag: str) -> dict[str, int]:
    if tag not in font:
        return {}
    return {name: adv for name, (adv, _) in font[tag].metrics.items()}


def vertical_origins(font: TTFont, glyph_set) -> dict[str, int]:
    """Each glyph's vertical origin Y as the CFF face defines it."""
    if "vmtx" not in font:
        return {}
    vorg = font["VORG"] if "VORG" in font else None
    origins = {}
    for name, (_, tsb) in font["vmtx"].metrics.items():
        if vorg is not None:
            origins[name] = vorg.VOriginRecords.get(name, vorg.defaultVertOriginY)
            continue
        pen = BoundsPen(glyph_set)
        glyph_set[name].draw(pen)
        if pen.bounds is not None:
            origins[name] = tsb + round(pen.bounds[3])
    return origins


def convert(src: Path, out: Path) -> bool:
    font = TTFont(str(src), recalcTimestamp=False)
    if "glyf" in font:
        return False
    assert "CFF " in font, f"{src.name}: neither glyf nor CFF outlines"
    flavor = font.flavor

    glyph_order = font.getGlyphOrder()
    before_tables = {tag: table_bytes(font, tag) for tag in KEPT_TABLES}
    before_h = metrics(font, "hmtx")
    before_v = metrics(font, "vmtx")
    glyph_set = font.getGlyphSet()
    origins = vertical_origins(font, glyph_set)
    cff = font["CFF "].cff
    cid_keyed = hasattr(cff[cff.fontNames[0]], "ROS")

    glyphs = {}
    for name in glyph_order:
        tt_pen = TTGlyphPen(None)
        glyph_set[name].draw(Cu2QuPen(tt_pen, MAX_ERR, reverse_direction=True))
        glyphs[name] = tt_pen.glyph()

    del font["CFF "]
    for tag in ("VORG", "DSIG"):
        if tag in font:
            del font[tag]

    font["loca"] = newTable("loca")
    glyf = font["glyf"] = newTable("glyf")
    glyf.glyphOrder = glyph_order
    glyf.glyphs = glyphs
    font["head"].glyphDataFormat = 0

    hmtx = font["hmtx"]
    vmtx = font["vmtx"] if "vmtx" in font else None
    for name in glyph_order:
        glyph = glyphs[name]
        glyph.recalcBounds(glyf)
        empty = glyph.numberOfContours == 0
        adv, _ = hmtx[name]
        hmtx[name] = (adv, 0 if empty else glyph.xMin)
        if vmtx is not None and not empty and name in origins:
            vadv, _ = vmtx[name]
            vmtx[name] = (vadv, origins[name] - glyph.yMax)

    maxp = font["maxp"] = newTable("maxp")
    maxp.tableVersion = 0x00010000
    maxp.numGlyphs = len(glyph_order)
    maxp.maxZones = 1
    for field in (
        "maxTwilightPoints",
        "maxStorage",
        "maxFunctionDefs",
        "maxInstructionDefs",
        "maxStackElements",
        "maxSizeOfInstructions",
        "maxComponentElements",
    ):
        setattr(maxp, field, 0)

    post = font["post"]
    post.formatType = 3.0 if cid_keyed else 2.0
    post.extraNames = []
    post.mapping = {}
    post.glyphOrder = glyph_order

    gasp = font["gasp"] = newTable("gasp")
    gasp.version = 1
    gasp.gaspRange = {0xFFFF: GASP_SMOOTH_ALL}

    font.sfntVersion = "\x00\x01\x00\x00"
    if flavor == "woff2":
        font.flavorData = WOFF2FlavorData(transformedTables=())
    out.parent.mkdir(parents=True, exist_ok=True)
    font.save(str(out))

    check = TTFont(str(out))
    assert "glyf" in check and "CFF " not in check and "VORG" not in check, out.name
    assert check.flavor == flavor, f"{out.name}: container changed"
    assert check.sfntVersion == "\x00\x01\x00\x00", out.name
    assert check["maxp"].numGlyphs == len(glyph_order), f"{out.name}: glyph count changed"
    if not cid_keyed:
        assert check.getGlyphOrder() == glyph_order, f"{out.name}: glyph order changed"
    # post format 3 renames glyphs on reload; compare metrics by glyph index
    order = check.getGlyphOrder()
    for tag, before in (("hmtx", before_h), ("vmtx", before_v)):
        after = metrics(check, tag)
        for gid, name in enumerate(glyph_order):
            if name in before:
                assert after[order[gid]] == before[name], (
                    f"{out.name}: {tag} advance of {name} changed"
                )
    for tag, data in before_tables.items():
        assert table_bytes(check, tag) == data, f"{out.name}: {tag} changed"
    return True


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("fonts", nargs="*", type=Path)
    parser.add_argument("--out-dir", type=Path, help="write here instead of in place")
    args = parser.parse_args()
    root = Path(__file__).resolve().parent.parent
    fonts = args.fonts or [root / f for f in DEFAULT_FONTS]

    total_before = total_after = 0
    for src in fonts:
        out = args.out_dir / src.name if args.out_dir else src
        size_before = src.stat().st_size
        if not convert(src, out):
            print(f"{src.name}: already glyf, skipped")
            continue
        size_after = out.stat().st_size
        total_before += size_before
        total_after += size_after
        print(f"{src.name}: {size_before:,} -> {size_after:,} bytes")
    if total_before:
        print(f"total: {total_before:,} -> {total_after:,} bytes")


if __name__ == "__main__":
    sys.exit(main())
