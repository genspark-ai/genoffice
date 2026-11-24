#!/usr/bin/env python3
"""Build GenOffice PMincho Kana JP: Noto Serif JP kana/JP punctuation set at
MS PMincho advances, full glyph height preserved.

MS PMincho kana are proportional glyph by glyph (0.50-0.98em in Word's private
msmincho.ttc, ideographic space and comma/period 0.664em, brackets 0.5em) while
the Hiragino fallback keeps them at 1em. A size-adjust alias per advance would
match the widths but draw each kana at a different height. This derivative gives
every glyph its MS PMincho advance: punctuation keeps the side its ink sits on,
kana are centred in the new advance and narrowed only when the ink would not fit. Word synthesizes the bold cut on
the same advances, so one Regular instance serves both.

Advance truth lives in tools/ms-pmincho-kana-advances.json; regenerate it from
a machine with Word installed via --dump-advances.

Usage:
  python3 tools/build-pmincho-kana-font.py <NotoSerifJP[wght].ttf> [outdir]
  python3 tools/build-pmincho-kana-font.py --dump-advances <msmincho.ttc>
"""

import json
import sys
from pathlib import Path

from fontTools.pens.boundsPen import BoundsPen
from fontTools.pens.recordingPen import DecomposingRecordingPen
from fontTools.pens.transformPen import TransformPen
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.subset import Options, Subsetter
from fontTools.ttLib import TTCollection, TTFont
from fontTools.ttLib.woff2 import WOFF2FlavorData
from fontTools.varLib.instancer import instantiateVariableFont

ROOT = Path(__file__).resolve().parent.parent
ADVANCES = ROOT / "tools/ms-pmincho-kana-advances.json"
OUT_DIR = ROOT / "apps/docs/src/renderer/fonts"
FAMILY = "GenOffice PMincho Kana JP"
PS_NAME = "GenOfficePMinchoKanaJP-Regular"
# kana, JP punctuation and the fullwidth brackets/colons; fullwidth letters and
# digits stay with the chain
RANGES = [(0x3000, 0x30FF), (0xFF01, 0xFF0F), (0xFF1A, 0xFF20), (0xFF3B, 0xFF40), (0xFF5B, 0xFF60)]
# the Hiragino face these glyphs sit next to
ASCENT, DESCENT = 880, -120
# side bearing kept on each side of the ink, em
MIN_BEARING = 0.03


def dump_advances(ttc: Path) -> None:
    face = next(
        f for f in TTCollection(str(ttc)).fonts if f["name"].getDebugName(1) == "MS PMincho"
    )
    upm = face["head"].unitsPerEm
    cmap = face.getBestCmap()
    hmtx = face["hmtx"]
    table: dict[str, float] = {}
    for lo, hi in RANGES:
        for cp in range(lo, hi + 1):
            glyph = cmap.get(cp)
            if glyph is None:
                continue
            adv = round(hmtx[glyph][0] / upm, 4)
            if adv != 1.0:
                table[f"{cp:04X}"] = adv
    ADVANCES.write_text(json.dumps(table, indent=1, sort_keys=True) + "\n")
    print(f"{ADVANCES.name}: {len(table)} advances")


def rename(font: TTFont) -> None:
    values = {
        1: FAMILY,
        2: "Regular",
        3: f"{FAMILY} Regular",
        4: f"{FAMILY} Regular",
        6: PS_NAME,
        16: FAMILY,
        17: "Regular",
        18: f"{FAMILY} Regular",
        20: PS_NAME,
        21: FAMILY,
        22: "Regular",
        25: PS_NAME.split("-")[0],
    }
    name = font["name"]
    for record in list(name.names):
        value = values.get(record.nameID)
        if value is not None:
            name.setName(value, record.nameID, record.platformID, record.platEncID, record.langID)
    for record in name.names:
        if record.nameID in values:
            assert "source" not in record.toUnicode().casefold(), record.toUnicode()


def fit(font: TTFont, advances: dict[str, float]) -> None:
    upm = font["head"].unitsPerEm
    cmap = font.getBestCmap()
    glyph_set = font.getGlyphSet()
    glyf = font["glyf"]
    hmtx = font["hmtx"]
    # record every outline from the untouched glyph set first: a composite written
    # back as a component of an already moved base would take the transform twice
    outlines: dict[str, tuple[DecomposingRecordingPen, tuple[float, float], float]] = {}
    for hex_cp, target in advances.items():
        glyph_name = cmap.get(int(hex_cp, 16))
        if glyph_name is None or glyph_name in outlines:
            continue
        rec = DecomposingRecordingPen(glyph_set)
        glyph_set[glyph_name].draw(rec)
        bounds = BoundsPen(None)
        rec.replay(bounds)
        ink = (bounds.bounds[0], bounds.bounds[2]) if bounds.bounds else (0.0, 0.0)
        outlines[glyph_name] = (rec, ink, target * upm)
    for glyph_name, (rec, (x_min, x_max), target) in outlines.items():
        ink = x_max - x_min
        bearing = MIN_BEARING * upm
        scale = min(1.0, (target - 2 * bearing) / ink) if ink > 0 else 1.0
        if x_max <= target - bearing:
            # comma, period, closing brackets: the ink stays against the left neighbour
            shift = 0.0
        elif x_min - (upm - target) >= bearing:
            # opening brackets keep their right bearing
            shift = target - upm
        else:
            shift = (target - ink * scale) / 2 - x_min * scale
        pen = TTGlyphPen(None)
        rec.replay(TransformPen(pen, (scale, 0, 0, 1, shift, 0)))
        glyph = pen.glyph()
        glyph.recalcBounds(glyf)
        glyf[glyph_name] = glyph
        hmtx[glyph_name] = (round(target), getattr(glyph, "xMin", 0))


def build(src: Path, out_dir: Path) -> None:
    table = json.loads(ADVANCES.read_text())
    font = instantiateVariableFont(TTFont(str(src), recalcTimestamp=False), {"wght": 400})
    cmap = font.getBestCmap()
    hmtx = font["hmtx"]
    # a zero-advance source (combining marks) cannot take a spacing advance
    unicodes = {
        int(cp, 16)
        for cp in table
        if cmap.get(int(cp, 16)) is not None and hmtx[cmap[int(cp, 16)]][0] > 0
    }
    opts = Options()
    opts.layout_features = []
    opts.name_IDs = ["*"]
    opts.drop_tables += ["DSIG", "BASE", "GDEF", "GPOS", "GSUB", "vhea", "vmtx", "STAT"]
    opts.notdef_outline = True
    subsetter = Subsetter(options=opts)
    subsetter.populate(unicodes=unicodes)
    subsetter.subset(font)
    fit(font, table)
    rename(font)
    font["OS/2"].usWeightClass = 400
    font["OS/2"].fsSelection |= 1 << 7
    font["hhea"].ascent = font["OS/2"].sTypoAscender = font["OS/2"].usWinAscent = ASCENT
    font["hhea"].descent = font["OS/2"].sTypoDescender = DESCENT
    font["OS/2"].usWinDescent = -DESCENT
    font["hhea"].lineGap = font["OS/2"].sTypoLineGap = 0
    font.flavor = "woff2"
    font.flavorData = WOFF2FlavorData(transformedTables=())
    out = out_dir / f"{PS_NAME}.woff2"
    font.save(str(out))
    covered = sorted(font.getBestCmap())
    print(f"{out.name}: {len(font.getGlyphOrder())} glyphs, {len(covered)} code points")
    print("unicode-range:", ranges(covered))


def ranges(cps: list[int]) -> str:
    out: list[str] = []
    start = prev = cps[0]
    for cp in cps[1:] + [None]:
        if cp is not None and cp == prev + 1:
            prev = cp
            continue
        out.append(f"U+{start:04X}" if start == prev else f"U+{start:04X}-{prev:04X}")
        start = prev = cp
    return ", ".join(out)


def main() -> None:
    args = sys.argv[1:]
    if len(args) == 2 and args[0] == "--dump-advances":
        dump_advances(Path(args[1]))
        return
    if not args:
        sys.exit(__doc__)
    build(Path(args[0]), Path(args[1]) if len(args) > 1 else OUT_DIR)


if __name__ == "__main__":
    main()
