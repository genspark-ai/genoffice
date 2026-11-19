#!/usr/bin/env python3
"""Build GenOffice UI Kana JP: Noto Sans JP kana/JP punctuation condensed to
Meiryo UI advances, full glyph height preserved.

Meiryo UI kana are proportional (Word probe 2026-09-03 from Word's private
meiryo.ttc: text-weighted mean 0.769em, ideographic space and comma/period
0.664em, corner brackets 0.5em) while the Hiragino fallback keeps them at 1em.
A size-adjust alias reproduces the advances but also shrinks glyph height, so
kana render at ~77% next to full-size kanji. This derivative gives every glyph
in U+3000-30FF whose Meiryo UI advance differs from 1em that exact advance and
scales the outline horizontally to fit; height, baseline and kanji (which fall
through to the chain) are untouched. Regular and Bold instances carry Meiryo
UI's regular and bold tables respectively.

Advance truth lives in tools/meiryo-ui-kana-advances.json; regenerate it from
a machine with Word installed via --dump-advances.

Usage:
  python3 tools/build-meiryo-ui-kana-font.py <NotoSansJP[wght].ttf> [outdir]
  python3 tools/build-meiryo-ui-kana-font.py --dump-advances <meiryo.ttc> <meiryob.ttc>
"""

import json
import sys
from pathlib import Path

from fontTools.pens.recordingPen import DecomposingRecordingPen
from fontTools.pens.transformPen import TransformPen
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.subset import Options, Subsetter
from fontTools.ttLib import TTCollection, TTFont
from fontTools.ttLib.woff2 import WOFF2FlavorData
from fontTools.varLib.instancer import instantiateVariableFont

ROOT = Path(__file__).resolve().parent.parent
ADVANCES = ROOT / "tools/meiryo-ui-kana-advances.json"
OUT_DIR = ROOT / "apps/docs/src/renderer/fonts"
FAMILY = "GenOffice UI Kana JP"
PS_PREFIX = "GenOfficeUIKanaJP"
WEIGHTS = {"regular": (400, "Regular"), "bold": (700, "Bold")}
# Chromium places fallback glyphs by their own ascent; mirror the Hiragino face
# these glyphs sit next to so mixed lines share one baseline geometry.
ASCENT, DESCENT = 880, -120


def dump_advances(regular_ttc: Path, bold_ttc: Path) -> None:
    def ui_face(path: Path) -> TTFont:
        for font in TTCollection(str(path)).fonts:
            if font["name"].getDebugName(4).startswith("Meiryo UI") and "Italic" not in font[
                "name"
            ].getDebugName(4):
                return font
        raise SystemExit(f"no Meiryo UI face in {path}")

    faces = {"regular": ui_face(regular_ttc), "bold": ui_face(bold_ttc)}
    table: dict[str, dict[str, float]] = {}
    for key, font in faces.items():
        upm = font["head"].unitsPerEm
        cmap = font.getBestCmap()
        hmtx = font["hmtx"]
        table[key] = {}
        for cp in range(0x3000, 0x3100):
            glyph = cmap.get(cp)
            if glyph is None:
                continue
            adv = round(hmtx[glyph][0] / upm, 4)
            if adv != 1.0:
                table[key][f"{cp:04X}"] = adv
    ADVANCES.write_text(json.dumps(table, indent=1, sort_keys=True) + "\n")
    print(f"{ADVANCES.name}: {len(table['regular'])} regular / {len(table['bold'])} bold")


def rename(font: TTFont, style: str) -> None:
    ps_name = f"{PS_PREFIX}-{style}"
    values = {
        1: FAMILY,
        2: style,
        3: f"{FAMILY} {style}",
        4: f"{FAMILY} {style}",
        6: ps_name,
        16: FAMILY,
        17: style,
        18: f"{FAMILY} {style}",
        20: ps_name,
        21: FAMILY,
        22: style,
        25: PS_PREFIX,
    }
    name = font["name"]
    for record in list(name.names):
        value = values.get(record.nameID)
        if value is not None:
            name.setName(value, record.nameID, record.platformID, record.platEncID, record.langID)
    for record in name.names:
        if record.nameID in values:
            assert "source" not in record.toUnicode().casefold(), record.toUnicode()


def condense(font: TTFont, advances: dict[str, float]) -> None:
    upm = font["head"].unitsPerEm
    cmap = font.getBestCmap()
    glyph_set = font.getGlyphSet()
    glyf = font["glyf"]
    hmtx = font["hmtx"]
    # decompose and record every outline from the untouched glyph set first: a
    # composite written back as a component of an already condensed base would
    # take the scale twice
    outlines: dict[str, tuple[DecomposingRecordingPen, float]] = {}
    for hex_cp, target in advances.items():
        glyph_name = cmap.get(int(hex_cp, 16))
        if glyph_name is None:
            continue
        source_adv = hmtx[glyph_name][0] / upm
        if source_adv == 0:
            continue
        rec = DecomposingRecordingPen(glyph_set)
        glyph_set[glyph_name].draw(rec)
        outlines[glyph_name] = (rec, target / source_adv)
    for glyph_name, (rec, scale) in outlines.items():
        pen = TTGlyphPen(None)
        rec.replay(TransformPen(pen, (scale, 0, 0, 1, 0, 0)))
        glyph = pen.glyph()
        glyph.recalcBounds(glyf)
        glyf[glyph_name] = glyph
        hmtx[glyph_name] = (round(hmtx[glyph_name][0] * scale), getattr(glyph, "xMin", 0))


def build(src: Path, out_dir: Path) -> None:
    table = json.loads(ADVANCES.read_text())
    for key, (weight, style) in WEIGHTS.items():
        font = instantiateVariableFont(TTFont(str(src), recalcTimestamp=False), {"wght": weight})
        # each weight carries only its own table; a zero-advance source (combining
        # marks) cannot be condensed to a spacing advance and falls through the chain
        cmap = font.getBestCmap()
        hmtx = font["hmtx"]
        unicodes = {
            int(cp, 16)
            for cp in table[key]
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
        condense(font, table[key])
        rename(font, style)
        font["OS/2"].usWeightClass = weight
        font["OS/2"].fsSelection |= 1 << 7
        font["hhea"].ascent = font["OS/2"].sTypoAscender = font["OS/2"].usWinAscent = ASCENT
        font["hhea"].descent = font["OS/2"].sTypoDescender = DESCENT
        font["OS/2"].usWinDescent = -DESCENT
        font["hhea"].lineGap = font["OS/2"].sTypoLineGap = 0
        font.flavor = "woff2"
        font.flavorData = WOFF2FlavorData(transformedTables=())
        out = out_dir / f"{PS_PREFIX}-{style}.woff2"
        font.save(str(out))
        print(f"{out.name}: {len(font.getGlyphOrder())} glyphs")


def main() -> None:
    args = sys.argv[1:]
    if len(args) == 3 and args[0] == "--dump-advances":
        dump_advances(Path(args[1]), Path(args[2]))
        return
    if not args:
        sys.exit(__doc__)
    build(Path(args[0]), Path(args[1]) if len(args) > 1 else OUT_DIR)


if __name__ == "__main__":
    main()
