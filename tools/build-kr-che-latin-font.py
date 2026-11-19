#!/usr/bin/env python3
"""Build GenOffice Che Latin KR: half-width Latin for -Che fixed-pitch faces.

Word renders BatangChe/GulimChe/DotumChe/GungsuhChe declares with the real
Office faces, whose Latin is fixed-pitch at exactly 0.5 em (probe 2026-08-24);
our KR chains laid those runs with proportional Latin instead, drifting line
breaks (prod100r3/50 uses DotumChe on 3.7k runs). This builds a tiny
ASCII-only face from the GenOffice Sans KR outlines (Noto-derived, OFL),
advances forced to 0.5 em and each glyph reshaped to DotumChe's per-glyph ink
box measured from the local Office font at build time. Only the transformed
Noto outlines ship.

Usage: python3 tools/build-kr-che-latin-font.py [out.woff2]
"""

import sys
from pathlib import Path

from fontTools.misc.transform import Transform
from fontTools.pens.boundsPen import BoundsPen
from fontTools.pens.t2CharStringPen import T2CharStringPen
from fontTools.pens.transformPen import TransformPen
from fontTools.subset import Options, Subsetter
from fontTools.ttLib import TTFont

SOURCE = "apps/docs/src/renderer/fonts/GenOfficeSansKR-Regular-subset.woff2"
DEFAULT_OUT = "apps/docs/src/renderer/fonts/GenOfficeCheLatinKR.woff2"
DOTUMCHE = ("/Applications/Microsoft Word.app/Contents/Resources/DFonts/gulim.ttc", 3)
FAMILY = "GenOffice Che Latin KR"
PS_NAME = "GenOfficeCheLatinKR"


def main() -> None:
    root = Path(__file__).resolve().parent.parent
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else root / DEFAULT_OUT

    target = TTFont(DOTUMCHE[0], fontNumber=DOTUMCHE[1])
    t_upm = target["head"].unitsPerEm
    t_cmap = target.getBestCmap()
    t_glyf = target["glyf"]
    t_hmtx = target["hmtx"]

    font = TTFont(str(root / SOURCE))
    upm = font["head"].unitsPerEm
    half = round(upm / 2)

    subsetter = Subsetter(options=Options(notdef_outline=True, glyph_names=False))
    subsetter.populate(unicodes=list(range(0x20, 0x7F)))
    subsetter.subset(font)

    cmap = font.getBestCmap()
    hmtx = font["hmtx"]
    cff = font["CFF "].cff
    charstrings = cff[cff.fontNames[0]].CharStrings
    glyph_set = font.getGlyphSet()

    for cp in range(0x20, 0x7F):
        name = cmap.get(cp)
        t_name = t_cmap.get(cp)
        if name is None or t_name is None:
            continue
        t_glyph = t_glyf[t_name]
        bp = BoundsPen(glyph_set)
        glyph_set[name].draw(bp)
        if not getattr(t_glyph, "numberOfContours", 0) or bp.bounds is None:
            hmtx[name] = (half, hmtx[name][1])
            continue
        target_ink = (t_glyph.xMax - t_glyph.xMin) * upm / t_upm
        target_lsb = t_hmtx[t_name][1] * upm / t_upm
        x_min, _, x_max, _ = bp.bounds
        ink = x_max - x_min
        if ink > 0:
            sx = target_ink / ink
            transform = Transform().translate(target_lsb, 0).scale(sx, 1).translate(-x_min, 0)
            pen = T2CharStringPen(half, glyph_set)
            glyph_set[name].draw(TransformPen(pen, transform))
            charstrings[name] = pen.getCharString(
                private=charstrings[name].private, globalSubrs=cff.GlobalSubrs
            )
        hmtx[name] = (half, round(target_lsb))

    for rec in list(font["name"].names):
        if rec.nameID in (1, 3, 4, 16):
            value = FAMILY if rec.nameID in (1, 16) else (
                PS_NAME if rec.nameID == 3 else FAMILY
            )
            font["name"].setName(value, rec.nameID, rec.platformID, rec.platEncID, rec.langID)
        elif rec.nameID == 6:
            font["name"].setName(PS_NAME, 6, rec.platformID, rec.platEncID, rec.langID)

    font.flavor = "woff2"
    font.save(str(out))
    print(f"built {out} ({out.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
