#!/usr/bin/env python3
"""Scale GenOffice Sans KR Latin outlines to Malgun Gothic's ink boxes.

normalize-kr-sans-hmtx.py rewrote the Basic Latin advances to Malgun's values
but left Noto's narrower outlines untouched, so wide-advance letters (M, W, &)
show a visible gap after the glyph ("SAM IN", "M onitoring" — prod100r3
samples 70/85). This pass reshapes each printable Latin glyph horizontally to
Malgun's per-glyph ink width and left side bearing (measured from the local
malgun.ttf at build time; only the transformed Noto outlines ship).

Idempotent: glyphs already within 2% of the target ink width AND within
3/1000 em of the target left side bearing are skipped; a matching width with
a drifted side bearing still gets translated into place.

Usage: python3 tools/scale-kr-sans-latin-ink.py [woff2-path] [malgun-path]
"""

import sys
from pathlib import Path

from fontTools.misc.transform import Transform
from fontTools.pens.boundsPen import BoundsPen
from fontTools.pens.t2CharStringPen import T2CharStringPen
from fontTools.pens.transformPen import TransformPen
from fontTools.ttLib import TTFont

DEFAULT = "apps/docs/src/renderer/fonts/GenOfficeSansKR-Regular-subset.woff2"
MALGUN = "/Applications/Microsoft Word.app/Contents/Resources/DFonts/malgun.ttf"


def main() -> None:
    root = Path(__file__).resolve().parent.parent
    path = Path(sys.argv[1]) if len(sys.argv) > 1 else root / DEFAULT
    malgun_path = Path(sys.argv[2]) if len(sys.argv) > 2 else Path(MALGUN)

    malgun = TTFont(str(malgun_path))
    m_upm = malgun["head"].unitsPerEm
    m_cmap = malgun.getBestCmap()
    m_glyf = malgun["glyf"]
    m_hmtx = malgun["hmtx"]

    font = TTFont(str(path))
    upm = font["head"].unitsPerEm
    cmap = font.getBestCmap()
    hmtx = font["hmtx"]
    cff = font["CFF "].cff
    td = cff[cff.fontNames[0]]
    charstrings = td.CharStrings
    glyph_set = font.getGlyphSet()

    scaled = skipped = 0
    for cp in range(0x21, 0x7F):
        name = cmap.get(cp)
        m_name = m_cmap.get(cp)
        if name is None or m_name is None:
            continue
        m_glyph = m_glyf[m_name]
        if not getattr(m_glyph, "numberOfContours", 0):
            continue
        target_ink = (m_glyph.xMax - m_glyph.xMin) * upm / m_upm
        target_lsb = m_hmtx[m_name][1] * upm / m_upm

        bp = BoundsPen(glyph_set)
        glyph_set[name].draw(bp)
        if bp.bounds is None:
            continue
        x_min, _, x_max, _ = bp.bounds
        ink = x_max - x_min
        width_ok = abs(ink - target_ink) / target_ink < 0.02
        lsb_ok = abs(x_min - target_lsb) < upm * 0.003
        if ink <= 0 or (width_ok and lsb_ok):
            skipped += 1
            continue

        sx = target_ink / ink
        transform = Transform().translate(target_lsb, 0).scale(sx, 1).translate(-x_min, 0)
        adv = hmtx[name][0]
        t2_pen = T2CharStringPen(adv, glyph_set)
        glyph_set[name].draw(TransformPen(t2_pen, transform))
        # CID-keyed CFF keeps Private dicts per FDArray entry; reuse the glyph's own
        new_cs = t2_pen.getCharString(
            private=charstrings[name].private, globalSubrs=cff.GlobalSubrs
        )
        charstrings[name] = new_cs
        hmtx[name] = (adv, round(target_lsb))
        scaled += 1

    print(f"scaled {scaled} glyphs, {skipped} already at target ink")
    if scaled:
        font.save(str(path))
        print(f"saved {path}")


if __name__ == "__main__":
    main()
