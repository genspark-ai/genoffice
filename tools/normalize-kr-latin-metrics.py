#!/usr/bin/env python3
"""Normalize GenOffice Serif KR Latin metrics to Batang.

Word renders Batang-class names with the real Office face (probe 2026-08-24:
Batang/바탕 declares lay Latin with real Batang, M 0.895em); the bundled
subset kept Noto's Latin advances above it (M 0.975), so line breaks drift
from Word on those documents. GenOffice Sans KR already went through the
same treatment for Malgun (normalize-kr-sans-hmtx.py +
scale-kr-sans-latin-ink.py). GenOffice Gothic KR is out of scope: it models
the NanumGothic downloadable asset with its real metrics unmodified
(build-gothic-kr-font.py) and must stay untouched.

For every printable Basic Latin glyph (U+0020-007E, U+00A0) this rewrites the
advance to the target face's value and reshapes the outline horizontally to
the target's per-glyph ink width and left side bearing, measured from the
local Office fonts at build time; only the transformed Noto outlines ship.

Idempotent: glyphs already at the target advance, ink width (2%), and side
bearing (3/1000 em) are skipped.

Usage: python3 tools/normalize-kr-latin-metrics.py [fonts-dir]
"""

import sys
from pathlib import Path

from fontTools.misc.transform import Transform
from fontTools.pens.boundsPen import BoundsPen
from fontTools.pens.t2CharStringPen import T2CharStringPen
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.pens.transformPen import TransformPen
from fontTools.ttLib import TTFont
from fontTools.ttLib.woff2 import WOFF2FlavorData

DFONTS = Path("/Applications/Microsoft Word.app/Contents/Resources/DFonts")
# subset filename -> (target collection, font number)
TARGETS = {
    "GenOfficeSerifKR-Regular-subset.woff2": ("batang.ttc", 0),  # Batang
}


def normalize(path: Path, target: TTFont) -> None:
    t_upm = target["head"].unitsPerEm
    t_cmap = target.getBestCmap()
    t_glyf = target["glyf"]
    t_hmtx = target["hmtx"]

    font = TTFont(str(path))
    upm = font["head"].unitsPerEm
    cmap = font.getBestCmap()
    hmtx = font["hmtx"]
    is_cff = "CFF " in font
    if is_cff:
        cff = font["CFF "].cff
        charstrings = cff[cff.fontNames[0]].CharStrings
    glyph_set = font.getGlyphSet()

    def replace_outline(name: str, transform: Transform, adv: int) -> None:
        if is_cff:
            pen = T2CharStringPen(adv, glyph_set)
            glyph_set[name].draw(TransformPen(pen, transform))
            charstrings[name] = pen.getCharString(
                private=charstrings[name].private, globalSubrs=cff.GlobalSubrs
            )
        else:
            pen = TTGlyphPen(glyph_set)
            glyph_set[name].draw(TransformPen(pen, transform))
            font["glyf"][name] = pen.glyph()

    changed = skipped = 0
    # hangul advances are out of scope (Gothic KR ships 0.94em vs Gulim's
    # 1.0em — a layout-wide change needing its own wave); pin them untouched
    hangul = cmap.get(0xAC00)
    hangul_before = hmtx[hangul][0] if hangul else None
    seen: set[str] = set()
    for cp in [*range(0x20, 0x7F), 0xA0]:
        name = cmap.get(cp)
        # nbsp shares the space glyph; first (space-width) mapping wins
        t_name = t_cmap.get(cp if cp != 0xA0 else 0x20)
        if name is None or t_name is None or name in seen:
            continue
        seen.add(name)
        target_adv = round(t_hmtx[t_name][0] * upm / t_upm)
        adv_ok = hmtx[name][0] == target_adv

        t_glyph = t_glyf[t_name]
        if not getattr(t_glyph, "numberOfContours", 0):
            # no target ink (space/nbsp): advance only
            if not adv_ok:
                hmtx[name] = (target_adv, hmtx[name][1])
                changed += 1
            else:
                skipped += 1
            continue
        target_ink = (t_glyph.xMax - t_glyph.xMin) * upm / t_upm
        target_lsb = t_hmtx[t_name][1] * upm / t_upm

        bp = BoundsPen(glyph_set)
        glyph_set[name].draw(bp)
        if bp.bounds is None:
            if not adv_ok:
                hmtx[name] = (target_adv, hmtx[name][1])
                changed += 1
            continue
        x_min, _, x_max, _ = bp.bounds
        ink = x_max - x_min
        width_ok = ink > 0 and abs(ink - target_ink) / target_ink < 0.02
        # 5/1000 em: TrueType point rounding can hold the bbox a few units off
        lsb_ok = abs(x_min - target_lsb) < upm * 0.005
        if adv_ok and width_ok and lsb_ok:
            skipped += 1
            continue

        if ink > 0:
            sx = target_ink / ink
            transform = Transform().translate(target_lsb, 0).scale(sx, 1).translate(-x_min, 0)
            replace_outline(name, transform, target_adv)
        hmtx[name] = (target_adv, round(target_lsb))
        changed += 1

    assert hangul and hmtx[hangul][0] == hangul_before, "hangul must stay untouched"
    print(f"{path.name}: {changed} glyphs normalized, {skipped} already at target")
    if changed:
        # keep glyf untransformed in the woff2: tests/helpers/woff2-metrics.ts
        # cannot read the transformed form
        if not is_cff:
            font.flavorData = WOFF2FlavorData(transformedTables=())
        font.save(str(path))


def main() -> None:
    root = Path(__file__).resolve().parent.parent
    fonts_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else root / "apps/docs/src/renderer/fonts"
    for fname, (coll, num) in TARGETS.items():
        normalize(fonts_dir / fname, TTFont(str(DFONTS / coll), fontNumber=num))


if __name__ == "__main__":
    main()
