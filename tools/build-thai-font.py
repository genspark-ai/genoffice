#!/usr/bin/env python3
"""Build the GenOffice Thai faces from upstream Noto Sans Thai, metric-fitted to
the Thai faces Word for Mac lays out with.

Word ships Cordia New and TH SarabunPSK in its DFonts; a declared Thai family
it does not have (TH SarabunIT9, Angsana New, the *UPC families...) renders
as Cordia New (probe 2026-11-03: prod sample 059, 16pt text paces 18.4pt =
Cordia New hhea 1.147). macOS has none of these; Chromium's Thai fallback
(Thonburi) advances ~1.75x wider and its glyphs are drawn for a 1em body, so
every line wrapped early and pages overflowed.

These faces take Noto Sans Thai's outlines and fit them to the Word face:
every glyph is scaled horizontally to the Word advance of its code point
(variant and ligature glyphs follow the code points in their names, marks and
the rest take the median letter factor), one uniform vertical factor (median
consonant ink-height ratio, ~0.72) brings the 1em Noto design down to the
small Thai design size, GPOS mark anchors move with their glyph, and hhea/OS2
take the Word face's line box. Hinting is dropped. An earlier build changed
only the em (scale_upem keeps the em-relative outline) and drew 1em Noto
letters over Word's 0.4em advances, piling every cluster onto the previous.

Renamed per OFL 1.1 ("Noto" is a Reserved Font Name; outlines and advances are
modified).
Upstream: https://github.com/notofonts/notofonts.github.io
          fonts/NotoSansThai/hinted/ttf/NotoSansThai-{Regular,Bold}.ttf

Usage: python3 tools/build-thai-font.py <NotoSansThai-*.ttf> <word-face.ttf|ttc[#index]> <family> [Regular|Bold] [out.woff2]
  e.g. build-thai-font.py NotoSansThai-Regular.ttf cordia.ttc#0 "GenOffice Cordia"
       build-thai-font.py NotoSansThai-Bold.ttf cordia.ttc#1 "GenOffice Cordia" Bold
"""

import re
import statistics
import sys
from pathlib import Path

from fontTools import subset
from fontTools.pens.boundsPen import BoundsPen
from fontTools.pens.recordingPen import DecomposingRecordingPen
from fontTools.pens.transformPen import TransformPen
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.ttLib import TTCollection, TTFont
from fontTools.ttLib.woff2 import WOFF2FlavorData

OUT_DIR = Path("apps/docs/src/renderer/fonts")
THAI_BLOCK = range(0x0E00, 0x0E80)
CONSONANTS = range(0x0E01, 0x0E2F)
GLYPH_CPS = re.compile(r"^uni((?:[0-9A-F]{4})+)(?:\.|$)")


def open_face(spec: str) -> TTFont:
    path, _, index = spec.partition("#")
    if path.lower().endswith(".ttc"):
        return TTCollection(path).fonts[int(index or 0)]
    return TTFont(path)


def rename(font: TTFont, family: str, style: str) -> None:
    ps_name = family.replace(" ", "") + "-" + style
    full = family if style == "Regular" else f"{family} {style}"
    name = font["name"]
    for rec in list(name.names):
        if rec.nameID in (1, 16):
            rec.string = family
        elif rec.nameID in (2, 17):
            rec.string = style
        elif rec.nameID in (3, 4):
            rec.string = full
        elif rec.nameID == 6:
            rec.string = ps_name
        elif rec.nameID in (7, 8, 9, 11, 12, 13, 14):
            name.removeNames(nameID=rec.nameID)


def ink_height(font: TTFont, gname: str) -> float | None:
    pen = BoundsPen(font.getGlyphSet())
    font.getGlyphSet()[gname].draw(pen)
    return None if pen.bounds is None else pen.bounds[3] - pen.bounds[1]


def name_codepoints(gname: str) -> list[int]:
    m = GLYPH_CPS.match(gname)
    if not m:
        return []
    hexes = m.group(1)
    return [int(hexes[i : i + 4], 16) for i in range(0, len(hexes), 4)]


def scale_anchor(anchor, sx: float, sy: float) -> None:
    if anchor is None:
        return
    anchor.XCoordinate = round(anchor.XCoordinate * sx)
    anchor.YCoordinate = round(anchor.YCoordinate * sy)
    if anchor.Format == 2:
        anchor.Format = 1


def strip_kern_and_hints(font: TTFont) -> None:
    # the renderer kerns only under w:kern and Word's advances are already baked in
    features = {fr.FeatureTag for fr in font["GSUB"].table.FeatureList.FeatureRecord}
    features |= {fr.FeatureTag for fr in font["GPOS"].table.FeatureList.FeatureRecord}
    opt = subset.Options()
    opt.layout_features = sorted(features - {"kern"})
    opt.hinting = False
    opt.glyph_names = True
    opt.notdef_outline = True
    opt.name_IDs = ["*"]
    opt.name_legacy = True
    opt.prune_unicode_ranges = False
    sub = subset.Subsetter(opt)
    sub.populate(glyphs=font.getGlyphOrder())
    sub.subset(font)


def scale_value(value, sx: float, sy: float) -> None:
    for attr, k in (("XPlacement", sx), ("XAdvance", sx), ("YPlacement", sy), ("YAdvance", sy)):
        if hasattr(value, attr):
            setattr(value, attr, round(getattr(value, attr) * k))


def scale_gpos(font: TTFont, sx_of, sy: float) -> None:
    for lookup in font["GPOS"].table.LookupList.Lookup:
        for st in lookup.SubTable:
            if lookup.LookupType == 9:
                st = st.ExtSubTable
                ltype = st.LookupType
            else:
                ltype = lookup.LookupType
            if ltype == 4:
                for g, rec in zip(st.MarkCoverage.glyphs, st.MarkArray.MarkRecord):
                    scale_anchor(rec.MarkAnchor, sx_of(g), sy)
                for g, rec in zip(st.BaseCoverage.glyphs, st.BaseArray.BaseRecord):
                    for a in rec.BaseAnchor:
                        scale_anchor(a, sx_of(g), sy)
            elif ltype == 6:
                for g, rec in zip(st.Mark1Coverage.glyphs, st.Mark1Array.MarkRecord):
                    scale_anchor(rec.MarkAnchor, sx_of(g), sy)
                for g, rec in zip(st.Mark2Coverage.glyphs, st.Mark2Array.Mark2Record):
                    for a in rec.Mark2Anchor:
                        scale_anchor(a, sx_of(g), sy)
            elif ltype == 1 and st.Format == 1:
                # one shared record: scale it once, by the covered glyphs' median factor
                scale_value(st.Value, statistics.median(sx_of(g) for g in st.Coverage.glyphs), sy)
            elif ltype == 1:
                for g, v in zip(st.Coverage.glyphs, st.Value):
                    scale_value(v, sx_of(g), sy)
            elif ltype == 2:
                raise SystemExit("unexpected GPOS pair adjustment; extend scale_gpos")


def main() -> None:
    if len(sys.argv) < 4:
        sys.exit(__doc__)
    src, word, family = sys.argv[1], sys.argv[2], sys.argv[3]
    style = sys.argv[4] if len(sys.argv) > 4 else "Regular"
    if style not in ("Regular", "Bold"):
        sys.exit(__doc__)
    out = (
        Path(sys.argv[5])
        if len(sys.argv) > 5
        else OUT_DIR / (family.replace(" ", "") + "-" + style + ".woff2")
    )

    noto = TTFont(src)
    strip_kern_and_hints(noto)
    ref = open_face(word)
    upm = noto["head"].unitsPerEm
    ref_upm = ref["head"].unitsPerEm
    ref_cmap = ref.getBestCmap()
    ref_hmtx = ref["hmtx"]
    cmap = noto.getBestCmap()
    hmtx = noto["hmtx"]
    glyph_order = noto.getGlyphOrder()

    def ref_adv(cp: int) -> float:
        return ref_hmtx[ref_cmap[cp]][0] / ref_upm * upm

    shared = [cp for cp in THAI_BLOCK if cp in cmap and cp in ref_cmap]
    ratios = {cp: ref_adv(cp) / hmtx[cmap[cp]][0] for cp in shared if hmtx[cmap[cp]][0] > 0}
    sx_median = statistics.median(ratios[cp] for cp in ratios if cp in CONSONANTS)
    sy = statistics.median(
        (ink_height(ref, ref_cmap[cp]) / ref_upm) / (ink_height(noto, cmap[cp]) / upm)
        for cp in shared
        if cp in CONSONANTS and ink_height(noto, cmap[cp])
    )

    # target advance and horizontal factor per glyph: code point glyphs from
    # the Word hmtx, variants (.narrow/.short/...) and ligatures from the code
    # points in their name, everything else (marks, .notdef) the median factor
    cp_of_glyph = {g: cp for cp, g in cmap.items()}
    targets: dict[str, tuple[int | None, float]] = {}
    for g in glyph_order:
        cps = [cp_of_glyph[g]] if g in cp_of_glyph else name_codepoints(g)
        if cps and all(cp in ref_cmap for cp in cps):
            adv = sum(ref_adv(cp) for cp in cps)
            base = sum(hmtx[cmap[cp]][0] for cp in cps if cp in cmap)
            targets[g] = (round(adv), adv / base if base > 0 else sx_median)
        else:
            targets[g] = (None, sx_median)

    recorded = {}
    glyph_set = noto.getGlyphSet()
    for g in glyph_order:
        pen = DecomposingRecordingPen(glyph_set)
        glyph_set[g].draw(pen)
        recorded[g] = pen
    glyf = noto["glyf"]
    for g in glyph_order:
        adv, sx = targets[g]
        tt = TTGlyphPen(None)
        recorded[g].replay(TransformPen(tt, (sx, 0, 0, sy, 0, 0)))
        glyph = tt.glyph()
        glyph.recalcBounds(glyf)
        glyf[g] = glyph
        old_adv, _ = hmtx[g]
        new_adv = adv if adv is not None else round(old_adv * sx)
        hmtx[g] = (new_adv, getattr(glyph, "xMin", 0))

    scale_gpos(noto, lambda g: targets[g][1], sy)

    hhea, os2 = noto["hhea"], noto["OS/2"]
    r_hhea, r_os2 = ref["hhea"], ref["OS/2"]
    to = lambda v: round(v / ref_upm * upm)  # noqa: E731
    hhea.ascent, hhea.descent, hhea.lineGap = to(r_hhea.ascent), to(r_hhea.descent), to(r_hhea.lineGap)
    os2.usWinAscent, os2.usWinDescent = to(r_os2.usWinAscent), to(r_os2.usWinDescent)
    os2.sTypoAscender, os2.sTypoDescender, os2.sTypoLineGap = (
        to(r_os2.sTypoAscender),
        to(r_os2.sTypoDescender),
        to(r_os2.sTypoLineGap),
    )
    os2.fsSelection &= ~(1 << 7)  # USE_TYPO_METRICS off: Chromium reads hhea
    os2.recalcAvgCharWidth(noto)
    bold = style == "Bold"
    os2.usWeightClass = 700 if bold else 400
    os2.fsSelection = (os2.fsSelection & ~((1 << 5) | (1 << 6))) | ((1 << 5) if bold else (1 << 6))
    noto["head"].macStyle = (noto["head"].macStyle & ~1) | (1 if bold else 0)
    noto["head"].flags &= ~(1 << 3)  # no ppem rounding without hinting

    rename(noto, family, style)
    noto.flavor = "woff2"
    noto.flavorData = WOFF2FlavorData(transformedTables=())
    out.parent.mkdir(parents=True, exist_ok=True)
    noto.save(str(out))
    print(
        f"{out}: {len(glyph_order)} glyphs, x per glyph (median {sx_median:.3f}), "
        f"y x{sy:.3f}, {len(shared)} Thai advances from {word}"
    )


if __name__ == "__main__":
    main()
