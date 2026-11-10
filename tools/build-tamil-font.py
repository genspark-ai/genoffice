#!/usr/bin/env python3
"""Build GenOffice Tamil from upstream Noto Sans Tamil, hmtx-normalized to Latha.

Word substitutes missing Tamil families (Latha/Vijaya/Noto Sans Tamil...) with
Latha; macOS has no Latha and Chromium's Tamil fallback (Tamil Sangam MN) is
~27% narrower shaped (M3 probe 2026-08-14: sentence R 0.728, space 0.39x), so
line breaks drift far from Word. This rewrites advances to Latha's values:
cmap-shared codepoints exactly, remaining glyphs (GSUB conjunct/matra outputs)
by the median Tamil-letter ratio. Outlines are untouched. Validated shaped
sentence R vs Latha: 0.994 mean, all sentences within +/-2.3%.

Renamed per OFL 1.1 ("Noto" is a Reserved Font Name; advances are modified).
Upstream: https://github.com/notofonts/notofonts.github.io
          fonts/NotoSansTamil/hinted/ttf/NotoSansTamil-Regular.ttf

Usage: python3 tools/build-tamil-font.py <NotoSansTamil-Regular.ttf> [out.woff2]
"""

import statistics
import sys
from pathlib import Path

from fontTools.ttLib import TTFont
from fontTools.ttLib.woff2 import WOFF2FlavorData

DEFAULT_OUT = "apps/docs/src/renderer/fonts/GenOfficeTamil-Regular.woff2"
FAMILY = "GenOffice Tamil"

# ord -> advance in 1/1000 em, measured from Word's latha.ttf (2048 upm).
# Latha has no A-Z/a-z overlap with Noto Sans Tamil (Noto ships no Latin
# letters), so the shared set is Tamil block + digits + punctuation + space.
LATHA_ADVANCES = {
    0x0020: 578, 0x0021: 250, 0x0022: 319, 0x0023: 500, 0x0025: 800, 0x0027: 172,
    0x0028: 300, 0x0029: 300, 0x002A: 350, 0x002B: 525, 0x002C: 250, 0x002D: 300,
    0x002E: 250, 0x002F: 250, 0x0030: 500, 0x0031: 500, 0x0032: 500, 0x0033: 500,
    0x0034: 500, 0x0035: 500, 0x0036: 500, 0x0037: 500, 0x0038: 500, 0x0039: 500,
    0x003A: 250, 0x003B: 250, 0x003C: 525, 0x003D: 525, 0x003E: 525, 0x003F: 500,
    0x005B: 250, 0x005C: 250, 0x005D: 250, 0x005E: 422, 0x005F: 497, 0x007B: 301,
    0x007C: 234, 0x007D: 301, 0x007E: 525, 0x0B82: 682, 0x0B83: 682, 0x0B85: 1133,
    0x0B86: 1352, 0x0B87: 1086, 0x0B88: 840, 0x0B89: 1070, 0x0B8A: 1363, 0x0B8E: 824,
    0x0B8F: 824, 0x0B90: 965, 0x0B92: 1098, 0x0B93: 1098, 0x0B94: 2109, 0x0B95: 797,
    0x0B99: 1035, 0x0B9A: 742, 0x0B9C: 973, 0x0B9E: 1234, 0x0B9F: 930, 0x0BA3: 1535,
    0x0BA4: 836, 0x0BA8: 746, 0x0BA9: 1184, 0x0BAA: 738, 0x0BAE: 965, 0x0BAF: 1016,
    0x0BB0: 621, 0x0BB1: 910, 0x0BB2: 1105, 0x0BB3: 1051, 0x0BB4: 960, 0x0BB5: 1117,
    0x0BB6: 1169, 0x0BB7: 1414, 0x0BB8: 1562, 0x0BB9: 1775, 0x0BBE: 617, 0x0BBF: 188,
    0x0BC1: 496, 0x0BC2: 832, 0x0BC6: 828, 0x0BC7: 785, 0x0BC8: 1202, 0x0BCA: 1479,
    0x0BCB: 1458, 0x0BCC: 1906, 0x0BD0: 1117, 0x0BD7: 1051, 0x0BE6: 581, 0x0BE7: 797,
    0x0BE8: 930, 0x0BE9: 918, 0x0BEA: 852, 0x0BEB: 1152, 0x0BEC: 1285, 0x0BED: 824,
    0x0BEE: 1133, 0x0BEF: 1258, 0x0BF0: 973, 0x0BF1: 926, 0x0BF2: 1102, 0x0BF3: 989,
    0x0BF4: 1348, 0x0BF5: 2208, 0x0BF6: 882, 0x0BF7: 1372, 0x0BF8: 1990, 0x0BF9: 1125,
    0x0BFA: 1032,
}

TAMIL_BLOCK = range(0x0B80, 0x0C00)


def rename(font: TTFont, family: str, ps_name: str) -> None:
    name = font["name"]
    for rec in list(name.names):
        if rec.nameID in (1, 16):
            name.setName(family, rec.nameID, rec.platformID, rec.platEncID, rec.langID)
        elif rec.nameID in (3, 4):
            name.setName(f"{family} Regular", rec.nameID, rec.platformID, rec.platEncID, rec.langID)
        elif rec.nameID == 6:
            name.setName(ps_name, 6, rec.platformID, rec.platEncID, rec.langID)


def main() -> None:
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    src = Path(sys.argv[1])
    root = Path(__file__).resolve().parent.parent
    out = Path(sys.argv[2]) if len(sys.argv) > 2 else root / DEFAULT_OUT

    font = TTFont(str(src))
    upm = font["head"].unitsPerEm
    cmap = font.getBestCmap()
    hmtx = font["hmtx"]

    exact = {}
    ratios = []
    for cp, adv in LATHA_ADVANCES.items():
        gname = cmap.get(cp)
        if gname is None:
            continue
        exact[gname] = round(adv * upm / 1000)
        old = hmtx[gname][0]
        if old > 0 and adv > 0 and cp in TAMIL_BLOCK:
            ratios.append(adv * upm / 1000 / old)
    scale = statistics.median(ratios)

    for gname in font.getGlyphOrder():
        adv, lsb = hmtx[gname]
        if gname in exact:
            hmtx[gname] = (exact[gname], lsb)
        elif adv > 0:
            hmtx[gname] = (round(adv * scale), lsb)

    rename(font, FAMILY, "GenOfficeTamil-Regular")
    space = cmap.get(0x20)
    assert space and hmtx[space][0] == round(0.578 * upm), "space must match Latha"
    font.flavor = "woff2"
    # plain glyf/loca (no woff2 transform), same as the other bundled subsets —
    # tests/helpers/woff2-metrics.ts reads the tables directly
    font.flavorData = WOFF2FlavorData(transformedTables=())
    font.save(str(out))
    print(f"{out.name}: {len(exact)} exact, scale {scale:.4f} for the rest")


if __name__ == "__main__":
    main()
