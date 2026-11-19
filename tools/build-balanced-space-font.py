#!/usr/bin/env python3
"""Build GenOffice Balanced Space: a single blank 0.5em U+0020 glyph.

Under balanceSingleByteDoubleByteWidth Word advances a space beside an East
Asian character or another space by half the run size whatever face draws
it (line-metrics.ts balancedSpaceOffsets). The renderer wraps such spaces in
.doc-balanced-space and draws them from this face; its 0.8/0.2 line box sits
inside every CJK face's, so the span never lifts the line.

Usage: python3 tools/build-balanced-space-font.py [out.woff2]
"""

import sys

from fontTools.fontBuilder import FontBuilder
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.ttLib.woff2 import WOFF2FlavorData

DEFAULT_OUT = "apps/docs/src/renderer/fonts/GenOfficeBalancedSpace.woff2"
FAMILY = "GenOffice Balanced Space"
UPM = 1000


def main() -> None:
    out = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_OUT
    fb = FontBuilder(UPM, isTTF=True)
    fb.setupGlyphOrder([".notdef", "space"])
    empty = TTGlyphPen(None).glyph()
    fb.setupGlyf({".notdef": empty, "space": empty})
    fb.setupCharacterMap({0x20: "space"})
    fb.setupHorizontalMetrics({".notdef": (UPM // 2, 0), "space": (UPM // 2, 0)})
    fb.setupHorizontalHeader(ascent=800, descent=-200)
    fb.setupOS2(sTypoAscender=800, sTypoDescender=-200, usWinAscent=800, usWinDescent=200)
    fb.setupNameTable({"familyName": FAMILY, "styleName": "Regular"})
    fb.setupPost()
    fb.font.flavor = "woff2"
    fb.font.flavorData = WOFF2FlavorData(transformedTables=())
    fb.save(out)
    print(f"wrote {out}")


if __name__ == "__main__":
    main()
