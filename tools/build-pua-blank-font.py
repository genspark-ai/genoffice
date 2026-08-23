#!/usr/bin/env python3
"""Build GenOffice PUA Blank: every BMP Private Use codepoint maps to one
blank 1em glyph.

Chromium never system-falls-back for PUA characters: an unmapped PUA
codepoint renders the chain's primary font's .notdef. Chains headed by a
real installed face (Calibri, Carlito GO...) therefore draw tofu boxes for
AI-residue PUA tokens, while Word and chains headed by the bundled CJK
subsets (whose subsetted .notdef is blank) show nothing. This font gives
those chains an explicit blank glyph so PUA stays invisible everywhere,
with the same 1em advance the subset .notdef had.

Usage: python3 tools/build-pua-blank-font.py [out.woff2]
"""

import sys

from fontTools.fontBuilder import FontBuilder
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.ttLib.woff2 import WOFF2FlavorData

DEFAULT_OUT = "apps/docs/src/renderer/fonts/GenOfficePUABlank.woff2"
FAMILY = "GenOffice PUA Blank"
UPM = 1000


def main() -> None:
    out = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_OUT
    fb = FontBuilder(UPM, isTTF=True)
    fb.setupGlyphOrder([".notdef", "blank"])
    empty = TTGlyphPen(None).glyph()
    fb.setupGlyf({".notdef": empty, "blank": empty})
    fb.setupCharacterMap({cp: "blank" for cp in range(0xE000, 0xF900)})
    fb.setupHorizontalMetrics({".notdef": (UPM, 0), "blank": (UPM, 0)})
    fb.setupHorizontalHeader(ascent=800, descent=-200)
    fb.setupOS2(sTypoAscender=800, sTypoDescender=-200, usWinAscent=800, usWinDescent=200)
    fb.setupNameTable({"familyName": FAMILY, "styleName": "Regular"})
    fb.setupPost()
    fb.font.flavor = "woff2"
    # untransformed glyf, like the other bundled subsets (test helper contract)
    fb.font.flavorData = WOFF2FlavorData(transformedTables=())
    fb.save(out)
    print(f"wrote {out}")


if __name__ == "__main__":
    main()
