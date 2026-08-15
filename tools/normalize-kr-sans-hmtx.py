#!/usr/bin/env python3
"""Normalize GenOffice Sans KR Basic Latin advances to Malgun Gothic.

Word substitutes missing Korean sans families with Malgun Gothic; the bundled
subset (Noto Sans CJK KR) ships hangul already normalized to 1.0em but keeps
Noto's Latin metrics (space 0.224em vs Malgun 0.352em, digits 0.555 vs 0.551),
so line breaks drift from Word. Rewrites hmtx advances for U+0020-007E and
U+00A0 to Malgun's values (measured from malgun.ttf 6.68, 2048upm, scaled to
1000upm); outlines are untouched. Idempotent.

Usage: python3 tools/normalize-kr-sans-hmtx.py [woff2-path]
"""

import sys
from pathlib import Path

from fontTools.ttLib import TTFont

DEFAULT = "apps/docs/src/renderer/fonts/GenOfficeSansKR-Regular-subset.woff2"

# ord -> advance in 1/1000 em, from Malgun Gothic Regular
MALGUN_ADVANCES = {
    0x20: 352, 0x21: 289, 0x22: 395, 0x23: 606, 0x24: 551, 0x25: 836,
    0x26: 818, 0x27: 232, 0x28: 305, 0x29: 305, 0x2A: 425, 0x2B: 701,
    0x2C: 219, 0x2D: 410, 0x2E: 219, 0x2F: 396, 0x30: 551, 0x31: 551,
    0x32: 551, 0x33: 551, 0x34: 551, 0x35: 551, 0x36: 551, 0x37: 551,
    0x38: 551, 0x39: 551, 0x3A: 219, 0x3B: 219, 0x3C: 701, 0x3D: 701,
    0x3E: 701, 0x3F: 460, 0x40: 979, 0x41: 658, 0x42: 583, 0x43: 635,
    0x44: 717, 0x45: 517, 0x46: 499, 0x47: 702, 0x48: 725, 0x49: 270,
    0x4A: 360, 0x4B: 590, 0x4C: 480, 0x4D: 917, 0x4E: 765, 0x4F: 773,
    0x50: 571, 0x51: 773, 0x52: 610, 0x53: 543, 0x54: 534, 0x55: 703,
    0x56: 634, 0x57: 954, 0x58: 601, 0x59: 563, 0x5A: 583, 0x5B: 305,
    0x5C: 764, 0x5D: 305, 0x5E: 701, 0x5F: 426, 0x60: 272, 0x61: 520,
    0x62: 601, 0x63: 473, 0x64: 602, 0x65: 535, 0x66: 316, 0x67: 602,
    0x68: 579, 0x69: 246, 0x6A: 246, 0x6B: 506, 0x6C: 246, 0x6D: 880,
    0x6E: 578, 0x6F: 599, 0x70: 601, 0x71: 602, 0x72: 354, 0x73: 433,
    0x74: 345, 0x75: 578, 0x76: 487, 0x77: 736, 0x78: 465, 0x79: 493,
    0x7A: 462, 0x7B: 305, 0x7C: 239, 0x7D: 305, 0x7E: 701, 0xA0: 352,
}


def main() -> None:
    root = Path(__file__).resolve().parent.parent
    path = Path(sys.argv[1]) if len(sys.argv) > 1 else root / DEFAULT
    font = TTFont(str(path))
    upm = font["head"].unitsPerEm
    cmap = font.getBestCmap()
    hmtx = font["hmtx"]
    changed = 0
    for cp, adv in MALGUN_ADVANCES.items():
        name = cmap.get(cp)
        if name is None:
            continue
        target = round(adv * upm / 1000)
        old, lsb = hmtx[name]
        if old != target:
            hmtx[name] = (target, lsb)
            changed += 1
    ga = cmap.get(0xAC00)
    assert ga and hmtx[ga][0] == upm, "hangul must stay 1.0em"
    if changed:
        font.flavor = "woff2"
        font.save(str(path))
    print(f"{path.name}: {changed} advances updated")


if __name__ == "__main__":
    main()
