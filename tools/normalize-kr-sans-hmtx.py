#!/usr/bin/env python3
"""Normalize GenOffice Sans KR Basic Latin advances to Malgun Gothic.

Word substitutes missing Korean sans families with Malgun Gothic; the bundled
subset (Noto Sans CJK KR) ships hangul already normalized to 1.0em but keeps
Noto's Latin metrics (space 0.224em vs Malgun 0.352em, digits 0.555 vs 0.551),
so line breaks drift from Word. Rewrites hmtx advances for U+0020-007E,
U+00A0-00FF, the General Punctuation glyphs Malgun carries and the
proportional CJK brackets to Malgun's values (measured from
malgun.ttf 6.68, 2048upm, scaled to 1000upm); outlines are untouched
(scale-kr-sans-latin-ink.py moves them into the new advances). Idempotent.

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
    # Latin-1 Supplement and General Punctuation: Malgun keeps them
    # proportional (middle dot 0.219em, dashes, quotes, bullet, ellipsis);
    # Noto's 0.56em middle dot alone wrapped Korean list lines early
    0xA1: 289, 0xA2: 551, 0xA3: 551, 0xA4: 570, 0xA5: 551, 0xA6: 239,
    0xA7: 458, 0xA8: 422, 0xA9: 914, 0xAA: 401, 0xAB: 514, 0xAC: 701,
    0xAD: 410, 0xAE: 914, 0xAF: 440, 0xB0: 387, 0xB1: 701, 0xB2: 373,
    0xB3: 373, 0xB4: 288, 0xB5: 590, 0xB6: 467, 0xB7: 219, 0xB8: 210,
    0xB9: 357, 0xBA: 440, 0xBB: 514, 0xBC: 916, 0xBD: 952, 0xBE: 950,
    0xBF: 460, 0xC0: 658, 0xC1: 658, 0xC2: 658, 0xC3: 658, 0xC4: 658,
    0xC5: 658, 0xC6: 877, 0xC7: 635, 0xC8: 517, 0xC9: 517, 0xCA: 517,
    0xCB: 517, 0xCC: 270, 0xCD: 270, 0xCE: 270, 0xCF: 270, 0xD0: 717,
    0xD1: 765, 0xD2: 773, 0xD3: 773, 0xD4: 773, 0xD5: 773, 0xD6: 773,
    0xD7: 701, 0xD8: 773, 0xD9: 703, 0xDA: 703, 0xDB: 703, 0xDC: 703,
    0xDD: 563, 0xDE: 571, 0xDF: 552, 0xE0: 520, 0xE1: 520, 0xE2: 520,
    0xE3: 520, 0xE4: 520, 0xE5: 520, 0xE6: 854, 0xE7: 473, 0xE8: 535,
    0xE9: 535, 0xEA: 535, 0xEB: 535, 0xEC: 246, 0xED: 246, 0xEE: 246,
    0xEF: 246, 0xF0: 571, 0xF1: 578, 0xF2: 599, 0xF3: 599, 0xF4: 599,
    0xF5: 599, 0xF6: 599, 0xF7: 701, 0xF8: 599, 0xF9: 578, 0xFA: 578,
    0xFB: 578, 0xFC: 578, 0xFD: 493, 0xFE: 601, 0xFF: 493, 0x2013: 513,
    0x2014: 1026, 0x2015: 1000, 0x2018: 231, 0x2019: 231, 0x201A: 231,
    0x201C: 379, 0x201D: 379, 0x201E: 379, 0x2020: 382, 0x2021: 382,
    0x2022: 417, 0x2025: 437, 0x2026: 740, 0x2030: 1238, 0x2032: 231,
    0x2033: 379, 0x2039: 322, 0x203A: 322, 0x203C: 578,
    0x2047: 920, 0x2048: 749, 0x2049: 749,
    # CJK brackets and the reference mark are proportional in Malgun (Noto
    # keeps them full-width): a law title in corner brackets wraps a line early
    0x300C: 571, 0x300D: 571, 0x300E: 517, 0x300F: 517, 0x3008: 542,
    0x3009: 542, 0x300A: 596, 0x300B: 596, 0x3010: 518, 0x3011: 518,
    0x3014: 488, 0x3015: 488, 0x203B: 800,
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
