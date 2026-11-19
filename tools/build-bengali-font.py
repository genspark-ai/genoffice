#!/usr/bin/env python3
"""Build GenOffice Bengali from upstream Noto Sans Bengali, hmtx-normalized to Nirmala UI.

Word for Mac renders Nirmala UI (the Windows Bengali UI face, an M365 cloud
font) for real; macOS has no Nirmala UI and Chromium's Bengali fallback
(Kohinoor Bangla) shapes ~13% narrower, so a two-page Bengali notice fits one
page. This rewrites advances to Nirmala UI's values: cmap-shared codepoints
exactly (86% of the shaped width of a Bengali legal notice), remaining glyphs
(GSUB conjunct/matra outputs) by one scale fitted so the corpus' shaped width
matches Nirmala UI (hb-shape, 26 paragraphs: body paragraphs land within
0.98-1.02, every paragraph within 0.93-1.07; the median letter ratio of 1.10
over-widened the conjuncts to 1.10x). Outlines are untouched.

Renamed per OFL 1.1 ("Noto" is a Reserved Font Name; advances are modified).
Upstream: https://github.com/notofonts/notofonts.github.io
          fonts/NotoSansBengali/hinted/ttf/NotoSansBengali-{Regular,Bold}.ttf

Usage: python3 tools/build-bengali-font.py <NotoSansBengali-Regular.ttf> <NotoSansBengali-Bold.ttf> [out-dir]
"""

import sys
from pathlib import Path

from fontTools.ttLib import TTFont
from fontTools.ttLib.woff2 import WOFF2FlavorData

DEFAULT_DIR = "apps/docs/src/renderer/fonts"
FAMILY = "GenOffice Bengali"

# ord -> advance in 1/1000 em, measured from Nirmala UI (2048 upm) for the
# codepoints Noto Sans Bengali also maps (no Latin letters or ASCII digits).
REGULAR = {
    0x0020: 274, 0x0021: 284, 0x003F: 448, 0x005B: 302, 0x005D: 302, 0x007B: 302,
    0x007D: 302, 0x00A0: 274, 0x02BC: 229, 0x0951: 0, 0x0952: 0, 0x0964: 432,
    0x0965: 634, 0x0980: 637, 0x0981: 0, 0x0982: 466, 0x0983: 466, 0x0985: 986,
    0x0986: 1208, 0x0987: 593, 0x0988: 842, 0x0989: 784, 0x098A: 840, 0x098B: 849,
    0x098C: 710, 0x098F: 748, 0x0990: 958, 0x0993: 702, 0x0994: 848, 0x0995: 878,
    0x0996: 690, 0x0997: 718, 0x0998: 658, 0x0999: 779, 0x099A: 604, 0x099B: 680,
    0x099C: 958, 0x099D: 803, 0x099E: 1096, 0x099F: 623, 0x09A0: 637, 0x09A1: 798,
    0x09A2: 623, 0x09A3: 633, 0x09A4: 771, 0x09A5: 654, 0x09A6: 644, 0x09A7: 589,
    0x09A8: 689, 0x09AA: 805, 0x09AB: 911, 0x09AC: 608, 0x09AD: 796, 0x09AE: 727,
    0x09AF: 642, 0x09B0: 608, 0x09B2: 721, 0x09B6: 733, 0x09B7: 643, 0x09B8: 727,
    0x09B9: 629, 0x09BC: 0, 0x09BD: 538, 0x09BE: 251, 0x09BF: 250, 0x09C0: 250,
    0x09C1: 0, 0x09C2: 0, 0x09C3: 0, 0x09C4: 0, 0x09C7: 411, 0x09C8: 411, 0x09CB: 1206,
    0x09CC: 1205, 0x09CD: 0, 0x09CE: 611, 0x09D7: 251, 0x09DC: 799, 0x09DD: 623,
    0x09DF: 642, 0x09E0: 849, 0x09E1: 953, 0x09E2: 0, 0x09E3: 0, 0x09E6: 651,
    0x09E7: 515, 0x09E8: 668, 0x09E9: 756, 0x09EA: 639, 0x09EB: 747, 0x09EC: 829,
    0x09ED: 637, 0x09EE: 764, 0x09EF: 688, 0x09F0: 608, 0x09F1: 608, 0x09F2: 665,
    0x09F3: 658, 0x09F4: 734, 0x09F5: 729, 0x09F6: 786, 0x09F7: 227, 0x09F8: 679,
    0x09F9: 505, 0x09FA: 764, 0x09FB: 619, 0x09FC: 590, 0x09FD: 371, 0x09FE: 0,
    0x1CD0: 0, 0x1CD2: 0, 0x1CD5: 0, 0x1CD6: 0, 0x1CD8: 0, 0x1CE1: 367, 0x1CEA: 1004,
    0x1CED: 0, 0x1CF2: 611, 0x1CF5: 602, 0x1CF6: 718, 0x1CF7: 398, 0x200B: 0,
    0x200C: 0, 0x200D: 0, 0x2010: 400, 0x2018: 229, 0x2019: 229, 0x201C: 377,
    0x201D: 377, 0x20B9: 539, 0x25CC: 594, 0xA8F1: 0,
}
BOLD = {
    0x0020: 276, 0x0021: 327, 0x003F: 438, 0x005B: 369, 0x005D: 369, 0x007B: 369,
    0x007D: 369, 0x00A0: 276, 0x02BC: 290, 0x0951: 0, 0x0952: 0, 0x0964: 481,
    0x0965: 724, 0x0980: 647, 0x0981: 0, 0x0982: 466, 0x0983: 466, 0x0985: 1099,
    0x0986: 1349, 0x0987: 689, 0x0988: 872, 0x0989: 835, 0x098A: 968, 0x098B: 986,
    0x098C: 746, 0x098F: 820, 0x0990: 1006, 0x0993: 775, 0x0994: 950, 0x0995: 957,
    0x0996: 762, 0x0997: 771, 0x0998: 729, 0x0999: 853, 0x099A: 659, 0x099B: 838,
    0x099C: 1101, 0x099D: 911, 0x099E: 1129, 0x099F: 693, 0x09A0: 636, 0x09A1: 852,
    0x09A2: 706, 0x09A3: 704, 0x09A4: 852, 0x09A5: 724, 0x09A6: 663, 0x09A7: 679,
    0x09A8: 702, 0x09AA: 866, 0x09AB: 1039, 0x09AC: 685, 0x09AD: 870, 0x09AE: 763,
    0x09AF: 761, 0x09B0: 685, 0x09B2: 787, 0x09B6: 827, 0x09B7: 722, 0x09B8: 795,
    0x09B9: 689, 0x09BC: 0, 0x09BD: 583, 0x09BE: 303, 0x09BF: 308, 0x09C0: 308,
    0x09C1: 0, 0x09C2: 0, 0x09C3: 0, 0x09C4: 0, 0x09C7: 473, 0x09C8: 473, 0x09CB: 1299,
    0x09CC: 1296, 0x09CD: 0, 0x09CE: 616, 0x09D7: 308, 0x09DC: 852, 0x09DD: 706,
    0x09DF: 761, 0x09E0: 986, 0x09E1: 1000, 0x09E2: 0, 0x09E3: 0, 0x09E6: 651,
    0x09E7: 557, 0x09E8: 691, 0x09E9: 817, 0x09EA: 649, 0x09EB: 747, 0x09EC: 815,
    0x09ED: 637, 0x09EE: 815, 0x09EF: 743, 0x09F0: 684, 0x09F1: 684, 0x09F2: 665,
    0x09F3: 809, 0x09F4: 734, 0x09F5: 740, 0x09F6: 823, 0x09F7: 274, 0x09F8: 679,
    0x09F9: 505, 0x09FA: 782, 0x09FB: 664, 0x09FC: 636, 0x09FD: 402, 0x09FE: 0,
    0x1CD0: 0, 0x1CD2: 0, 0x1CD5: 0, 0x1CD6: 0, 0x1CD8: 0, 0x1CE1: 426, 0x1CEA: 1073,
    0x1CED: 0, 0x1CF2: 637, 0x1CF5: 650, 0x1CF6: 823, 0x1CF7: 433, 0x200B: 0,
    0x200C: 0, 0x200D: 0, 0x2010: 404, 0x2018: 290, 0x2019: 290, 0x201C: 493,
    0x201D: 493, 0x20B9: 575, 0x25CC: 594, 0xA8F1: 0,
}

# fitted scale for glyphs outside the cmap (conjuncts, half forms)
REST_SCALE = {"Regular": 0.936, "Bold": 1.011}


def rename(font: TTFont, family: str, style: str) -> None:
    name = font["name"]
    ps = f"{family.replace(' ', '')}-{style}"
    for rec in list(name.names):
        if rec.nameID in (1, 16):
            name.setName(family, rec.nameID, rec.platformID, rec.platEncID, rec.langID)
        elif rec.nameID in (3, 4):
            name.setName(f"{family} {style}", rec.nameID, rec.platformID, rec.platEncID, rec.langID)
        elif rec.nameID == 6:
            name.setName(ps, 6, rec.platformID, rec.platEncID, rec.langID)


def build(src: Path, advances: dict, style: str, out: Path) -> None:
    font = TTFont(str(src))
    upm = font["head"].unitsPerEm
    cmap = font.getBestCmap()
    hmtx = font["hmtx"]

    exact = {}
    for cp, adv in advances.items():
        gname = cmap.get(cp)
        if gname is not None:
            exact[gname] = round(adv * upm / 1000)
    scale = REST_SCALE[style]

    for gname in font.getGlyphOrder():
        adv, lsb = hmtx[gname]
        if gname in exact:
            hmtx[gname] = (exact[gname], lsb)
        elif adv > 0:
            hmtx[gname] = (round(adv * scale), lsb)

    rename(font, FAMILY, style)
    space = cmap.get(0x20)
    assert space and hmtx[space][0] == round(advances[0x20] * upm / 1000), "space must match Nirmala UI"
    font.flavor = "woff2"
    # plain glyf/loca (no woff2 transform) like the other bundled subsets —
    # tests/helpers/woff2-metrics.ts reads the tables directly
    font.flavorData = WOFF2FlavorData(transformedTables=())
    font.save(str(out))
    print(f"{out.name}: {len(exact)} exact, scale {scale} for the rest")


def main() -> None:
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    root = Path(__file__).resolve().parent.parent
    out_dir = Path(sys.argv[3]) if len(sys.argv) > 3 else root / DEFAULT_DIR
    build(Path(sys.argv[1]), REGULAR, "Regular", out_dir / "GenOfficeBengali-Regular.woff2")
    build(Path(sys.argv[2]), BOLD, "Bold", out_dir / "GenOfficeBengali-Bold.woff2")


if __name__ == "__main__":
    main()
