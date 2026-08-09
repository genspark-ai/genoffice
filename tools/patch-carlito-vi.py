#!/usr/bin/env python3
"""Build-time patch for bundled Carlito (1.103 Beta1) Vietnamese stacked-diacritic
defect: some precomposed dot-below+above-mark glyphs (e.g. U+1EAC) reuse the
dot-below-only composition, dropping the circumflex/breve. Rebuilds them as
<above-mark glyph components> + <dot-below component>. Advance widths untouched.

Also renames the family to "Carlito GO" (OFL Reserved Font Name rule, since the
glyph data is modified). Idempotent: files are rewritten only when changed.

Usage:
  patch-carlito-vi.py            patch fonts in place
  patch-carlito-vi.py --verify   scan only; exit 1 if any broken glyph remains
"""
import copy
import sys
import unicodedata
from pathlib import Path

from fontTools.ttLib import TTFont

FONT_DIR = Path(__file__).resolve().parent.parent / "apps/docs/src/renderer/fonts"
FACES = ["Regular", "Bold", "Italic", "BoldItalic"]
DOT_BELOW = 0x0323
ABOVE_MARKS = {0x0302, 0x0306}  # circumflex, breve


def targets(cmap):
    """(cp, dot-only counterpart cp, above-mark-only counterpart cp) triples."""
    out = []
    for cp in cmap:
        nfd = unicodedata.normalize("NFD", chr(cp))
        marks = [ord(c) for c in nfd[1:]]
        above = [m for m in marks if m in ABOVE_MARKS]
        if DOT_BELOW not in marks or not above:
            continue
        base = nfd[0]
        dot_cp = ord(unicodedata.normalize("NFC", base + chr(DOT_BELOW)))
        above_cp = ord(unicodedata.normalize("NFC", base + chr(above[0])))
        if dot_cp in cmap and above_cp in cmap:
            out.append((cp, dot_cp, above_cp))
    return sorted(out)


def comp_set(glyph):
    return sorted((c.glyphName, c.x, c.y) for c in glyph.components)


def contour_shapes(glyf, gn):
    """Translation-invariant contour shapes (composites are expanded)."""
    coords, ends, flags = glyf[gn].getCoordinates(glyf)
    cl, fl, shapes, start = list(coords), list(flags), [], 0
    for e in ends:
        pts = cl[start : e + 1]
        x0 = min(p[0] for p in pts)
        y0 = min(p[1] for p in pts)
        shapes.append(tuple(((x - x0, y - y0), f) for (x, y), f in zip(pts, fl[start : e + 1])))
        start = e + 1
    return shapes


def find_broken(font):
    """[(cp, glyphname, fixable)] — fixable means the composite rebuild applies."""
    glyf, cmap = font["glyf"], font.getBestCmap()
    broken = []
    for cp, dot_cp, above_cp in targets(cmap):
        gn, gd, ga = cmap[cp], cmap[dot_cp], cmap[above_cp]
        g = glyf[gn]
        if g.isComposite() and glyf[gd].isComposite():
            if comp_set(g) == comp_set(glyf[gd]):
                broken.append((cp, gn, glyf[ga].isComposite()))
            continue
        # decomposed outlines: both mark shapes must be present
        base_cp = ord(unicodedata.normalize("NFD", chr(cp))[0])
        base = set(contour_shapes(glyf, cmap[base_cp])) if base_cp in cmap else set()
        mine = set(contour_shapes(glyf, gn))
        above_marks = set(contour_shapes(glyf, ga)) - base
        dot_marks = set(contour_shapes(glyf, gd)) - base
        if (above_marks and not above_marks & mine) or (dot_marks and not dot_marks & mine):
            broken.append((cp, gn, False))
    return broken


def fix(font, broken):
    glyf, cmap, hmtx = font["glyf"], font.getBestCmap(), font["hmtx"]
    fixed = []
    for cp, gn, fixable in broken:
        if not fixable:
            raise SystemExit(f"U+{cp:04X} {gn}: broken but not composite-fixable, bailing")
        nfd = unicodedata.normalize("NFD", chr(cp))
        above = next(m for m in nfd[1:] if ord(m) in ABOVE_MARKS)
        dot_gn = cmap[ord(unicodedata.normalize("NFC", nfd[0] + chr(DOT_BELOW)))]
        above_gn = cmap[ord(unicodedata.normalize("NFC", nfd[0] + above))]
        new = copy.deepcopy(glyf[above_gn])
        above_names = {c.glyphName for c in new.components}
        dot_comps = [c for c in glyf[dot_gn].components if c.glyphName not in above_names]
        assert dot_comps, f"U+{cp:04X}: no dot-below component found in {dot_gn}"
        new.components.extend(copy.deepcopy(c) for c in dot_comps)
        new.recalcBounds(glyf)
        glyf[gn] = new
        adv, _lsb = hmtx[gn]
        hmtx[gn] = (adv, new.xMin)
        fixed.append((cp, gn))
    return fixed


def rename(font):
    changed = False
    for rec in font["name"].names:
        s = rec.toUnicode()
        if "Carlito" not in s or "Carlito GO" in s or "CarlitoGO" in s:
            continue
        if rec.nameID == 6:  # PostScript name: no spaces
            rec.string = s.replace("Carlito", "CarlitoGO")
            changed = True
        elif rec.nameID in (1, 3, 4, 16, 17):
            rec.string = s.replace("Carlito", "Carlito GO")
            changed = True
    return changed


def main():
    verify = "--verify" in sys.argv
    bad = False
    for face in FACES:
        path = FONT_DIR / f"Carlito-{face}.ttf"
        font = TTFont(path, recalcTimestamp=False)
        broken = find_broken(font)
        if verify:
            names = [f"U+{cp:04X} {chr(cp)} ({gn})" for cp, gn, _ in broken]
            print(f"{face}: {'OK, no broken glyphs' if not names else 'BROKEN ' + ', '.join(names)}")
            bad |= bool(names)
            continue
        adv_before = {gn: font["hmtx"][gn][0] for gn in font.getGlyphOrder()}
        fixed = fix(font, broken)
        renamed = rename(font)
        if not fixed and not renamed:
            print(f"{face}: already patched, unchanged")
            continue
        assert not find_broken(font)
        font.save(path)
        reloaded = TTFont(path)
        adv_after = {gn: reloaded["hmtx"][gn][0] for gn in reloaded.getGlyphOrder()}
        assert adv_after == adv_before, f"{face}: advance widths changed!"
        print(
            f"{face}: fixed {[f'U+{cp:04X} {chr(cp)}' for cp, _ in fixed]}, "
            f"renamed={renamed}, advances unchanged ({len(adv_before)} glyphs)"
        )
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
