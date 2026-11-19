#!/usr/bin/env python3
"""Build advance-matched stand-in fonts for Office-only faces (manifest refit
method 'advance-matched').

Per-bucket size-adjust faces (hmtx-buckets / aptos-buckets) reproduce Word's
advances but scale every glyph's height with its width, so letters inside one
word sit at different heights ("SUSUNAN" with a 1.3x S next to a 1.04x U). This
pass instead derives one font file per cut from the open base font: the glyphs
shared with the Word font are subset out, each outline is scaled horizontally
to the Word advance (times the cut's kerning factor), the vertical axis gets
one uniform factor (mean cap-height / x-height ratio of the two fonts), and
hhea/OS2 take the line box of the family's general face. GPOS/GSUB/kern are
dropped (the kerning factor is baked into the advances; the renderer kerns
only under w:kern) and so is hinting. Only the transformed base outlines ship.

refit: { method: 'advance-matched', range, kern?: { cut: factor },
         cuts: { 'weight/style': { word, base, out } } }
word / base use the build-font-aliases.mjs locators (dfonts: / cloud: /
repo: / system:), out is the Assets path of the derived TTF. The manifest's rangeless
general faces stay (they carry the non-Latin fallback); the script rewrites
the ranged faces to one per cut and leaves fonts.generated.css to
`node tools/build-font-aliases.mjs`.

Usage: python3 tools/build-metric-fonts.py [--family NAME] [--dfonts DIR]
Needs Word's DFonts (or the files named by the locators) on this machine.
"""

import argparse
import json
import os
import re
import subprocess
import sys
from pathlib import Path

from fontTools import subset
from fontTools.pens.boundsPen import BoundsPen
from fontTools.pens.recordingPen import DecomposingRecordingPen
from fontTools.pens.transformPen import TransformPen
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.ttLib import TTFont

ROOT = Path(__file__).resolve().parent.parent
MANIFEST_DIR = ROOT / "packages/font-metrics/measurements/families"
DFONTS = "/Applications/Microsoft Word.app/Contents/Resources/DFonts"
CLOUD = os.path.expanduser("~/Library/Group Containers/UBF8T346G9.Office/FontCache/4/CloudFonts")
CUT_SUBFAMILY = {
    "normal/normal": "Regular",
    "bold/normal": "Bold",
    "normal/italic": "Italic",
    "bold/italic": "Bold Italic",
}
DROP_TABLES = ["GPOS", "GSUB", "GDEF", "kern", "DSIG", "FFTM", "LTSH", "VDMX", "hdmx", "meta", "STAT"]
os.environ.setdefault("SOURCE_DATE_EPOCH", "0")


def codepoints(spec):
    out = set()
    for part in spec.split(","):
        m = re.match(r"\s*U\+([0-9A-Fa-f]+)(?:-([0-9A-Fa-f]+))?", part)
        a = int(m.group(1), 16)
        b = int(m.group(2), 16) if m.group(2) else a
        out.update(range(a, b + 1))
    return out


def format_range(cps):
    runs = []
    for cp in sorted(cps):
        if runs and runs[-1][1] == cp - 1:
            runs[-1][1] = cp
        else:
            runs.append([cp, cp])
    return ", ".join(f"U+{a:04X}" if a == b else f"U+{a:04X}-{b:04X}" for a, b in runs)


def locate(spec, dfonts):
    kind, _, rest = spec.partition(":")
    path, _, index = rest.partition("#")
    if kind == "dfonts":
        return TTFont(os.path.join(dfonts, path), fontNumber=int(index or 0))
    if kind == "cloud":
        folder, _, full_name = path.partition("/")
        for file in sorted(os.listdir(os.path.join(CLOUD, folder))):
            font = TTFont(os.path.join(CLOUD, folder, file))
            if font["name"].getDebugName(4) == full_name:
                return font
        raise SystemExit(f"{full_name} not found under {folder}")
    if kind == "repo":
        return TTFont(ROOT / path, fontNumber=int(index or 0))
    if kind == "system":
        return TTFont(path, fontNumber=int(index or 0))
    raise SystemExit(f"unsupported locator {spec}")


def ink_height(font, ch):
    gs = font.getGlyphSet()
    name = font.getBestCmap().get(ord(ch))
    if not name:
        return None
    pen = BoundsPen(gs)
    gs[name].draw(pen)
    return pen.bounds[3] / font["head"].unitsPerEm if pen.bounds else None


def y_scale(word, base):
    ratios = []
    for ch in "Hx":
        w, b = ink_height(word, ch), ink_height(base, ch)
        if w and b:
            ratios.append(w / b)
    return round(sum(ratios) / len(ratios), 3) if ratios else 1.0


def general_box(family, cut):
    """em line box of the cut's rangeless face: overrides (or the base hhea) times size-adjust"""
    weight, style = cut.split("/")
    face = next(
        f
        for f in family["faces"]
        if not f.get("unicodeRange")
        and f.get("weight", "normal") == weight
        and f.get("style", "normal") == style
    )
    return face


def effective_box(face, base):
    scale = face.get("sizeAdjust", 100) / 100
    upm = base["head"].unitsPerEm
    hhea = base["hhea"]
    asc = face.get("ascentOverride", hhea.ascent / upm * 100) / 100 * scale
    desc = face.get("descentOverride", -hhea.descent / upm * 100) / 100 * scale
    gap = face.get("lineGapOverride", hhea.lineGap / upm * 100) / 100 * scale
    return asc, desc, gap


def set_names(font, family, subfamily, version_note):
    name = font["name"]
    keep = {0, 13, 14}
    name.names = [n for n in name.names if n.nameID in keep]
    bold = "Bold" in subfamily
    italic = "Italic" in subfamily
    ps = f"{family.replace(' ', '')}-{subfamily.replace(' ', '')}"
    for nid, text in (
        (1, family),
        (2, subfamily),
        (3, f"GenOffice;{ps}"),
        (4, f"{family} {subfamily}" if subfamily != "Regular" else family),
        (5, f"Version 1.000; {version_note}"),
        (6, ps),
    ):
        name.setName(text, nid, 3, 1, 0x409)
    os2 = font["OS/2"]
    os2.fsSelection = (os2.fsSelection & ~0b1100001) | (0b100000 if bold else 0) | (1 if italic else 0)
    if not bold and not italic:
        os2.fsSelection |= 0b1000000
    font["head"].macStyle = (1 if bold else 0) | (2 if italic else 0)


def build_cut(family, cut, spec, dfonts, out_dir):
    word = locate(spec["word"], dfonts)
    base = locate(spec["base"], dfonts)
    w_cmap, b_cmap = word.getBestCmap(), base.getBestCmap()
    cps = sorted(cp for cp in codepoints(family["refit"]["range"]) if cp in w_cmap and cp in b_cmap)
    # Carlito has no U+00A0; give it the space glyph rather than rely on shaper fallback
    nbsp_alias = 0xA0 in w_cmap and 0xA0 not in b_cmap and 0x20 in b_cmap
    kern = family["refit"].get("kern", {}).get(cut, 1.0)
    sy = y_scale(word, base)

    opts = subset.Options()
    opts.drop_tables += DROP_TABLES
    opts.layout_features = []
    opts.hinting = False
    opts.notdef_outline = True
    opts.name_IDs = ["*"]
    opts.name_legacy = True
    opts.glyph_names = False
    opts.recalc_bounds = True
    sub = subset.Subsetter(opts)
    sub.populate(unicodes=cps)
    sub.subset(base)

    upm, w_upm = base["head"].unitsPerEm, word["head"].unitsPerEm
    glyf, hmtx, gs = base["glyf"], base["hmtx"], base.getGlyphSet()
    cmap = base.getBestCmap()
    w_hmtx = word["hmtx"]
    # record every outline before rewriting any: composites decompose through the
    # live glyph set, and a shared glyph must be scaled once
    recorded = {}
    for cp in cps:
        name = cmap[cp]
        if name in recorded:
            continue
        rec = DecomposingRecordingPen(gs)
        gs[name].draw(rec)
        recorded[name] = (cp, rec)
    errors = []
    for name, (cp, rec) in recorded.items():
        target = w_hmtx[w_cmap[cp]][0] * upm / w_upm * kern
        adv, lsb = hmtx[name]
        sx = target / adv if adv else 1.0
        pen = TTGlyphPen(None)
        rec.replay(TransformPen(pen, (sx, 0, 0, sy, 0, 0)))
        glyph = pen.glyph()
        glyph.recalcBounds(glyf)
        glyf[name] = glyph
        new_adv = round(target)
        hmtx[name] = (new_adv, glyph.xMin if glyph.numberOfContours else round(lsb * sx))
        errors.append(abs(new_adv / upm - target / upm))

    if nbsp_alias:
        for table in base["cmap"].tables:
            if table.isUnicode():
                table.cmap[0xA0] = cmap[0x20]
        cps.append(0xA0)

    asc, desc, gap = effective_box(general_box(family, cut), base)
    hhea, os2 = base["hhea"], base["OS/2"]
    hhea.ascent = os2.sTypoAscender = os2.usWinAscent = round(asc * upm)
    hhea.descent = os2.sTypoDescender = -round(desc * upm)
    os2.usWinDescent = round(desc * upm)
    hhea.lineGap = os2.sTypoLineGap = round(gap * upm)
    if os2.version >= 4:
        os2.fsSelection |= 1 << 7  # USE_TYPO_METRICS
    set_names(
        base,
        family["family"],
        CUT_SUBFAMILY[cut],
        f"advance-matched to {Path(spec['word'].partition(':')[2]).stem} hmtx x{kern:.6f}, y x{sy}",
    )
    out = ROOT / spec["out"]
    out.parent.mkdir(parents=True, exist_ok=True)
    base.save(out)
    print(
        f"{family['family']} {cut}: {len(cps)} glyphs, y x{sy}, kern {kern:.6f}, "
        f"max rounding {max(errors) * 1000:.3f} mem -> {out.relative_to(ROOT)} ({out.stat().st_size // 1024} KB)"
    )
    return cps


def css_src(out_path):
    rel = os.path.relpath(ROOT / out_path, ROOT / "apps/docs/src/renderer/fonts")
    return f"url('./{rel}')"


def rewrite_manifest(path, family, ranges):
    general = [f for f in family["faces"] if not f.get("unicodeRange")]
    derived = []
    for cut, spec in family["refit"]["cuts"].items():
        weight, style = cut.split("/")
        face = {"src": css_src(spec["out"]), "unicodeRange": format_range(ranges[cut])}
        if weight != "normal":
            face["weight"] = weight
        if style != "normal":
            face["style"] = style
        derived.append(face)
    family["faces"] = general + derived
    text = json.dumps(family, indent=2, ensure_ascii=True) + "\n"
    path.write_text(text, encoding="utf8")
    subprocess.run(["npx", "prettier", "--write", str(path)], cwd=ROOT, check=True, capture_output=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--family")
    ap.add_argument("--dfonts", default=DFONTS)
    args = ap.parse_args()
    built = 0
    for path in sorted(MANIFEST_DIR.glob("*.json")):
        family = json.loads(path.read_text(encoding="utf8"))
        refit = family.get("refit") or {}
        if refit.get("method") != "advance-matched":
            continue
        if args.family and family["family"] != args.family:
            continue
        ranges = {cut: build_cut(family, cut, spec, args.dfonts, path) for cut, spec in refit["cuts"].items()}
        rewrite_manifest(path, family, ranges)
        built += 1
    if not built:
        sys.exit("no advance-matched family matched")
    print("now run: node tools/build-font-aliases.mjs && python3 tools/build-stand-in-boxes.py")


if __name__ == "__main__":
    main()
