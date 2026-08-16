# Vendored: emf-converter

Upstream: emf-converter 2.0.2 (Apache-2.0) — https://github.com/ChristopherVR/emf-converter

`index.mjs` / `index.d.mts` are the published npm dist files, plus local WMF fixes
(marked below); wrapper-level changes still go in `src/metafile.ts`.

Local modifications to `index.mjs` (2026-08, GenOffice):
- LOGFONT16 FaceName read at offset +18 (upstream read +14, landing in the
  precision/quality bytes, so every WMF font family came out empty)
- Font escapement parsed and applied (rotated text), TA_* text alignment honored
- Symbol fonts (Wingdings/Webdings/Symbol/…) remap chars to the U+F0xx PUA
- META_PATBLT filled with the current brush (PATCOPY/BLACKNESS/WHITENESS)
- META_DIBCREATEPATTERNBRUSH approximated as a solid brush (DIB average color)
- Unhandled object-creating records still claim a WMF object-table slot, keeping
  SELECTOBJECT indices aligned
- EMF window→viewport mapping normalized back into canvas space: gmx/gmy/gmw/gmh
  now offset by rclBounds and scale by canvas/bounds (files with
  SETVIEWPORTEXTEX drew at reference-device scale — ignoring dpiScale — and
  files with a non-zero bounds origin drew fully off-canvas); viewport defaults
  changed to an identity window→viewport mapping to match
- META_DIBBITBLT / META_DIBSTRETCHBLT / META_STRETCHDIB implemented (upstream
  dropped every WMF-embedded DIB, e.g. OLE preview icons); SRCAND/SRCPAINT/
  SRCINVERT approximated with multiply/lighter/difference composites; the
  bitmap-less variant is detected per MS-WMF 2.3.1.2/2.3.1.3
  (RecordSize words == (RecordFunction >> 8) + 3 — an earlier local version
  added the 6-byte header on top and never matched)
- WMF without a placeable header derives logical bounds from the leading
  SETWINDOWORG/SETWINDOWEXT records instead of assuming 800×600
- EMR_ALPHABLEND implemented (Word 2007-era OLE icon previews draw the icon
  bitmap with it): 32bpp DIB decoded with its real alpha channel (AC_SRC_ALPHA
  premultiplied colors un-premultiplied), SrcConstantAlpha applied via
  globalAlpha, source sub-rect scaled onto the destination
- EMF canvases replay in rclFrame space (converted to device units via
  szlDevice/szlMillimeters) when rclBounds clearly disagree with the frame
  (>5% per edge). Word maps the frame onto the picture extent, so a
  bounds-tight canvas got stretched to the frame's aspect by the display box
  (e.g. OLE preview text filling a third of a page-wide frame drew giant and
  deformed); matching frames keep the established bounds mapping
