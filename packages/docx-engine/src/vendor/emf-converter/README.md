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
  (>1% per edge). Word maps the frame onto the picture extent, so a
  bounds-tight canvas got stretched to the frame's aspect by the display box
  (e.g. OLE preview text filling a third of a page-wide frame drew giant and
  deformed); matching frames keep the established bounds mapping
- EMF+ dual-mode files (EmfPlusHeader flag bit 0) replay the EMF+ records only; the
  GDI fallback records are painted just inside an EmfPlusGetDC window (MS-EMFPLUS
  2.3.1.3), instead of drawing every shape and caption twice
- EMR_EXTTEXTOUTW honours the Dx advance array (per-glyph placement, ETO_PDY),
  TA_UPDATECP current-position text, ETO_CLIPPED / ETO_OPAQUE rectangles and the
  symbol-font PUA remap; substituted fonts no longer overrun or hide neighbours
- EmfPlusDrawDriverString: GlyphPos follows the glyph array without padding (odd
  glyph counts read garbage positions), every glyph is placed at its own position,
  the optional TransformMatrix is applied, glyph-index strings (no CmapLookup) skip
- EMF+ page coordinates map through the GDI bounds origin and canvas scale
  (deviceMap) so dual-mode fills line up with their GDI text and survive canvas
  clamping; nested metafiles keep the plain dpi scale
- EMF+ images are painted in z-order: files with image objects replay twice — a dry
  pass collects the payloads, they are decoded, the real pass draws each in place
- Continued EMF+ objects (64 KiB chunks) reassemble across EMR_COMMENT records and
  finish when the declared TotalObjectSize is reached; every chunk drops its 4-byte
  size prefix (only the first was skipped, so multi-chunk PNG/JPEG payloads and nested
  metafiles came out corrupt or never completed)
- LOGFONTW FaceName read at offset +32 (upstream read +28, landing in the
  OutPrecision/ClipPrecision/Quality/PitchAndFamily bytes — GDI+-generated EMFs
  set those non-zero, prefixing the family with control chars; the invalid CSS
  ident made `ctx.font` assignment fail silently and text drew at the default
  10px); mapFontFamily now also strips control chars and quotes non-generic
  families so a corrupt facename can never drop the font size again
- EmfPlusBitmap BitmapDataType compared against the MS-EMFPLUS values (Pixel = 0,
  Compressed = 1; upstream tested 1/2, so no EMF+ image object ever decoded and
  every EmfPlusDrawImage painted nothing — e.g. a PowerPoint master background
  wrapped as EMF+ rendered as a blank white page)
- EMR_CREATEDIBPATTERNBRUSHPT / EMR_CREATEMONOBRUSH implemented as repeating canvas
  patterns (upstream left them unhandled, so the previously selected solid brush
  painted the PATCOPY blits — Excel OLE previews drew their dotted cell borders as
  solid lines); the tile is blown up to the device scale without smoothing, since GDI
  tiles a pattern brush in device pixels (a 1-px checkerboard at 2x otherwise averages
  to grey once the picture is fitted to the page)
- Font strings carry a `sans-serif` generic fallback after the (quoted) facename; an
  unknown facename otherwise fell to the browser default serif, so CJK Office text in
  EMF previews (Meiryo UI, MS PGothic) drew in Mincho/Song
- EMR_EXTTEXTOUTW horizontal alignment tests the TA_CENTER/TA_RIGHT bit pair as a
  value (upstream tested `& 2` after `& 6`, so TA_CENTER (6) fell through to
  right-aligned and an OLE icon's centered caption lost its leading characters)
- EMF+ ObjectType constants follow MS-EMFPLUS 2.1.1.22 (Region = 4,
  ImageAttributes = 8; upstream had the two swapped) and a region object with
  RegionNodeCount 0 (a single leaf node) parses. Office chart pictures clip the
  plot with SetClipRect and widen the clip back with SetClipRegion; the region
  never resolved, so titles, legends and axis labels outside the plot area
  vanished
