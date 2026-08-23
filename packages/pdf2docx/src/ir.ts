/**
 * Intermediate representation between extraction (PDFium) and rebuild
 * (docx-engine): Page → Block(Text/Image) → Line → Span → Char, plus
 * page-level quality flags. All coordinates are PDF points in page space
 * (origin bottom-left, y up); unit conversion happens only in the rebuild layer.
 */
import type { Rect } from './geometry'
import type { UnicodeScript } from './script'

export type Dir = 'ltr' | 'rtl'

/** One extracted character (the analysis layer's input unit). */
export interface PdfChar {
  /** unicode code point (UTF-32 from FPDFText_GetUnicode) */
  code: number
  /** string form; combining-mark clusters append here, so may exceed one code point */
  text: string
  /** tight glyph box */
  box: Rect
  /** loose box (advance-width based, from FPDFText_GetLooseCharBox) */
  looseBox: Rect
  /** baseline origin */
  originX: number
  originY: number
  /** glyph rotation in radians (0 = horizontal) */
  angle: number
  /** font size in points */
  fontSize: number
  /** CSS-style weight from the font descriptor (400 normal / 700 bold) */
  fontWeight: number
  /** family name with any `ABCDEF+` subset prefix stripped */
  fontFamily: string
  italic: boolean
  /** fill color, hex RRGGBB without '#' */
  color: string
  /**
   * horizontal glyph scale from the char's text matrix (PDF Tz / docx w:w),
   * 1 = unscaled; absent means unmeasured (treated as 1)
   */
  hscale?: number
  /** invisible render mode (PDF Tr 3/7) — Word's hidden formatting marks;
   * carried to the docx as w:vanish so it stays present but unseen (P20) */
  invisible?: boolean
  /** inserted by PDFium's layout analysis (soft spaces / newlines), not in the content stream */
  isGenerated: boolean
  /** source char index in the PDFium text page (extraction bookkeeping, P16 B) */
  pdfCharIndex?: number
  /** end-of-line hyphenation hyphen (drop when joining the split word) */
  isHyphen: boolean
  script: UnicodeScript
  // ── P4 text×shape style mapping (set by analyze/styling, absent otherwise) ──
  /** fill-derived highlight color, hex RRGGBB without '#' */
  highlight?: string
  /** stroke hugging the baseline */
  underline?: boolean
  /** stroke through the glyph's middle band */
  strike?: boolean
  /**
   * footnote anchor (P6): this char stands where a superscript footnote
   * marker was; it renders as a w:footnoteReference to the IrPage footnote
   * with this id (its own text is empty — Word regenerates the number)
   */
  noteRef?: string
}

/** A run of same-styled, same-script text inside one line. */
export interface Span {
  text: string
  box: Rect
  fontSize: number
  fontFamily: string
  bold: boolean
  italic: boolean
  /** hex RRGGBB without '#' */
  color: string
  dir: Dir
  script: UnicodeScript
  /** P4 style mapping — see PdfChar */
  highlight?: string
  underline?: boolean
  strike?: boolean
  // ── character-compression restore (P5): AI-generated documents squeeze text
  // via w:w scale + negative w:spacing; without them rebuilt lines wrap
  // earlier and pages overflow ──
  /** horizontal glyph scale (docx w:w = charScale × 100); absent = 1 */
  charScale?: number
  /** extra per-char advance beyond the font's own, in points (docx w:spacing) */
  charSpacingPt?: number
  /** footnote anchor span (P6) — see PdfChar.noteRef */
  noteRef?: string
  /** invisible source text (PDF Tr 3/7) → w:vanish (P20) */
  invisible?: boolean
}

export interface Line {
  spans: Span[]
  box: Rect
  /** dominant baseline y */
  baseline: number
  /** ends in a hyphenation hyphen: join with the next line without a space, dropping the hyphen */
  endsWithHyphen: boolean
  /**
   * intra-paragraph hard break (P7): the PREVIOUS line ended well short of the
   * wrap edge although this line's first word would still have fit — an
   * intentional line break, rebuilt as <w:br/> instead of a soft join
   */
  hardBreakBefore?: boolean
}

/** list-item annotation (P4): the paragraph is one item of a detected list */
export interface ListInfo {
  kind: 'bullet' | 'ordered'
  /** nesting level, 0-based (capped at 2; deeper items stay plain paragraphs) */
  level: number
  /** ordered items: id shared by one sequential run — the rebuild layer maps it to a docx numId */
  seqId?: number
  /** ordered items: first ordinal of the run (numbering restart value) */
  start?: number
  /** ordered marker style: "1." / "1)" / "(1)" / "3.1.15." (multi-level outline) */
  style?: 'dot' | 'paren' | 'parens' | 'multi'
  /** multi style: the run's first marker split into per-level start ordinals */
  startValues?: number[]
  /** original marker text (already stripped from the item's runs) */
  marker: string
}

export interface TextBlock {
  kind: 'text'
  lines: Line[]
  box: Rect
  /** inferred paragraph alignment ('left' also stands for undecided); visual, not logical */
  align: 'left' | 'center' | 'right' | 'justify'
  /** first-line indent in points (0 = none) */
  firstLineIndentPt: number
  /** paragraph base direction (first-strong over the block's text) */
  dir: Dir
  /** present = this paragraph is a list item (P4); the marker is auto-generated on rebuild */
  list?: ListInfo
  /**
   * present = the paragraph is one TOC dot-leader line (P6). Lines hold the
   * TITLE only; the dots become a right leader tab and `pageNumber` follows it.
   */
  tocEntry?: { level: number; pageNumber: string }
  /**
   * P3 before_space chain: extra whitespace above this block (pt) beyond the
   * page's normal line leading — absolute vertical positions turned flowable.
   * Absent/0 = no extra spacing.
   */
  spacingBeforePt?: number
  /**
   * decorative rule attached to this paragraph (P7): a stray horizontal line
   * mapped onto w:pBdr. `side` is where the LINE sits relative to the text.
   */
  border?: DecorBorder
  /**
   * card membership (P20): index into IrPage.cards. On flow pages the member
   * blocks leave the flow and rebuild INSIDE the card's anchored text box —
   * pinned at absolute coordinates the plate cannot follow a reflow, and one
   * gained line above slides it off its own text (byte-deck FINAL TAKEAWAY).
   */
  cardId?: number
}

/**
 * A card region (P20): an opaque backdrop plate plus the text blocks laid on
 * it. Flow pages rebuild the group as ONE paragraph-anchored text box (wrap
 * topAndBottom) — plate fill, measured insets, text riding the flow together.
 * Canvas pages ignore the tag: plate and text are both absolute there.
 */
export interface CardRegion {
  /** plate box (pt, PDF y-up) */
  box: Rect
  /** plate fill, hex RRGGBB without '#' */
  color: string
  /** plate came from a curved (rounded-rect) path */
  rounded?: boolean
  /** source paint order of the plate (relativeHeight discipline, P16 A) */
  z?: number
}

/** a stray rule rebuilt as a paragraph border (P7; 'left' = accent bar, P14 C) */
export interface DecorBorder {
  side: 'top' | 'bottom' | 'left'
  /** hex RRGGBB without '#' */
  color: string
  /** line thickness in points */
  widthPt: number
  /** gap between the line and the text edge it decorates, in points */
  spacePt: number
  /**
   * how far the line's right end stops short of its column's right edge (pt);
   * rebuilt as w:ind right so the border does not span the whole column.
   * Only set when the line is at least as wide as the text it decorates.
   */
  indentRightPt?: number
  /** left inset from the column edge — standalone (line-only) paragraphs only */
  indentLeftPt?: number
}

/** floating-image placement (P3): wrap mode + offset from the content left edge */
export interface FloatPlacement {
  wrap: 'square-left' | 'square-right' | 'behind'
  /** image left edge relative to the text body's left edge, in points */
  xOffsetPt: number
}

export interface ImageBlock {
  kind: 'image'
  box: Rect
  /** encoded bytes ready to embed */
  data: Uint8Array
  mime: 'image/png' | 'image/jpeg'
  pixelWidth: number
  pixelHeight: number
  /** present = the image floats outside the text flow (anchored, not inline) */
  float?: FloatPlacement
  /** see TextBlock.spacingBeforePt (inline images only) */
  spacingBeforePt?: number
  /**
   * source paint order (top-level page-object index, P16 A): behindDoc
   * anchors on one page stack by this via relativeHeight, so a full-page
   * wallpaper drawn first stays under card panels drawn later.
   */
  z?: number
  /**
   * this panel is a card plate (P20): index into IrPage.cards. The flow
   * rebuild skips its behindDoc pin (the anchored text box paints the plate);
   * canvas pages pin it as usual.
   */
  cardId?: number
}

// ── vector shapes (P2): normalized from page path objects ──

/** A thin horizontal/vertical line — table-border / underline / strikethrough candidate. */
export interface Stroke {
  /** the line's extent as a (thin) rect */
  box: Rect
  orientation: 'h' | 'v'
  /** line thickness in points */
  widthPt: number
  /** hex RRGGBB without '#' */
  color: string
  /**
   * extracted from inside a form XObject (P14 C). Slide exports wrap design
   * dividers in forms; these stay out of the LATTICE stroke pool (a lone
   * cross of dividers would mint a bordered 2×2 table) but still serve the
   * decor / column-separator passes.
   */
  fromForm?: boolean
}

/** A filled rectangle — cell-shading / highlight candidate. */
export interface Fill {
  box: Rect
  /** hex RRGGBB without '#' */
  color: string
  /**
   * source fill alpha (128–254) when the PDF painted this translucently —
   * absent means opaque. Background panels keep it so a dark scrim pinned
   * over a photo backdrop doesn't turn into a solid slab (P11 A).
   */
  alpha?: number
  /** source paint order (see ImageBlock.z) */
  z?: number
}

export interface PageShapes {
  strokes: Stroke[]
  fills: Fill[]
  /** path subpaths P2 ignores (curves, diagonals, non-rect fills) — surfaced as a warning */
  ignoredPaths: number
  /**
   * bboxes of opaque FILLED subpaths with curves (rounded cards, banners) —
   * not part of the fill pool, but light-text backdrop candidates (P10 B)
   */
  curvedFills?: Fill[]
}

// ── raw path data (extraction output; the analysis layer normalizes it) ──

export interface RawSubpath {
  /** page-space points (object matrix already applied) */
  points: Array<{ x: number; y: number }>
  closed: boolean
  /** contains bezier segments (P2 ignores the subpath) */
  hasCurves: boolean
}

export interface RawPath {
  subpaths: RawSubpath[]
  /** fill mode is not none */
  filled: boolean
  stroked: boolean
  /** hex RRGGBB without '#' */
  fillColor: string
  strokeColor: string
  /** stroke width in page units (matrix scale applied) */
  strokeWidth: number
  /** fill alpha 0–255 (P10 C): translucent washes are glows/tints, not shading */
  fillAlpha?: number
  /** lives inside a form XObject (P14 C) — see Stroke.fromForm */
  fromForm?: boolean
  /** top-level page-object index (form children share the form's, P16 A) */
  z?: number
}

// ── tables (P2, lattice) ──

export interface TableCellBlock {
  /** grid-space cell box (merged cells span their full extent) */
  box: Rect
  /** columns this cell spans (docx w:gridSpan) */
  gridSpan: number
  /** vertical merge: 'restart' opens, 'continue' rows are covered placeholders */
  vMerge?: 'restart' | 'continue'
  /** cell shading, hex RRGGBB without '#' */
  fill?: string
  /** vertical content alignment (docx w:vAlign); absent = top */
  vAlign?: 'center' | 'bottom'
  /**
   * edges recovered by the char-respecting split of a merged run (P27):
   * tables that draw vertical rules only in the header row otherwise merge
   * every data row into one full-width cell. A soft edge was never drawn in
   * the source — rebuilders must not paint a border there. Filler cells of a
   * merged side-by-side panel grid (P28) suppress all four edges.
   */
  softEdges?: { left?: boolean; right?: boolean; top?: boolean; bottom?: boolean }
  /** cell content through the regular text pipeline */
  blocks: TextBlock[]
}

export interface TableBlock {
  kind: 'table'
  box: Rect
  /** column widths in points (from the solved grid's column boundaries) */
  colWidthsPt: number[]
  /** row-major; every row covers all grid columns via gridSpan + vMerge placeholders */
  rows: TableCellBlock[][]
  /**
   * detection confidence 0..1 — set by the borderless detectors (stream /
   * form) only; absent = lattice (vector borders, unconditionally trusted).
   * P4's low-confidence downgrade path keys off this.
   */
  confidence?: number
  /**
   * dominant grid stroke color (hex RRGGBB) when it isn't black — the rebuild
   * layer writes it into w:tblBorders (white rulings between zebra fills
   * otherwise render as default black and repaint the table dark)
   */
  borderColor?: string
  /** detected by the checkbox-form pass (P6): cells are form-field results */
  form?: boolean
  /**
   * rule-separated side-by-side zone (P22 A): the drawn separator between the
   * two cells survives as the table's inside-vertical border
   */
  sepRule?: 'single' | 'double'
  /** see TextBlock.spacingBeforePt */
  spacingBeforePt?: number
}

export type PageBlock = TextBlock | ImageBlock | TableBlock

/** one detected footnote (P6): body anchors carry `id` via Span.noteRef */
export interface FootnoteIR {
  /** document-unique id (page-scoped numbering encoded in) */
  id: string
  /** note content without its leading marker digits */
  blocks: TextBlock[]
}

/** Full-page raster fallback for scanned / degraded pages. */
export interface PageRender {
  data: Uint8Array
  mime: 'image/png'
  pixelWidth: number
  pixelHeight: number
}

export interface IrPage {
  /** 0-based page index */
  index: number
  widthPt: number
  heightPt: number
  /** /Rotate in 90° steps (0..3) */
  rotation: number
  /** reading-order blocks; empty when the page fell back to `render` */
  blocks: PageBlock[]
  /** bad ToUnicode map (U+FFFD / private-use ratio over threshold) → bitmap fallback */
  degraded: boolean
  /** no text layer + one page-covering image → bitmap fallback */
  scanned: boolean
  /** page had a structure tree (Tagged PDF); P1 only probes, P2+ may consume it */
  hasStructTree: boolean
  /** why the page degraded (for warnings), e.g. 'bad-tounicode' | 'rotated' | 'vertical-text' */
  degradedReason?: string
  /** present iff scanned or degraded */
  render?: PageRender
  /** full-page background fill color (hex RRGGBB); absent when white/none */
  bgColor?: string
  /**
   * page-covering background stack (gradient/wallpaper) rendered to a bitmap
   * (P9 B); the rebuild layer pins it behind the text as a full-page
   * behindDoc float. Only on normal (non-fallback) pages.
   */
  bgRender?: PageRender
  /**
   * region background panels (P10 B): bottom-z edge-flush solid fills (cover
   * spines, chapter banners) as ready behindDoc floats — the rebuild layer
   * pins them to the page after bgRender, and the text on them stays in flow.
   */
  bgPanels?: ImageBlock[]
  /**
   * card regions (P20): backdrop plates whose text blocks rebuild inside an
   * anchored text box on flow pages. The plate's bgPanels entry carries the
   * matching cardId so the flow rebuild skips its behindDoc pin (canvas pages
   * keep it — everything is absolute there).
   */
  cards?: CardRegion[]
  /** normalized vector shapes (absent on fallback pages) */
  shapes?: PageShapes
  /**
   * vertical layout sections (P3): every non-fallback page carries at least
   * one; a plain page is one section with one column. `blocks` stays the
   * flattened reading-order view of the same content.
   */
  sections?: PageSection[]
  /** analysis-layer notes for this page (negative gaps clamped, …) */
  warnings?: string[]
  /**
   * aggregated layout confidence 0..1 (P4): ToUnicode quality + stream-table
   * confidence + layout warnings. Absent on fallback (scanned/degraded) pages.
   * Below PAGE_CONFIDENCE_MIN the caller downgrades the page to its bitmap.
   */
  confidence?: number
  /** footnotes lifted off the page bottom (P6); anchored via Span.noteRef */
  footnotes?: FootnoteIR[]
  /** decorative lines the P7 decor pass left out: vertical, or hostless over the per-page cap */
  ignoredVerticalDecor?: number
  /**
   * canvas page (P19): high-confidence absolutely-positioned layout (slides).
   * The rebuild layer emits its text blocks as page-anchored containers
   * (w:framePr) in a per-page section instead of stacking them into a flow.
   */
  canvas?: boolean
  /**
   * decor gradient shadings rasterized transparent (P19): pinned behindDoc
   * by the CANVAS path only — flow pages ignore them (dropped, as always).
   */
  decorImages?: ImageBlock[]
}

// ── multi-column layout (P3) ──

export interface PageColumn {
  box: Rect
  /** column blocks, top → bottom */
  blocks: PageBlock[]
}

export interface PageSection {
  box: Rect
  /** columns in READING order (RTL-dominant sections list right column first) */
  columns: PageColumn[]
  /** gaps between adjacent columns in LAYOUT order (pt); length = columns − 1 */
  gutterWidthsPt: number[]
  /** column reading direction: 'rtl' = right column first */
  dir: Dir
  /** a drawn rule runs inside a gutter → w:cols w:sep (P14 C) */
  colSep?: boolean
}
