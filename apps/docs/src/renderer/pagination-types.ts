// Records shared by the pagination modules: measured blocks, page slices,
// section geometry and the patch outputs of a slicing pass.
import type { SectionInfo } from '@genoffice/docx-engine'

export interface BlockBox {
  top: number
  height: number
  /** paragraph pageBreakBefore: force a page break before the block */
  breakBefore?: boolean
  /** breakBefore comes from a leading w:br (real break character, not the
   *  pageBreakBefore property): honored even on the document's blank first page */
  breakBeforeBr?: boolean
  /** block contains a page-break field (w:br type=page): force a page break after it */
  breakAfter?: boolean
  /** zero-height break carrier (floating-textbox anchor): the break survives a blank page */
  breakForce?: boolean
  /** block contains a column break (w:br type=column): force a column change after it (new page on last column) */
  colBreakAfter?: boolean
  /** source DOM block (filled during canvas measurement, used to position page-gap decorations) */
  el?: HTMLElement
  /** the block's docxIndex (DOM data-idx; new unsaved blocks lack one) */
  docxIndex?: number
  /** owning section index (filled by assignSections) */
  section?: number
  /** CSS-floated block (square/tight/through image wrap, w:tblpPr table): the
   *  wrapped text beside it carries the vertical extent, so it consumes no
   *  column height itself (block boxes in normal flow stack ignoring floats) */
  floated?: boolean
  /** floating-anchor wrapper whose height is a column-spanning wrap band:
   *  Word keeps the box on its anchor's page and lets the band overflow the
   *  bottom margin — fill the column instead of pushing the block whole */
  bandKeep?: boolean
  /** true bottom of the bandKeep wrapper's floating boxes (px, may exceed
   *  top + height by border/rounding): the closing page must reach it or the
   *  preview treats the box as spilling and re-pins it */
  floatBottom?: number
  /** floated w:tblpPr table with vertAnchor page/margin: target Y on the landing
   *  page (px from the page top for 'page', from the content top for 'margin');
   *  the engine resolves it into a --tblp-dy shift (SliceOutputs.floatVShifts) */
  pageRelVyPx?: number
  pageRelVAnchor?: 'page' | 'margin'
  /** in-block line boundaries (relative to block top, ascending, each = a line's starting Y): used to split page-crossing blocks by line */
  lineOffsets?: number[]
  /** min lines kept on each side of a split (widow/orphan control): paragraphs 2, table rows 1 (default) */
  splitMinLines?: number
  /** no visible text/image (empty paragraph mark): flows but doesn't count toward column-balance quotas */
  emptyPara?: boolean
  /** non-reflowable block's rendered width (tables, protected textboxes/objects):
   *  such a block never advances into a narrower column (Word turns the page instead) */
  fixedWidthPx?: number

  // ── F2 line-level page-split extensions ─────────────────────────────────
  /**
   * Paragraph line-box list (from computeLineMetrics, for line-level page splitting).
   * When absent, degrades to F1 block-level greedy placement.
   */
  lineBoxes?: Array<{ offsetInBlock: number; height: number }>
  /** space before (px), from line-metrics output */
  spaceBeforePx?: number
  /** space after (px), from line-metrics output */
  spaceAfterPx?: number
  /** total page-bottom footnote reservation folded into `height` by
   *  applyBlockMeta (never into spaceAfterPx: page-bottom exemptions must not
   *  hand it back) */
  footnoteExtraPx?: number
  /** per-reference reservations at their marker offsets (same in-block space
   *  as lineBoxes): a split paragraph charges each note on the page holding
   *  its reference line. Absent = the whole reservation rides the last line. */
  noteBands?: Array<{ offset: number; height: number }>
  /** break-only paragraph (page-break field, no other content): single line's height, drives absorb-vs-blank-page placement */
  breakOnlyLineH?: number
  /** ink height of the block's sole line when it exceeds the column capacity
   *  (an oversized inline picture): the line has no break points, so Word
   *  starts it at a fresh column top and overflow-clips it at the column
   *  bottom instead of painting into later pages (SliceOutputs.oversizeClips) */
  oversizeLineH?: number

  // ── F2 pagination constraints ───────────────────────────────────────────
  /** keepLines: all lines of the paragraph must be on the same page */
  keepLines?: boolean
  /** keepNext: the paragraph and the next paragraph's first line must be on the same page */
  keepNext?: boolean
  /** widowControl: false = widow/orphan protection off (Word default on) */
  widowControl?: boolean

  // ── F2 table row-level page-split extensions ─────────────────────────────
  /**
   * Table row data (from parseDocx).
   * When present, table rows become the page-split unit (instead of hard pixel cuts).
   */
  tableRows?: TableRowBox[]

  /** virtual endnotes-area block (appended by appendEndnotesBlock; no DOM/docxIndex) */
  isEndnotes?: boolean

  /** virtual trailing block reserving pages for overflowing floating boxes (appendFloatSpillBlock) */
  isFloatSpill?: boolean

  /** table block under Word 2013+ layout rules (see BlockMeta.modernTableHeaders) */
  modernTableHeaders?: boolean
}

/**
 * Table row box (for F2 table row-level page splitting).
 */
export interface TableRowBox {
  /** row height (px) */
  height: number
  /** cantSplit: the row cannot be broken internally (whole row stays on one page) */
  cantSplit?: boolean
  /** tblHeader: the row is a header row, repeated at the top of the next page after a break */
  isHeader?: boolean
  /** vertical merge (vMerge continue): the row continues a merged row; its height is not counted independently */
  vMergeContinue?: boolean
  /** in-row safe cut points (relative to row top, px, ascending): spanning all cells without splitting any text line/image.
   *  Word allows in-row page breaks by default; without cut points or with cantSplit the row is atomic */
  cutYs?: number[]
  /** bottom of the lowest content band (text/image rects) relative to row top (px, 0 = empty row):
   *  lets pagination clip declared-height fill below the content instead of pushing pages */
  contentBottom?: number
  /** declared atLeast trHeight (px): reserved space Word never breaks inside —
   *  when it overflows the page remainder the whole row pushes to the next page */
  minHPx?: number
}

/** One column of a multi-column page: a content range in the continuous flow (in-column break semantics match pages) */
export interface PageColumn {
  start: number
  end: number
  /** table continued into the column: header-row range repeated at column top (virtual coordinates) */
  repeatHeader?: { top: number; height: number }
}

/** A column-flow region within a page (a continuous column-count change can stack multiple regions vertically on one page) */
export interface PageRegion {
  /** region top relative to the page content-area top (px) */
  top: number
  /** available height per column within the region (px) */
  height: number
  /** owning section index of the region (for column count/width) */
  section: number
  /** content ranges per column (ascending; single-column regions have length 1) */
  columns: PageColumn[]
}

/** Content range [start, end) shown on one page; height ≤ the owning section's page content height */
export interface PageSlice {
  start: number
  end: number
  /** owning section index of this page */
  section: number
  /**
   * tblHeader repetition: this page starts mid-table, so the source table's header
   * rows must render first. top/height is the header rows' range in the continuous
   * flow (virtual coordinates); the preview clones and crops accordingly.
   */
  repeatHeader?: { top: number; height: number }
  /**
   * Column flow: provided when this page has cols>1 regions (omitted for single-column
   * pages; consumers use the original path). start/end is still the whole-page flow
   * range (= first column start .. last column end); the span can reach columns × column height.
   */
  regions?: PageRegion[]
  /**
   * Physical content height actually used on a regioned page (last region top +
   * its tallest column). The virtual span (end - start) can exceed the page's
   * content height by stacking columns; canvas gap padding/compression uses this.
   */
  physHeight?: number
}

/** Pagination geometry for one section */
export interface SectionGeom {
  contentHeight: number
  /** content height of the section's FIRST page when it differs (w:titlePg
   *  header/footer variant push-down); pages render per-variant margins, so
   *  slicing with the default capacity clips/underfills the first page */
  firstContentHeight?: number
  /** content width (px, page width minus side margins); optional so height-only callers/tests can omit it */
  contentWidth?: number
  /** content-area top offset from the page edge (px, header-expanded top margin);
   *  resolves page-anchored w:tblpY targets into content coordinates */
  topPx?: number
  /** section start forces a page break (nextPage/evenPage/oddPage, or continuous with different page geometry) */
  forceBreak: boolean
  /** section break type: evenPage/oddPage need physical blank pages inserted to align parity */
  startType?: SectionInfo['startType']
  /** equal-width column count (w:cols w:num, default 1): page capacity = columns × column height */
  cols?: number
  /** per-column widths (px, length cols) — differ under w:equalWidth="0" */
  colWidths?: number[]
  /** nextColumn start with the same column count as the previous section: advance one column at the boundary */
  colBreakStart?: boolean
}

/** A split declared-height row needs its DOM stretched: target total height for the row */
export interface RowFillPatch {
  /** owning table block's virtual top (block identity within one layout pass) */
  blockTop: number
  /** row index (same tr filtering/order as domTableRows) */
  row: number
  /** row total target height (px); the DOM tr gets it as a minimum */
  targetPx: number
}

/** Page-anchored floated table: downward shift placing it at its w:tblpY target */
export interface FloatVShiftPatch {
  /** owning table block's virtual top (block identity within one layout pass) */
  blockTop: number
  /** shift below the flow position (px, ≥0); applied as the --tblp-dy margin */
  dyPx: number
}

/** Single line taller than its column (oversized inline picture): renderer-baked page-bottom clip */
export interface OversizeClipPatch {
  /** owning block's virtual top (block identity within one layout pass) */
  blockTop: number
  /** landing column's capacity (px); the block DOM is clipped to it */
  clipPx: number
}

export interface SliceOutputs {
  rowFills?: RowFillPatch[]
  floatVShifts?: FloatVShiftPatch[]
  oversizeClips?: OversizeClipPatch[]
}

/** Per-section header/footer content heights (px); sectionGeoms uses these to compute body push-down */
export interface SectionHfHeights {
  headerPx: number
  footerPx: number
  /** titlePg first-page variant heights (only set for sections with w:titlePg):
   *  the section's first page renders these strips, so its capacity differs */
  firstHeaderPx?: number
  firstFooterPx?: number
}

/** One block's mixed-column canvas placement (consumed by editor/column-layout.ts) */
export interface ColumnBlockPlacement {
  el: HTMLElement
  /** column width (px); absent in single-column regions */
  widthPx?: number
  /** owning section's content width (--doc-content-w): tables resolve their spill/centering caps against it */
  contentWPx?: number
  /** owning section's side margins (--doc-margin-left/right overrides) */
  marginLeftPx?: number
  marginRightPx?: number
  /** owning section's typed docGrid pitch (pt); 0 = untyped section in a
   *  mixed-grid doc (opts out like snapToGrid=0) */
  gridPitchPt?: number
  /** owning section's character-grid letter-spacing delta (pt, docGrid charSpace/4096) */
  charSpacePt?: number
  dx: number
  dy: number
}

export interface MeasuredContent {
  blocks: BlockBox[]
  totalHeight: number
  /** absolutely-positioned boxes of floating-textbox anchors (virtual coords, shift-neutral) */
  floats: FloatBox[]
  /** docxIndex of section-break chips present in the DOM but excluded from the block
   *  list (zero-height in Word): liveSections must still see their boundaries */
  sectBreaks: Set<number>
}

/** one floating textbox/shape box: DOM element + gapless virtual position */
export interface FloatBox {
  el: HTMLElement
  top: number
  height: number
  /** gapless virtual position of the anchor wrapper (the box's page follows it) */
  anchorTop: number
  /** page-pinned box: `top` is raw page-relative Y, not a flow position */
  pinned: boolean
  /** page/margin-relative V rendered from the anchor: `top` - `anchorTop` is
   *  the page-relative Y; the box belongs at that offset on the anchor's page */
  pageRelV: boolean
}

/** Page-bottom footnote entry (number/text/estimated height): shared by canvas page gaps and the pagination preview */
export interface PageNoteItem {
  no: number
  id: string
  text: string
  /** note body has no self-reference mark run: Word draws the entry without a numeral */
  noRefMark?: true
  height: number
  /** resolved note-style line height (px): the renderers set it inline so the entry's line boxes match the reservation model */
  lineHeightPx?: number
  /** resolved note-style base font size (pt) */
  fontSizePt?: number
  /** rich display runs (one group per paragraph); omitted for unformatted footnotes, rendering falls back to plain text */
  richParas?: Array<
    Array<{
      text: string
      bold?: boolean
      italic?: boolean
      underline?: boolean
      strike?: boolean
      color?: string
      sizeHalfPoints?: number
      caps?: 'all' | 'small' | 'none'
    }>
  >
}

/**
 * Block-level pagination constraints (injection channel from parse-layer results into DOM-measured blocks).
 * The canvas/preview measureBlocks only has geometry; semantics like keepNext are attached by docxIndex.
 */
export interface BlockMeta {
  keepNext?: boolean
  keepLines?: boolean
  /** pageBreakBefore (direct or style-level): force a page break before the block */
  breakBefore?: boolean
  /** false only when explicitly disabled (Word default on) */
  widowControl?: false
  /** table blocks: per-tr header/unsplittable/reserved-height flags (applied by fillLineBoxes when collecting rows) */
  tableRowFlags?: Array<{ isHeader: boolean; cantSplit: boolean; minHPx?: number }>
  /** Word 2013+ layout (settings compatibilityMode >= 15): a multirow tblHeader block that doesn't fit the remaining space pushes the table to a fresh page */
  modernTableHeaders?: boolean
  /** page-bottom height reserved for footnote refs inside the block (px): merged into the block height (consumes page capacity like Word's note area) */
  footnoteExtraPx?: number
  /** per-reference reservation heights in run order (matched to the block's
   *  marker elements so a split paragraph charges each note at its line) */
  footnoteBands?: Array<{ heightPx: number }>
}

export type BlockMetaOf = (docxIndex: number) => BlockMeta | undefined
