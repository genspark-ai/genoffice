/** author/date of one tracked change (a w:ins or w:del wrapper). */
export interface RevisionInfo {
  author: string
  /** ISO timestamp from w:date */
  date?: string
  /** original w:id, kept so an edited paragraph re-emits stable ids */
  id?: string
}

/** A styled text run inside a paragraph-like block. */
export interface Run {
  text: string
  /**
   * Original <w:rPr> slice (serialized from the parse tree). Written back via
   * mergeRPrModel when the run is rebuilt: unmodeled properties (caps/vanish/dstrike/
   * bdr/double underline/themeColor/all four rFonts slots…) are kept verbatim, and
   * modeled fields are only rebuilt when they differ from the raw encoding (i.e. edited).
   */
  rawRPr?: string
  /** character style id (w:rStyle); "Hyperlink" is implied by `link` and never stored here */
  styleId?: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  /** hex color without '#', e.g. "FF0000" */
  color?: string
  /** font size in half-points (OOXML w:sz) */
  sizeHalfPoints?: number
  /** primary font family (w:rFonts, eastAsia ?? ascii ?? hAnsi) */
  font?: string
  /** Latin-slot font (w:rFonts ascii/hAnsi) when the run declares one; may equal `font`.
   * Kept separate so editing one script's font never flattens the other slot. */
  fontAscii?: string
  /** Complex-script font (w:rFonts cs/cstheme). Display only; saving is kept faithful by rawRPr */
  csFont?: string
  /** Character spacing (w:spacing, twips, may be negative). Display only; saving is kept faithful by rawRPr */
  charSpacingTwips?: number
  /** Horizontal character scale percent (w:w). Display only (approximated as spacing); saving is kept faithful by rawRPr */
  charScalePct?: number
  /** OOXML named highlight color (w:highlight), e.g. "yellow" */
  highlight?: string
  /** superscript / subscript (w:vertAlign) */
  vertAlign?: 'superscript' | 'subscript'
  /** East Asian emphasis mark (w:em). Display only; saving is kept faithful by rawRPr */
  em?: 'dot' | 'comma' | 'circle' | 'underDot'
  /** hyperlink info; rId references an existing relationship in the original docx */
  link?: { href: string; rId?: string; tooltip?: string }
  /**
   * ids of comments whose range covers this run. Only set when the whole
   * commentRangeStart..End pair lives inside the same paragraph, so an edited
   * paragraph can re-emit the markers without touching its neighbors.
   */
  commentIds?: string[]
  /** run is inside a w:ins (tracked insertion) */
  ins?: RevisionInfo
  /** run is inside a w:del (tracked deletion; text came from w:delText) */
  del?: RevisionInfo
  /**
   * The run is a footnote/endnote reference marker (w:footnoteReference /
   * w:endnoteReference). `text` holds the display number; the marker itself
   * is what saves (Word renumbers automatically).
   */
  noteRef?: { kind: 'footnote' | 'endnote'; id: string }
  /**
   * An index entry (XE) field attached AFTER this run's text. Invisible in
   * Word; kept in the run model so marked paragraphs stay editable.
   */
  xeTerm?: string
  /**
   * The run is a cross-reference (REF field) to the named bookmark; `text`
   * holds the cached display result. Regenerates as a full REF field.
   */
  refField?: string
  /** exact original REF instruction text; written back verbatim so switches (\r, \p, ...) survive */
  refInstr?: string
  /** Generic inline field (DATE/TIME/NUMPAGES/FILENAME etc.): full instruction text; run text is the cached result */
  instrField?: string
  /**
   * Run-level formatting revision (w:rPrChange): records the review info and the modeled
   * subset of the pre-revision formatting. Accept = keep the current formatting and clear
   * the record; reject = restore the formatting from `old` and clear the record.
   */
  rPrChange?: RevisionInfo & {
    old?: {
      bold?: boolean
      italic?: boolean
      underline?: boolean
      strike?: boolean
      color?: string
      sizeHalfPoints?: number
      font?: string
      fontAscii?: string
      charSpacingTwips?: number
      charScalePct?: number
      highlight?: string
      vertAlign?: 'superscript' | 'subscript'
      styleId?: string
    }
  }
  /**
   * The run is an atomic inline formula. `text` holds the flat OMML token
   * strip (previews / word count); `omml` is the exact <m:oMath> fragment,
   * emitted verbatim when the paragraph regenerates.
   */
  math?: { omml: string }
  /**
   * Phonetic guide (w:ruby). `text` holds the base characters; `rt` the
   * phonetic text; `xml` the exact <w:ruby> fragment, re-wrapped in a w:r
   * and emitted verbatim when the paragraph regenerates.
   */
  ruby?: { rt: string; xml: string }
  /**
   * Inline picture in a table cell. `dataUrl` is for display; `xml` is the
   * exact <w:drawing> fragment, re-wrapped in a w:r and emitted verbatim
   * when the cell regenerates.
   */
  image?: { dataUrl: string; widthPx?: number; heightPx?: number; xml: string }
}

/** One comment from word/comments.xml (read-only display model). */
export interface CommentInfo {
  id: string
  author: string
  initials?: string
  /** ISO timestamp from w:date */
  date?: string
  /** plain text, paragraphs joined with \n */
  text: string
  /** Reply: the parent comment's w:id (resolved via the paraId relation in commentsExtended) */
  parentId?: string
  /** Resolved (w15:done) */
  done?: boolean
  /** w14:paraId of the comment's last paragraph (the commentsExtended link key; assigned on save for new comments) */
  paraId?: string
}

export type ParaAlign = 'left' | 'center' | 'right' | 'justify' | 'distribute'

/**
 * Preserved shell of a structured document tag (w:sdt) wrapping this block.
 * The original <w:sdt>…<w:sdtPr>…</w:sdtPr>…<w:sdtContent> opening bytes and
 * the closing </w:sdtContent></w:sdt> bytes are stored here so that saving can
 * re-wrap the regenerated paragraph XML in the original SDT shell (byte-fidelity rule).
 */
export interface SdtShell {
  /** human-visible alias label (w:sdtPr/w:alias @w:val), may be empty */
  alias: string
  /** programmatic tag (w:sdtPr/w:tag @w:val), may be empty */
  tag: string
  /**
   * control type: 'text' = plain/rich text, 'date' = date picker,
   * 'dropdown' = drop-down list/combo box, 'checkbox' = Word 2013 checkbox,
   * 'unknown' = unrecognised sdtPr content
   */
  controlType: 'text' | 'date' | 'dropdown' | 'checkbox' | 'unknown'
  /**
   * Original XML bytes from the opening <w:sdt> up to and including the
   * <w:sdtContent> tag (i.e. everything BEFORE the paragraph content).
   * This is emitted verbatim before the regenerated paragraph XML on save.
   */
  openXml: string
  /**
   * Closing bytes "</w:sdtContent></w:sdt>" (emitted after the paragraph XML).
   */
  closeXml: string
  /**
   * Set when one w:sdt was split into several blocks (multi-paragraph
   * sdtContent, e.g. a Word TOC). Blocks sharing a group belong to the same
   * sdt; only the first carries openXml and only the last carries closeXml.
   */
  group?: number
}

/** One custom tab stop from w:tabs/w:tab. Positions are in twips. */
export interface TabStop {
  /** position in twips (w:pos) */
  pos: number
  /** tab alignment (w:val) */
  val: 'left' | 'center' | 'right' | 'decimal' | 'bar' | 'clear'
  /** leader character (w:leader): none/dot/hyphen/underscore/heavy/middleDot */
  leader?: 'none' | 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot'
}

/** Paragraph-level formatting that survives regeneration (subset of w:pPr). */
export interface ParaFormat {
  /** w:jc */
  align?: ParaAlign
  /** line spacing as a multiple of single (w:spacing w:line / 240, lineRule="auto") */
  lineSpacing?: number
  /**
   * F1 precise line-height rule (from w:spacing w:lineRule):
   * 'auto'    = multiple spacing; lineSpacing is the multiplier (e.g. 1.0 = single, 1.5 = 1.5x)
   * 'atLeast' = line height at least N twips (lineRawTwips), else the font's natural height
   * 'exact'   = fixed line height of N twips (lineRawTwips), never expands
   * undefined = same as 'auto' (backward compatible)
   */
  lineRule?: 'auto' | 'atLeast' | 'exact'
  /** raw w:spacing w:line value in twips (used in atLeast/exact modes) */
  lineRawTwips?: number
  /** left indent in twips (w:ind w:left) */
  indentLeft?: number
  /** right indent in twips (w:ind w:right) */
  indentRight?: number
  /** first-line indent in twips; positive = w:firstLine, negative = w:hanging */
  indentFirstLine?: number
  /** space above the paragraph in twips (w:spacing w:before) */
  spaceBefore?: number
  /** space below the paragraph in twips (w:spacing w:after) */
  spaceAfter?: number
  /** start the paragraph on a new page (w:pageBreakBefore) */
  pageBreakBefore?: boolean
  /** keep this paragraph on same page as the next paragraph (w:keepNext) */
  keepNext?: boolean
  /** keep all lines of this paragraph on the same page (w:keepLines) */
  keepLines?: boolean
  /** widow/orphan control (w:widowControl; Word default=true) */
  widowControl?: boolean
  /** CJK-Latin/digit auto spacing (w:autoSpaceDE/DN; Word default=true); false when both are explicitly off */
  autoSpace?: boolean
  /** ignore space-after when followed by same-style para (w:contextualSpacing) */
  contextualSpacing?: boolean
  /** paragraph shading fill, hex without '#' (w:shd w:fill) */
  shadingFill?: string
  /** paragraph borders, subset of "tblr" e.g. "b" or "tblr" (w:pBdr, single lines) */
  borders?: string
  /** custom tab stops from w:tabs (non-empty overrides default 0.5in grid) */
  tabStops?: TabStop[]
  /** first-line drop cap (w:framePr w:dropCap="drop|margin") */
  dropCap?: { type: 'drop' | 'margin'; lines: number }
  /**
   * RTL paragraph (w:bidi). align stores the visual value: per Word's quirk,
   * w:jc left/right swap meaning in bidi paragraphs; parsing converts to the
   * visual direction and writing back converts again.
   */
  bidi?: boolean
}

/**
 * Document grid (w:docGrid) — a key property for Chinese documents.
 * With type=lines, each line's height is rounded up to align with linePitch (twips).
 */
export interface DocGrid {
  /** w:type: 'default'|'lines'|'linesAndChars'|'snapToChars', default 'default' */
  type: 'default' | 'lines' | 'linesAndChars' | 'snapToChars'
  /** line pitch (twips, w:linePitch); effective for lines/linesAndChars */
  linePitch?: number
  /** characters per line (derived from w:charSpace; effective for linesAndChars/snapToChars) */
  charSpace?: number
}

/** Page setup stored in the trailing w:sectPr. All lengths in twips. */
export interface SectionSettings {
  pageWidth: number
  pageHeight: number
  orientation: 'portrait' | 'landscape'
  marginTop: number
  marginRight: number
  marginBottom: number
  marginLeft: number
  /** simple single-line box page border (w:pgBorders) */
  pageBorder: boolean
  /** number of text columns (w:cols w:num), 1 = normal */
  columns: number
  /** column gap (w:cols w:space, twips; OOXML default 720) */
  colSpace?: number
  /** Header distance from the page top (w:pgMar w:header, twips; default 720). A tall header pushes the body down */
  headerDist?: number
  /** Footer distance from the page bottom (w:pgMar w:footer, twips; default 720). A tall footer pushes the body up */
  footerDist?: number
  /** Vertical alignment of page content (sectPr w:vAlign): center/both/bottom; default top */
  vAlign?: 'top' | 'center' | 'both' | 'bottom'
  /** line grid (w:docGrid), key for Chinese official documents; affects line-height rounding */
  docGrid?: DocGrid
  /** text flow direction (sectPr w:textDirection), e.g. tbRl = vertical CJK; absent = horizontal lrTb */
  textDirection?: string
}

/** One section: settings + start type + block ownership + header/footer refs (read-only enumeration; editing still goes through sectPr XML) */
export interface SectionInfo {
  settings: SectionSettings
  /** Section-break type w:type (how this section starts; meaningless for the first section), default nextPage */
  startType: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage'
  /** docxIndex range of this section's blocks (inclusive; the section-break paragraph / trailing hidden sectPr block belong to this section) */
  firstBlockIndex: number
  lastBlockIndex: number
  /** raw sectPr XML fragment of this section */
  sectPrXml: string
  /** different first page for this section (w:titlePg) */
  titlePg: boolean
  /** page renumbering start value (w:pgNumType w:start); absent = continue from the previous section */
  pageNumberStart?: number
  /** page number format (w:pgNumType w:fmt): decimal/lowerRoman/upperLetter/... */
  pageNumberFmt?: string
  /** header/footer reference rIds, by variant */
  headerRefs: Partial<Record<'default' | 'first' | 'even', string>>
  footerRefs: Partial<Record<'default' | 'first' | 'even', string>>
}

/** one rich paragraph of a header / footer part */
export interface HfParagraph extends ParaFormat {
  runs: Run[]
  /** layout-table row: one entry per cell, rendered as columns. Display-only —
   *  saving keeps the part's original w:tbl bytes and never serializes cells. */
  cells?: HfTableCell[]
}

/** one cell of a header/footer layout-table row */
export interface HfTableCell {
  runs: Run[]
  align?: ParaAlign
  /** column width as % of the row (w:tcW, falling back to tblGrid) */
  widthPct?: number
}

/** one w:lvl of a numbering definition */
export interface NumberingLevel {
  /** w:numFmt, e.g. "decimal" | "bullet" | "lowerLetter" | "upperRoman" | "chineseCountingThousand" */
  numFmt: string
  /** w:lvlText, e.g. "%1." or "%1.%2" (placeholders %n = counter of level n-1) or a bullet glyph */
  lvlText: string
  /** w:start, default 1 */
  start: number
  /** w:pPr/w:ind left (twips); fallback geometry for items without their own w:ind */
  indentLeft?: number
  /** w:pPr/w:ind hanging (twips): width reserved for the number/bullet marker */
  hanging?: number
  /** w:rPr/w:sz of the marker itself (half-points) */
  szHalfPoints?: number
  /** w:rPr/w:rFonts of the marker (bullet glyphs in Symbol/Wingdings need decoding) */
  font?: string
}

/** one w:num entry from word/numbering.xml (abstractNum levels + overrides applied) */
export interface NumberingDef {
  numId: string
  /** counters continue across w:num entries sharing an abstractNum */
  abstractNumId: string
  /** ilvl -> level definition */
  levels: Record<number, NumberingLevel>
  /** ilvl -> w:lvlOverride/w:startOverride value (restart markers) */
  startOverrides: Record<number, number>
}

/** Content for the default page header / footer. */
export interface HeaderFooter {
  text: string
  /** append an automatic page number (PAGE field), footer only */
  pageNumber?: boolean
  /**
   * Rich paragraphs; when present they are the content source of truth and
   * the part regenerates one w:p per entry (PAGE_MARK \u2014 or a user-typed '#'
   * when no PAGE_MARK exists \u2014 becomes the PAGE field when pageNumber is
   * set). Absent = legacy single centered line.
   */
  paras?: HfParagraph[]
}

/** Placeholder for NUMPAGES (total pages) fields in headers/footers; the renderer substitutes the total page count and saving writes the field back (PAGE uses PAGE_MARK the same way) */
export const TOTAL_PAGES_MARK = '\uE000'
/** Placeholder for PAGE fields in headers/footers; a private-use character so literal '#' text in the part can never be mistaken for the field position */
export const PAGE_MARK = '\uE001'

/** display-only image in a header/footer part (logos etc.; saving preserves the part's bytes) */
export interface HfImage {
  dataUrl: string
  widthPx?: number
  heightPx?: number
  /** wp:anchor floating (otherwise inline) */
  floating?: boolean
}

/** Parsed content of one header/footer part (any w:type variant). */
export interface HfPartInfo {
  /** plain text, PAGE fields shown as PAGE_MARK, NUMPAGES as TOTAL_PAGES_MARK */
  text: string
  hasPageNumber: boolean
  paras: HfParagraph[]
  /** images in the part (display-only; text edits do not affect their bytes) */
  images?: HfImage[]
}

/** One display run of footnote/endnote body text (for page-bottom rendering; editing still uses plain text) */
export interface NoteRun {
  text: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  /** hex without '#' */
  color?: string
  sizeHalfPoints?: number
}

/** One footnote or endnote (display/edit model; separators are preserved separately). */
export interface NoteInfo {
  /** original w:id, unique within its part */
  id: string
  /** plain text, paragraphs joined with \n */
  text: string
  /** rich display runs (one group per paragraph); gone when an edit rebuilds the note as plain text */
  richParas?: NoteRun[][]
}

/** Display/edit model for a protected field paragraph. */
export interface FieldDisplay {
  /** tocLine: "text .... page"; text: plain visible result; pageBreak: invisible marker */
  kind: 'tocLine' | 'text' | 'pageBreak'
  /** visible text (tocLine: entry title; text: whole result) */
  left?: string
  /** page number for tocLine */
  right?: string
  /** TOC outline level 1-9 */
  level?: number
  /** bookmark anchor of the tocLine hyperlink (_Toc...), for click-to-jump */
  anchor?: string
  /** numbering marker of the entry's w:numPr ("1.", "1.1."), computed at parse time */
  num?: string
}

/** Editable text tokens inside an OMML formula; the surrounding math tree is preserved. */
export interface FormulaDisplay {
  tokens: string[]
  /** MathML Core markup for native 2D rendering (absent = flat token fallback) */
  mathml?: string
  /** the paragraph's <m:oMath> fragments, so token edits can re-derive mathml */
  omml?: string
  /** LaTeX source recovered by ommlToLatex; enables full re-editing (absent = token edits only) */
  latex?: string
}

/** One data series of an embedded chart, read from the cached values in its chart part. */
export interface ChartSeries {
  name?: string
  /** cached numeric values by category position; null = gap in the cache */
  values: (number | null)[]
}

/**
 * Display/edit model of an embedded chart, extracted from its chart part
 * (word/charts/chartN.xml). Edits patch only cached texts/numbers in that
 * part; the chart structure stays untouched. The embedded workbook is NOT
 * updated, so Word's "Edit Data" sheet will still show the pre-edit numbers.
 */
export interface ChartDisplay {
  /** zip path of the chart part this model was read from */
  partPath: string
  kind: 'bar' | 'line' | 'pie' | 'area' | 'other'
  /** chart title; undefined when the chart has no title part (then not editable) */
  title?: string
  categories: string[]
  series: ChartSeries[]
  /** display size (from wp:extent), editable via the corner resize handle */
  widthPx?: number
  heightPx?: number
}

/** A new chart to embed at save time (becomes word/charts/chartN.xml + relationship). */
export interface NewChart {
  kind: 'bar' | 'line' | 'pie'
  title?: string
  categories: string[]
  series: Array<{ name: string; values: (number | null)[] }>
}

/** Editable interchange model of one table cell (type === 'table' blocks). */
export interface TableParagraph extends ParaFormat {
  runs: Run[]
  /** list paragraph inside the cell (w:numPr) */
  list?: { kind: 'bullet' | 'ordered'; numId: string; ilvl: number }
}

/** One cell border (a w:tcBorders child element) */
export interface CellBorder {
  /** w:val, e.g. single/dashed/double; 'none' means explicitly no border */
  style: string
  /** w:sz, in 1/8 pt units */
  szEighths?: number
  /** hex without '#', or 'auto' */
  color?: string
}

export interface CellBorders {
  top?: CellBorder
  left?: CellBorder
  bottom?: CellBorder
  right?: CellBorder
}

/** Table-level borders (w:tblBorders), including inner horizontal/vertical lines */
export interface TableBorders extends CellBorders {
  insideH?: CellBorder
  insideV?: CellBorder
}

/** Cell margins (twips); unspecified sides inherit the table-level/Word defaults (0 top/bottom, 108 left/right) */
export interface CellMargins {
  top?: number
  left?: number
  bottom?: number
  right?: number
}

export interface TableCell {
  /** paragraph texts inside the cell */
  paras: string[]
  /** rich paragraph content; when present this is the formatting source of truth */
  richParas?: TableParagraph[]
  /** this cell's margin overrides (w:tcMar) */
  cellMarTwips?: CellMargins
  /** nested tables inside this cell (shown as read-only sub-tables; bytes kept faithful by the outer table) */
  nestedTables?: TableModel[]
  /** per nested table: how many cell paragraphs precede it (reading-order anchor); absent = after all paragraphs */
  nestedTableAnchors?: number[]
  colSpan?: number
  /** 'restart' opens a vertical merge, 'continue' is a merged-away cell */
  vMerge?: 'restart' | 'continue'
  /** cell shading fill, hex without '#' (w:shd w:fill) */
  fill?: string
  /** first-run text color, hex without '#' */
  color?: string
  bold?: boolean
  align?: ParaAlign
  /** vertical alignment (w:vAlign): top (default)/center/bottom */
  vAlign?: 'top' | 'center' | 'bottom'
  /** vertical text (w:textDirection): tbRl = vertical right-to-left, btLr = horizontal rotated 90° (bottom-up) */
  textDirection?: 'tbRl' | 'btLr'
  /** legacy horizontal merge (w:hMerge, intermediate parse state): continue cells fold into the restart cell to their left */
  hMerge?: 'restart' | 'continue'
  /** cell borders (w:tcBorders); undefined = no tcBorders (table-level/style borders apply) */
  borders?: CellBorders
  /** raw <w:tcPr>…</w:tcPr> bytes: surgically patched on regeneration so unmodeled
   *  properties (tcMar/textDirection etc.) are not lost */
  rawTcPr?: string
  /** cell-level revision (tcPr w:cellIns / w:cellDel) */
  cellRevision?: { kind: 'ins' | 'del' } & RevisionInfo
}

/** Editable interchange model of a table block, extracted from or generated to OOXML. */
export interface TableModel {
  rows: TableCell[][]
  /** relative column widths in percent (from w:tblGrid) */
  colWidthsPct?: number[]
  /** absolute column widths in twips (w:tblGrid); only when every gridCol > 0 */
  colWidthsTwips?: number[]
  /** table width as a percentage of the body width (w:tblW type="pct"); absolute widths are ignored when set */
  widthPct?: number
  /** table-level default cell margins (w:tblCellMar) */
  cellMarTwips?: CellMargins
  /** table-level borders (w:tblBorders); undefined = not declared (default gridlines apply) */
  borders?: TableBorders
  /** table horizontal alignment (tblPr w:jc); default left */
  align?: 'left' | 'center' | 'right'
  /** table left indent (w:tblInd, twips; effective only with left alignment) */
  indentTwips?: number
  /** per-row height (twips, w:trHeight; null = not set), aligned with rows */
  rowHeightsTwips?: Array<number | null>
  /** per-row height rule (w:trHeight w:hRule; absent/legacy models = atLeast), aligned with rows */
  rowHeightRules?: Array<'atLeast' | 'exact' | null>
  /** per-row raw <w:trPr>…</w:trPr> bytes (passed through verbatim), aligned with rows */
  rawTrPrs?: Array<string | null>
  /** row-level revisions (trPr w:ins/w:del = inserted/deleted row), aligned with rows */
  rowRevisions?: Array<({ kind: 'ins' | 'del' } & RevisionInfo) | null>
  /** table style reference (w:tblStyle w:val); '' = explicitly no style, undefined = leave as-is */
  tblStyleId?: string
  /** RTL table (tblPr w:bidiVisual): columns display right to left */
  bidiVisual?: boolean
}

/**
 * One ink annotation (freehand strokes) to write at save time. Becomes a floating
 * anchored picture (wp:anchor, in front of text) inside its anchor paragraph;
 * word/media/... + relationship are created like NewImage.
 */
export interface NewInkImage {
  /** index into the finalBlocks array of the anchor block */
  blockIndex: number
  /** transparent PNG bytes, base64 encoded */
  base64: string
  widthPx: number
  heightPx: number
  /** offset from the text column's left edge (may be negative) */
  offsetXPx: number
  /** offset from the top of the anchor paragraph (may be negative) */
  offsetYPx: number
  /** opaque editor payload stored in wp:docPr descr (stroke vectors) */
  payload?: string
}

/** One ink annotation read back from a document (display/restore model). */
export interface InkInfo {
  /** docxIndex of the paragraph carrying the anchored drawing */
  anchorIndex: number
  offsetXPx: number
  offsetYPx: number
  widthPx: number
  heightPx: number
  /** embedded PNG as data URL; null when the media part is unreadable */
  dataUrl: string | null
  /** editor stroke payload from wp:docPr descr; null when absent */
  payload: string | null
}

/**
 * Text wrapping of a floating image (wp:anchor):
 * square-* = square wrap (text flows around, image on left/right), topBottom = top and
 * bottom, behind = behind text, front = in front of text.
 */
/**
 * Wrap modes. tight/through still render approximated as square, but the model keeps
 * them distinct for fidelity: the wrap element is not rebuilt unless the wrap is
 * explicitly changed, and repositioning reuses the original wrapTight/wrapThrough
 * bytes (including wp:wrapPolygon) as a whole.
 */
export type ImageWrap =
  | 'square-left'
  | 'square-right'
  | 'tight-left'
  | 'tight-right'
  | 'through-left'
  | 'through-right'
  | 'topBottom'
  | 'behind'
  | 'front'

/** A new image to embed at save time (becomes word/media/... + relationship). */
export interface NewImage {
  /** raw image bytes, base64 encoded */
  base64: string
  mime: 'image/png' | 'image/jpeg' | 'image/gif'
  widthPx: number
  heightPx: number
  /** paragraph alignment for the image (w:jc) */
  align?: 'left' | 'center' | 'right'
  /** floating wrap mode; absent = inline */
  wrap?: ImageWrap
  /** rotation in degrees clockwise (a:xfrm rot) */
  rotDeg?: number
  /** mirror flips (a:xfrm flipH/flipV) */
  flipH?: boolean
  flipV?: boolean
}

export type BlockType = 'paragraph' | 'heading' | 'listItem' | 'table' | 'image' | 'passthrough'

export interface Block {
  /** stable id for editor bookkeeping */
  id: string
  type: BlockType
  /**
   * Index of the top-level element in <w:body> of the ORIGINAL document.xml.
   * This is the patch anchor. null for blocks created in the editor.
   */
  docxIndex: number | null
  /** exact XML slice from the original document.xml (patch passthrough basis) */
  originalXml: string | null
  /** heading level 1-9 (type === 'heading') */
  level?: number
  /** original w:pStyle val, if any */
  styleId?: string
  /** list info (type === 'listItem') */
  list?: { kind: 'bullet' | 'ordered'; numId: string; ilvl: number }
  /** paragraph-level formatting (paragraph / heading / listItem) */
  format?: ParaFormat
  /** user bookmarks starting in this paragraph (w:bookmarkStart, _internal names excluded) */
  bookmarks?: string[]
  /** Word internal bookmarks (_Ref/_Toc/_Hlk…): hidden from UI, re-emitted on rebuild so REF/TOC anchors survive editing */
  hiddenBookmarks?: string[]
  /** Cross-paragraph comment ranges: ids whose start falls in this paragraph only (end is in a later one). Re-emitted at paragraph start on rebuild to avoid orphaned endpoints */
  commentStarts?: string[]
  /** Cross-paragraph comment ranges: ids whose end falls in this paragraph only. end+reference are re-emitted at paragraph end on rebuild */
  commentEnds?: string[]
  /**
   * Exact original <w:pPr> slice (paragraph / heading / listItem). Regenerated
   * paragraphs reuse it verbatim (text-only edits) or merge format changes
   * into it, so pPr constructs the format model cannot express (keepNext,
   * tabs, paragraph-mark rPr, pPrChange revisions...) survive editing.
   */
  rawPPr?: string
  /** rich text content (paragraph / heading / listItem) */
  runs?: Run[]
  /** human label for protected blocks (localized strings like "Table 3×4", "Image", "Chart") */
  label?: string
  /** plain-text preview for protected blocks */
  previewText?: string
  /** data URL for inline display (type === 'image', or the packaged OLE preview on passthrough) */
  imageDataUrl?: string
  /** OLE embed ProgID (o:OLEObject), e.g. "Excel.Sheet.12"; drives the friendly type caption */
  oleProgId?: string
  /** display size from wp:extent in CSS px (type === 'image') */
  imageWidthPx?: number
  imageHeightPx?: number
  /** paragraph alignment of the image (w:jc) */
  imageAlign?: 'left' | 'center' | 'right'
  /** text wrapping of a floating image (wp:anchor); absent = inline (in line with text) */
  imageWrap?: ImageWrap
  /**
   * Horizontal/vertical posOffset (in EMU) of a floating image (wp:anchor
   * with wp:positionH/V using posOffset, not align). Used for free-position
   * drag. Absent when the image uses named alignment (left/center/right) or
   * is inline.
   */
  imageOffsetXEmu?: number
  imageOffsetYEmu?: number
  /** margin-relative wp:align of a floating image (Word position-gallery presets) */
  imagePosH?: 'left' | 'center' | 'right'
  imagePosV?: 'top' | 'center' | 'bottom'
  /** picture rotation in degrees clockwise 0-359 (pic a:xfrm rot / 60000) */
  imageRotDeg?: number
  /** picture mirror flips (pic a:xfrm flipH/flipV) */
  imageFlipH?: boolean
  imageFlipV?: boolean
  /** editable structure (type === 'table'); untouched original XML still saves byte-identically */
  table?: TableModel
  /** display-only rendering for field passthrough paragraphs (TOC lines etc.) */
  fieldDisplay?: FieldDisplay
  /** true for body-trailing elements (w:sectPr) that are never shown in the editor */
  hidden?: boolean
  /** protected drawing that is just a decorative rule; render as a line, not a chip */
  decorative?: boolean
  /** stroke display of a decorative rule (a:ln); absent = default 1px full-width line */
  ruleColorHex?: string
  ruleThicknessPx?: number
  ruleWidthPx?: number
  /** picture whose rel/media is broken (or metafile-only): render an empty frame + alt text */
  brokenImage?: boolean
  /**
   * Invisible range marker leaked to body top level (w:bookmarkEnd, w:proofErr…):
   * renders as nothing but keeps its body position. Not `hidden` — hidden blocks
   * are moved to the body tail on save, which would relocate the marker.
   */
  invisibleMarker?: boolean
  /** display-only content of anchored textboxes (code boxes, callouts) */
  textboxes?: TextboxDisplay[]
  /** text tokens of an OMML formula, in document order */
  formulaDisplay?: FormulaDisplay
  /** display/edit model of an embedded chart (set on chart-labeled blocks) */
  chartDisplay?: ChartDisplay
  /**
   * When this block was wrapped in a w:sdt (structured document tag), the
   * original sdtPr + shell bytes are stored here for save-time re-wrapping.
   * The block's runs/format are from sdtContent and are fully editable;
   * saving patches only those runs while keeping the sdt shell verbatim.
   */
  sdtShell?: SdtShell
  /**
   * Move revision: 'from' = content was moved away (shown with strikethrough + red bg),
   * 'to' = content was moved here (shown with green bg). Accept/reject via del/ins marks.
   */
  moveRevision?: 'from' | 'to'
  /**
   * When this paragraph has a tracked paragraph-property change (w:pPrChange),
   * this carries the revision author/date for the review UI badge and navigation.
   * The raw pPrChange XML is preserved in rawPPr and used for accept/reject.
   */
  pPrChangeInfo?: {
    author: string
    date?: string
    id?: string
    old?: ParaFormat & {
      type?: 'docParagraph' | 'docHeading' | 'docListItem'
      styleId?: string
      level?: number
      kind?: 'bullet' | 'ordered'
      numId?: string
      ilvl?: number
    }
  }
  /** Top-level block insertion/deletion wrapper (w:ins/w:del around w:p or w:tbl). */
  blockRevision?: { kind: 'ins' | 'del' } & RevisionInfo
}

/** one paragraph inside an anchored textbox */
export interface TextboxParaDisplay extends ParaFormat {
  runs: Run[]
}

/**
 * Display model of an anchored textbox. Text edits are patched back into the
 * original XML via patchTextboxParas; everything else saves byte-identical.
 */
export interface TextboxDisplay {
  /** hex without '#' */
  fill?: string
  /** hex without '#' */
  borderColor?: string
  /** box width from wp:extent in CSS px (render fidelity) */
  widthPx?: number
  /**
   * fixed box height in CSS px. Only set when the shape's autofit is off,
   * i.e. Word renders the box at this exact height and clips overflow.
   */
  heightPx?: number
  /** original fixed height; edited boxes may grow but never shrink below this */
  minHeightPx?: number
  /** text body insets from wps:bodyPr, in CSS px */
  insetTopPx?: number
  insetRightPx?: number
  insetBottomPx?: number
  insetLeftPx?: number
  /**
   * DrawingML preset geometry name (a:prstGeom prst="...") for shape rendering.
   * Absent for plain rectangular textboxes (rect = default).
   */
  prst?: string
  /**
   * WordArt preset id (e.g. 'wordArt-1').  When set, the editor applies
   * large-text CSS approximation (color + optional text-stroke).
   * This field is display-only; it is never written to OOXML.
   */
  wordArtId?: string
  /**
   * Content holds tables / content controls flattened into display lines, so
   * paras do not map 1:1 onto the w:txbxContent w:p children the patch-save
   * path rewrites. The editor must not offer text editing for this box or a
   * commit would corrupt the table (sidebars).
   */
  readOnly?: boolean
  paras: TextboxParaDisplay[]
}

/** Final body content decided by the editor at save time. */
export type FinalBlock =
  { source: 'original'; docxIndex: number } | { source: 'generated'; block: GeneratedBlock }

/** Content the generator can turn into an OOXML <w:p> fragment. */
export interface GeneratedBlock {
  type: 'paragraph' | 'heading' | 'listItem'
  level?: number
  styleId?: string
  list?: { kind: 'bullet' | 'ordered'; numId: string; ilvl: number }
  format?: ParaFormat
  /** emit this exact <w:pPr> instead of rebuilding it (see Block.rawPPr) */
  rawPPr?: string
  /** bookmark names to re-emit at the paragraph start (w:bookmarkStart/End pair) */
  bookmarks?: string[]
  /** internal (_-prefixed) bookmark names, re-emitted before `bookmarks` */
  hiddenBookmarks?: string[]
  /** cross-paragraph comment range starts to re-emit at paragraph start */
  commentStarts?: string[]
  /** cross-paragraph comment range ends (+reference) to re-emit at paragraph end */
  commentEnds?: string[]
  runs: Run[]
  /**
   * When this block was originally wrapped in a w:sdt, re-wrap the generated
   * paragraph XML with this shell on save (byte-fidelity rule for the sdt structure).
   */
  sdtShell?: SdtShell
  /**
   * JSON {author,date?,id?} when the block originally had a pPrChange.
   * Set to null by the editor when the user accepts/rejects the pPrChange revision;
   * applyRawPPr then strips <w:pPrChange> from rawPPr so the saved file is clean.
   */
  pPrChange?: string | null
  /** Used by editor dirty detection; SaveBlock carries the actual wrapper metadata. */
  blockRevision?: ({ kind: 'ins' | 'del' } & RevisionInfo) | null
}

/** display-only formatting a paragraph style contributes (for on-screen fidelity) */
export interface StyleDisplay {
  sizeHalfPoints?: number
  /** hex without '#' */
  color?: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  font?: string
  /** latin-slot font when it differs from the east-asian one (w:ascii/w:hAnsi) */
  fontAscii?: string
  /** complex-script font (w:rFonts cs/cstheme) */
  csFont?: string
  /** character spacing (rPr w:spacing, twips, may be negative) */
  charSpacingTwips?: number
  /** multiple of single line spacing (w:spacing w:line, lineRule auto) */
  lineSpacing?: number
  /** F1: lineRule for accurate height calculation */
  lineRule?: 'auto' | 'atLeast' | 'exact'
  /** F1: raw line twips for atLeast/exact */
  lineRawTwips?: number
  spaceBeforeTwips?: number
  spaceAfterTwips?: number
  /** indents from the style definition (w:pPr w:ind, twips; hanging is stored as a negative firstLine) */
  indentLeftTwips?: number
  indentRightTwips?: number
  indentFirstLineTwips?: number
  /** F1: keepNext from style definition */
  keepNext?: boolean
  /** keepLines from style definition */
  keepLines?: boolean
  /** F1: contextualSpacing from style definition */
  contextualSpacing?: boolean
  /** paragraph alignment from the style (w:pPr w:jc) */
  align?: 'left' | 'center' | 'right' | 'justify'
  /** CJK-Latin/digit auto spacing from the style pPr (w:autoSpaceDE/DN) */
  autoSpace?: boolean
}

/** display-only formatting a table style contributes (fills applied per tblLook flags) */
export interface TableStyleDisplay {
  /** whole-table cell shading (hex without '#') */
  fill?: string
  /** whole-table run formatting (style-level w:rPr) */
  wholeTable?: { color?: string; bold?: boolean }
  /** conditional first-row formatting (w:tblStylePr w:type="firstRow") */
  firstRow?: { fill?: string; bold?: boolean; color?: string }
  /** conditional first/last-column and last-row formatting (w:tblStylePr) */
  firstCol?: { fill?: string; bold?: boolean; color?: string }
  lastCol?: { fill?: string; bold?: boolean; color?: string }
  lastRow?: { fill?: string; bold?: boolean; color?: string }
  /** odd-band row shading (band1Horz) */
  band1Fill?: string
  /** even-band row shading (band2Horz) */
  band2Fill?: string
  /** style-level table borders (w:tblPr w:tblBorders); fallback when the document level has none */
  borders?: TableBorders
  /** style-level cell margins (w:tblPr w:tblCellMar) */
  cellMarTwips?: CellMargins
  /** style-level w:pPr spacing applied to every cell paragraph (TableGrid carries
   * after=0/line=240 — the reason Word tables are tight while body text is loose) */
  paraSpacing?: {
    beforeTwips?: number
    afterTwips?: number
    lineRawTwips?: number
    lineRule?: 'auto' | 'atLeast' | 'exact'
    lineSpacing?: number
  }
}

export interface StyleInfo {
  styleId: string
  name: string
  type: 'paragraph' | 'character' | 'table'
  headingLevel?: number
  /** w:semiHidden — Word itself hides it from the style gallery (e.g. DefaultParagraphFont) */
  semiHidden?: boolean
  /** w:qFormat — candidate for Word's quick style gallery */
  qFormat?: boolean
  /** character shell linked to a paragraph style ("Heading 1 Char"); Word does not show it separately */
  linkedCharShell?: boolean
  /** rendering hints from the style definition, basedOn chain resolved; never written back on save */
  display?: StyleDisplay
  /** table-style rendering hints (type === 'table') */
  tableDisplay?: TableStyleDisplay
  /** w:pPr/w:numPr on the style — list numbering referenced via pStyle (ListBullet/ListNumber) */
  numPr?: { numId: string; ilvl: number }
  /** w:default="1" — Word applies this style to every paragraph without a w:pStyle */
  isDefault?: boolean
}

/** document-wide defaults from styles.xml w:docDefaults (display-only) */
export interface DocDefaults {
  sizeHalfPoints?: number
  /** default Latin font (rPrDefault w:rFonts w:ascii) */
  asciiFont?: string
  /** default Chinese font (rPrDefault w:rFonts w:eastAsia) — determines the CJK line-height factor */
  eastAsiaFont?: string
  /** bold/italic/color from rPrDefault (display-layer defaults, used by canvas CSS) */
  bold?: boolean
  italic?: boolean
  /** hex without '#' */
  color?: string
  lineSpacing?: number
  /** F1: lineRule for accurate height calculation */
  lineRule?: 'auto' | 'atLeast' | 'exact'
  /** F1: raw line twips for atLeast/exact */
  lineRawTwips?: number
  /** F1: paragraph spacing before default (twips); undefined = not set in docDefaults */
  spaceBeforeTwips?: number
  /** F1: paragraph spacing after default (twips); undefined = not set in docDefaults */
  spaceAfterTwips?: number
}

/** word/settings.xml w:documentProtection (only the editing restriction subset). */
export interface DocProtection {
  /** w:edit value, e.g. "readOnly" | "comments" | "trackedChanges" | "forms" */
  edit: string
  /** w:enforcement="1" — restriction is active */
  enforced: boolean
  /** password hash (base64, Word 2013+ iterated SHA-512 scheme); absent when there is no password protection */
  hash?: string
  /** salt (base64) */
  salt?: string
  /** iteration count (w:cryptSpinCount), Word default 100000 */
  spinCount?: number
  /** w:cryptAlgorithmSid, 14 = SHA-512 (the only supported value) */
  algorithmSid?: number
}

/** One bibliography source (subset of the Word b:Sources schema). */
export interface SourceInfo {
  /** unique tag (b:Tag), e.g. "Wang2024" */
  tag: string
  /** b:SourceType, e.g. "Book" | "JournalArticle" | "InternetSite" | "Report" */
  type: string
  /** display author string (b:Author/b:Author/b:NameList or b:Corporate) */
  author: string
  title: string
  year: string
  /** publisher / journal / site name, optional */
  publisher?: string
  /** url for internet sources, optional */
  url?: string
}

/** Font pair from word/theme/theme1.xml (latin typefaces + east-asian). */
export interface ThemeFonts {
  /** a:majorFont a:latin typeface (headings) */
  major: string
  /** a:minorFont a:latin typeface (body) */
  minor: string
  /** a:minorFont a:ea typeface (body east-asian), optional */
  eastAsia?: string
  /** a:majorFont a:ea typeface (heading east-asian), optional */
  majorEastAsia?: string
  /** a:majorFont / a:minorFont a:cs typefaces (complex script), optional */
  majorCs?: string
  minorCs?: string
  /** settings.xml w:themeFontLang w:eastAsia — picks Word's face for empty EA theme slots */
  eaLang?: string
}

/** Color scheme values (hex without '#') from a:clrScheme. */
export interface ThemeColors {
  name?: string
  /** read-only slots, populated by readThemeColors for w:themeColor resolution */
  dk1?: string
  lt1?: string
  hlink?: string
  folHlink?: string
  dk2?: string
  lt2?: string
  accent1?: string
  accent2?: string
  accent3?: string
  accent4?: string
  accent5?: string
  accent6?: string
}

export interface ParsedDoc {
  blocks: Block[]
  /** comments from word/comments.xml, in file order */
  comments: CommentInfo[]
  /** footnotes from word/footnotes.xml (separators excluded), file order */
  footnotes: NoteInfo[]
  /** endnotes from word/endnotes.xml (separators excluded), file order */
  endnotes: NoteInfo[]
  /** bibliography sources from the customXml b:Sources part */
  sources: SourceInfo[]
  /** aidocs-ink annotations (freehand strokes) found on body paragraphs, file order */
  inks: InkInfo[]
  /** theme font pair from word/theme/theme1.xml, null when the doc has no theme part */
  themeFonts?: ThemeFonts | null
  /** color scheme from word/theme/theme1.xml, null when the doc has no theme part */
  themeColors?: ThemeColors | null
  /** editing restriction from word/settings.xml, null when none */
  protection: DocProtection | null
  /** plain text of the default page header, null when the document has none */
  headerText?: string | null
  /** rich paragraphs of the default header (PAGE fields appear as PAGE_MARK runs) */
  headerParas?: HfParagraph[] | null
  /** rich paragraphs of the default footer */
  footerParas?: HfParagraph[] | null
  /** display-only images (logos etc.) in the default header / footer */
  headerImages?: HfImage[] | null
  footerImages?: HfImage[] | null
  /** text watermark (VML textpath) in the default header, null when none */
  watermarkText?: string | null
  /** plain text of the default page footer (PAGE fields appear as PAGE_MARK) */
  footerText?: string | null
  /** the default footer contains an automatic page number field */
  footerHasPageNumber?: boolean
  /** the default header contains an automatic page number field */
  headerHasPageNumber?: boolean
  /** "different first page" (w:titlePg in a sectPr) */
  titlePg?: boolean
  /** "different odd & even pages" (settings.xml w:evenAndOddHeaders) */
  evenAndOddHeaders?: boolean
  /** first-page header/footer parts (w:type="first"), null when absent */
  headerFirst?: HfPartInfo | null
  footerFirst?: HfPartInfo | null
  /** even-page header/footer parts (w:type="even"), null when absent */
  headerEven?: HfPartInfo | null
  footerEven?: HfPartInfo | null
  /** rId → header/footer part content (multi-section docs look these up via each section's sectPr refs) */
  hfParts?: Record<string, HfPartInfo>
  /** styleId -> info, from word/styles.xml */
  styles: Map<string, StyleInfo>
  /** document-wide display defaults from styles.xml */
  docDefaults?: DocDefaults
  /** heading level -> styleId present in this document */
  headingStyleIds: Map<number, string>
  /** styleId used for list paragraphs, if the doc has one (usually "ListParagraph") */
  listParagraphStyleId?: string
  /** numId -> full numbering definition, from word/numbering.xml */
  numbering: Map<string, NumberingDef>
  /** internal: everything needed to patch-save */
  internal: {
    originalBytes: Uint8Array
    documentXml: string
    /** inner body range [start, end) in documentXml covered by top-level elements */
    bodyInnerStart: number
    bodyInnerEnd: number
  }
}
