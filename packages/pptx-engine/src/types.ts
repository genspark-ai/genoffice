/**
 * pptx-engine data model
 *
 * Design principles:
 * 1. The in-house model is the core asset (Canva-style): rendering / editing /
 *    saving all revolve around it.
 * 2. Byte-level fidelity: every parseable element carries a dual anchor —
 *    `slideXmlRange` (original XML byte range) + `originalXml` (original slice).
 *    On save, untouched elements pass through as their original bytes.
 * 3. Inheritance chain pre-resolved: elements carry `resolved` final styles
 *    (merged slide→layout→master→theme); the render layer uses them directly and
 *    never walks inheritance. The original pptx's layout/master/theme are never
 *    written back.
 * 4. Risky content is passthrough: only a placeholder is shown, original bytes
 *    pass through on save.
 *
 * Coordinate unit: EMU (English Metric Unit, 1 inch = 914400 EMU, 1 px@96dpi = 9525 EMU).
 * Angle unit: 60000ths of a degree (OOXML native), with a convenience `deg` field.
 */

// ── Geometry ───────────────────────────────────────────────────────────

export interface EmuRect {
  /** Top-left x (EMU), relative to the containing coordinate system (slide or owning group's child coords) */
  x: number
  /** Top-left y (EMU) */
  y: number
  /** Width (EMU) */
  cx: number
  /** Height (EMU) */
  cy: number
}

export interface Transform {
  offset: EmuRect
  /** Rotation in 1/60000 degree units; deg = rot / 60000 */
  rot: number
  flipH: boolean
  flipV: boolean
}

// ── Color / fill / stroke (inheritance resolved to final values) ────────

/** Resolved final color: uniform #RRGGBB(AA) or 'none' */
export type ResolvedColor = string

export type Fill =
  | { type: 'none' }
  | { type: 'solid'; color: ResolvedColor }
  | {
      type: 'gradient'
      stops: Array<{ pos: number; color: ResolvedColor }>
      angle?: number
      /** <a:path path>: circle/rect/shape = radial/path gradient (linear by default) */
      path?: 'circle' | 'rect' | 'shape'
      /** <a:fillToRect> insets as fractions (may exceed 0..1); defines the gradient focus */
      fillTo?: { l: number; t: number; r: number; b: number }
    }
  | {
      type: 'image'
      mediaRef: string
      mode?: 'stretch' | 'tile'
      /** <a:blip><a:alphaModFix amt> (0-1, translucent picture fills e.g. washed-out backgrounds) */
      alpha?: number
      /** <a:stretch><a:fillRect> insets as fractions: the image maps into this subrect of the shape */
      fillRect?: { l: number; t: number; r: number; b: number }
      /** <a:blip><a:duotone>: [dark, light] colors mapped over image luminance (theme texture backgrounds) */
      duotone?: [string, string]
      /** <a:blip><a:clrChange>: pixels matching `from` are replaced with `to` (#RRGGBB or #RRGGBBAA; alpha 0 = color-to-transparent) */
      clrChange?: { from: string; to: string }
      /** <a:tile>: offsets (EMU), scale fractions and anchor alignment of the tile grid */
      tile?: { tx: number; ty: number; sx: number; sy: number; algn: string }
    }
  | { type: 'pattern'; fg: ResolvedColor; bg: ResolvedColor; preset: string }

/** OOXML arrowhead size: sm/med/lg (width or length direction), med by default */
export type ArrowEndSize = 'sm' | 'med' | 'lg'

/** Line-end arrow decoration (<a:headEnd>/<a:tailEnd>) */
export interface ArrowEnd {
  /** Arrowhead type */
  type: 'arrow' | 'triangle' | 'stealth' | 'diamond' | 'oval' | 'none'
  /** Width size (horizontal, med by default) */
  w?: ArrowEndSize
  /** Length size (vertical/depth, med by default) */
  len?: ArrowEndSize
}

export interface Stroke {
  fill: Fill
  /** Line width (EMU) */
  width: number
  dash?: string
  cap?: 'flat' | 'round' | 'square'
  /** Line head decoration <a:headEnd> (omitted when absent or none) */
  headEnd?: ArrowEnd
  /** Line tail decoration <a:tailEnd> (omitted when absent or none) */
  tailEnd?: ArrowEnd
}

/** Glow <a:glow>: color + radius (EMU) */
export interface GlowEffect {
  color: ResolvedColor
  radius: number
}

/** Outer shadow <a:outerShdw> (the most common effectLst entry) */
export interface ShadowEffect {
  color: ResolvedColor
  /** Blur radius (EMU) */
  blurRad: number
  /** Offset distance (EMU) */
  dist: number
  /** Direction (degrees, clockwise, 0 = right) */
  dirDeg: number
}

// ── Text ───────────────────────────────────────────────────────────────

/** A run of contiguous same-format text (maps to <a:r>); line breaks/soft returns split into separate runs or paragraphs */
export interface TextRun {
  text: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  /** Original underline style (sng/dbl/wavy…); underline is the display boolean, write-back restores from this */
  underlineStyle?: string
  strike?: boolean
  /** Original strikethrough style (sngStrike/dblStrike); strike is the display boolean, write-back restores from this */
  strikeStyle?: string
  /** Font size (pt) */
  fontSize?: number
  /** Run has no explicit sz (inherits); rebuild/injected rPr omits sz to avoid baking in the master font size */
  fontSizeImplicit?: boolean
  /** Letter spacing <a:rPr spc> (pt, may be negative; PowerPoint stores 1/100pt) */
  letterSpacing?: number
  /** Font family (final font name after theme inheritance, for render/editor display) */
  fontFamily?: string
  /**
   * Original <a:latin>/<a:ea> typeface text (incl. +mj-lt/+mn-ea theme refs).
   * Present = the user has not changed the font: patches keep the original bytes
   * and regeneration writes the original values back, so CJK/Latin dual fonts and
   * theme linkage survive; cleared when the user changes the font (along with
   * fontImplicit)
   */
  latinFont?: string
  eaFont?: string
  /** Original <a:cs> typeface (complex-script bucket), written back on the rebuild path */
  csFont?: string
  /** Run has no explicit font declaration (inherits/theme); patches don't inject latin/ea when the font is unchanged */
  fontImplicit?: boolean
  /**
   * Effective character casing ('all' | 'small', explicit rPr cap or inherited from
   * placeholder styles). Display-only: the render layer uppercases; never written back.
   */
  cap?: string
  color?: ResolvedColor
  /** Text highlight color <a:rPr><a:highlight> (drawn as a background behind the run) */
  highlight?: ResolvedColor
  /** color is display-only (from schemeClr/inheritance, not an explicit run srgbClr);
   * the patch path won't write srgbClr from it, avoiding baking in theme colors and breaking theme switches */
  colorFollowsTheme?: boolean
  /** rPr has no solidFill at all (color purely inherited); the rebuild path writes no fill.
   * Explicit schemeClr is not included here (materialized as srgbClr on rebuild to keep visuals) */
  colorInherited?: boolean
  /** Superscript/subscript baseline offset (%) */
  baseline?: number
  /** Hyperlink target, rId resolved at parse: external url, or "slide:N" (0-based) for in-doc jumps */
  hyperlink?: string
  /** Run-level <a:hlinkClick> original r:id/action/tooltip (written back verbatim; absent on a run
   * whose hyperlink was set this session until ensureRunLinkRels allocates the relationship) */
  hyperlinkRId?: string
  hyperlinkAction?: string
  hyperlinkTooltip?: string
  /** underline=true came from hlink styling, not an explicit u attr; write-back never emits u from it */
  underlineImplicit?: boolean
  /** Dynamic field <a:fld type="…"> (slidenum/datetime…); text is the cached value */
  field?: string
  /** Text outline <a:rPr><a:ln> (common in WordArt); width in EMU */
  outline?: { color: ResolvedColor; widthEmu: number }
  /** Run-level outer shadow (<a:rPr>/defRPr <a:effectLst><a:outerShdw>) */
  shadow?: ShadowEffect
}

export type TextAlign = 'left' | 'center' | 'right' | 'justify'

export interface Paragraph {
  runs: TextRun[]
  align?: TextAlign
  /** Indent level (bullet level) */
  level?: number
  /** Line spacing (%, 100 = single) or absolute (pt, via lineExact) */
  lineHeight?: number
  lineExact?: number
  /** Space before/after (pt, spcPts) or as a percentage of single line height (spcPct, 100 = one line) */
  spaceBefore?: number
  spaceAfter?: number
  spaceBeforePct?: number
  spaceAfterPct?: number
  bullet?: {
    type: 'none' | 'char' | 'number'
    char?: string
    color?: ResolvedColor
    /** <a:buFont> typeface (symbol fonts like Wingdings) */
    font?: string
    /** <a:buSzPct> (%, 100 = same size as text) */
    sizePct?: number
    /** <a:buAutoNum type> (arabicPeriod/romanLcParen…) */
    numType?: string
    /** <a:buAutoNum startAt>: first number of the sequence (default 1) */
    startAt?: number
  }
  /** Paragraph left indent marL (EMU) */
  marL?: number
  /** First-line indent (EMU, negative = hanging indent, common with bullets) */
  indent?: number
  /**
   * Which paragraph properties come from an explicit <a:pPr> (rather than display
   * values inherited from lstStyle/placeholder/master). The rebuild path writes
   * only explicit items to avoid baking in inheritance; paragraphs produced by
   * parse always have this field, newly created elements may omit it (omitted =
   * all treated as explicit). Set the matching bit when the user edits a property.
   */
  pPrExplicit?: {
    align?: boolean
    lnSpc?: boolean
    spcBef?: boolean
    spcAft?: boolean
    bullet?: boolean
    marL?: boolean
    indent?: boolean
  }
}

/** Text-box body-level properties */
export interface TextBody {
  paragraphs: Paragraph[]
  /** Vertical alignment */
  anchor?: 'top' | 'middle' | 'bottom'
  /** Insets (EMU): left/top/right/bottom */
  insets?: { l: number; t: number; r: number; b: number }
  /** Autofit: none | shrink font to fit | resize box */
  autofit?: 'none' | 'shrink' | 'resize'
  /** <a:normAutofit fontScale>: font-shrink ratio precomputed by PowerPoint (0-1), used directly for rendering */
  fontScale?: number
  /** <a:normAutofit lnSpcReduction>: line-spacing reduction ratio (0-1, at most 0.2) */
  lnSpcReduction?: number
  wrap?: boolean
  /** <a:bodyPr vert>: vertical text (Japanese tategaki etc.). Read-only display — write-back keeps original bodyPr bytes */
  vert?: 'eaVert' | 'vert' | 'vert270' | 'wordArtVert'
  /** <a:bodyPr numCol>: body text flows across N columns (fill one, then the next) */
  numCol?: number
  /** <a:bodyPr spcCol>: gap between columns (EMU) */
  spcCol?: number
}

// ── Elements ───────────────────────────────────────────────────────────

export type ElementType =
  | 'text' // text box (may be text-only, no fill)
  | 'shape' // autoshape (may carry text)
  | 'picture' // picture
  | 'group' // group
  | 'table' // table (a:tbl inside a graphicFrame; read-only render, saved as original bytes)
  | 'chart' // chart (graphicFrame referencing a chart part; read-only render, saved as original bytes)
  | 'passthrough' // protected block (smartart/ole/connector/animation anchor etc., display-only for now)

/** Dual anchor — the core of byte-level fidelity */
export interface ByteAnchor {
  /** Index of this element among the top-level elements of its slideN.xml (0-based) */
  spIndex: number
  /** Original XML slice (passed through verbatim on save when untouched) */
  originalXml: string
  /** Range [start, end) of the original XML within the slideN.xml string */
  range: [number, number]
  /** Non-shape spTree bytes between this shape and the next (mc:AlternateContent etc.), replayed verbatim on rebuild */
  gapAfter?: string
}

/** Paragraph-property dirty flags: which pPr groups need a surgical patch (set by element-level paragraph formatting ops). */
export interface PPrDirty {
  bullet?: boolean
  /** Indent level lvl (Tab/⇧Tab multi-level lists) */
  level?: boolean
  lnSpc?: boolean
  spcBef?: boolean
  spcAft?: boolean
  align?: boolean
  /** marL + indent as a pair (bullet indent linkage) */
  indents?: boolean
  /** Restrict the patch to these paragraph indices; absent = all paragraphs */
  paraIndices?: number[]
}

interface ElementBase {
  id: string
  type: ElementType
  anchor: ByteAnchor
  transform: Transform
  /** Editor dirty flag: when true, save must regenerate OOXML; otherwise originalXml passes through */
  dirty?: boolean
  /** Geometry dirty flag: when true, save patches <a:xfrm> in place (orthogonal to text patches) */
  dirtyTransform?: boolean
  /** Fill dirty flag: when true, save patches the spPr fill node in place */
  dirtyFill?: boolean
  /** Stroke dirty flag: when true, save patches the spPr <a:ln> in place */
  dirtyStroke?: boolean
  /** Picture-crop dirty flag: when true, save patches p:blipFill's <a:srcRect> in place */
  dirtySrcRect?: boolean
  /** Paragraph-property dirty flags: on save, surgically patch pPr per <a:p> (run bytes untouched) */
  dirtyPPr?: PPrDirty
  /** Placeholder type (title/body/…), located via layout/master inheritance */
  placeholder?: string
  name?: string
  /**
   * <p:cNvPr descr="…">: editor-owned metadata payload (e.g. vector points of
   * freehand ink), written back verbatim with originalXml on save so the editable
   * state can be restored on reopen.
   */
  descr?: string
  /**
   * <p:cNvPr id>: matching key between group children and slices of the group's
   * originalXml (for in-group editing). parseGroup iterates children grouped by
   * tag, which differs from document order, so index alignment is not possible.
   */
  nvId?: string
  /** Connector attachment <a:stCxn>/<a:endCxn>: target shape cNvPr id + connection point index (for move-following) */
  connection?: {
    start?: { id: number; idx: number }
    end?: { id: number; idx: number }
  }
}

/**
 * <a:custGeom> custom geometry (read-only render; save passes originalXml through,
 * never written back). Each field is a normalized SVG path (M/L/C/Q/Z, coords 0..1
 * relative to the w/h declared by <a:path>; arcTo converted to cubic); the render
 * layer scales it to the element frame.
 */
export interface CustomGeometry {
  /** Main fill+stroke path */
  path?: string
  /** stroke="0" subpath: fill only (e.g. the pie sector of arc shapes) */
  fillPath?: string
  /** fill="none" subpath: stroke only (e.g. guide arcs) */
  strokePath?: string
}

export interface TextElement extends ElementBase {
  type: 'text' | 'shape'
  /** Shape's preset geometry (rect/ellipse/roundRect/…); absent for text */
  presetGeometry?: string
  /** Preset geometry adjust values <a:avLst>: e.g. roundRect's adj (1/1000 %, 50000 = pill) */
  adjust?: Record<string, number>
  /** Custom geometry (mutually exclusive with presetGeometry) */
  customGeometry?: CustomGeometry
  fill?: Fill
  /** <p:sp useBgFill="1">: painted with the slide's effective background fill (fill is only a fallback) */
  useBgFill?: boolean
  /** <a:effectLst><a:fillOverlay>: second fill composited over the base (PowerPoint blends
   *  with the record's blend mode; the renderer approximates every mode as multiply) */
  fillOverlay?: Fill
  stroke?: Stroke
  shadow?: ShadowEffect
  glow?: GlowEffect
  text?: TextBody
}

export interface PictureElement extends ElementBase {
  type: 'picture'
  /** Picture media path inside the zip (word/ppt/media/imageN.ext) */
  mediaRef: string
  /** Decoded dataURL (for rendering, may be lazy-loaded) */
  dataUrl?: string
  /** Source crop <a:srcRect>: fraction cropped from each edge (0..1) */
  srcRect?: { l: number; t: number; r: number; b: number }
  /** Whole-image opacity <a:blip><a:alphaModFix amt> (0..1, 1 = opaque; default 1) */
  opacity?: number
  /** Soft edges <a:softEdge rad> (EMU feather radius) */
  softEdge?: number
  /**
   * Audio/video (a:videoFile/a:audioFile under p:nvPr): blipFill is the poster
   * frame; target is the media file's zip path or an external URL (external).
   */
  media?: { kind: 'video' | 'audio'; target?: string; external?: boolean }
  /** Picture outline geometry <a:prstGeom> (ellipse avatars/rounded-corner frames etc. from picture styles; rect omitted) */
  presetGeometry?: string
  adjust?: Record<string, number>
  /** Shape fill from the pic's own spPr, drawn as a backdrop behind the image */
  fill?: Fill
  /** <a:blip><a:duotone> on the picture blip */
  duotone?: [string, string]
  /** <a:blip><a:clrChange> on the picture blip */
  clrChange?: { from: string; to: string }
  stroke?: Stroke
  shadow?: ShadowEffect
  glow?: GlowEffect
}

export interface GroupElement extends ElementBase {
  type: 'group'
  children: SlideElement[]
  /** Child coordinate system (<a:chOff>/<a:chExt>), used to map child coords to the parent */
  childOffset?: EmuRect
}

export interface PassthroughElement extends ElementBase {
  type: 'passthrough'
  /** Protection reason category, for the UI placeholder chip */
  kind: 'chart' | 'smartart' | 'table' | 'ole' | 'connector' | 'media' | 'unknown'
  /**
   * SmartArt read-only preview: shapes parsed from the pptx's built-in prerendered
   * drawing part (diagrams/drawingN.xml), coordinate origin (0,0), size ≈ transform
   * ext (the diagram canvas). Render-only; save still uses the anchor's original bytes.
   */
  previewShapes?: SlideElement[]
  /** OLE read-only preview: the preview picture embedded in the graphicFrame (stretched to fill the frame when rendering). */
  previewPicture?: PictureElement
  /** Render nothing (no placeholder chip): unparseable mc:AlternateContent kept only for byte fidelity */
  noChip?: boolean
}

// ── Table (p:graphicFrame → a:tbl) ───────────────────────────────────

/** Cell borders on all four edges (no border by default; table-style inheritance not yet supported). */
export interface TableCellBorders {
  l?: Stroke
  r?: Stroke
  t?: Stroke
  b?: Stroke
}

export interface TableCell {
  /** Cell text (reuses the TextBody model; anchor/insets already overridden per tcPr) */
  text?: TextBody
  /** Cell fill (explicit tcPr fill; table-style inheritance not yet supported) */
  fill?: Fill
  borders?: TableCellBorders
  /** Horizontal merge span in columns (gridSpan, default 1) */
  gridSpan?: number
  /** Vertical merge span in rows (rowSpan, default 1) */
  rowSpan?: number
  /** Placeholder cell covered by a merge (hMerge/vMerge), not rendered */
  merged?: boolean
}

export interface TableElement extends ElementBase {
  type: 'table'
  /** Column widths (EMU, from a:tblGrid/a:gridCol) */
  colWidths: number[]
  /** Row heights (EMU, from a:tr@h) */
  rowHeights: number[]
  /** rows[r][c], aligned with rowHeights/colWidths */
  rows: TableCell[][]
  /** tblPr's header-row/banded-rows toggles (echoed in the Ribbon's "Table Design") */
  styleFlags?: { firstRow: boolean; bandRow: boolean }
  /** Table-style <a:tblBg>: drawn under the cells (alpha band fills composite over it) */
  bgFill?: Fill
}

// ── Chart (p:graphicFrame → c:chart reference) ───────────────────────

export interface ChartElement extends ElementBase {
  type: 'chart'
  /** Parsed chart data model (see chart.ts) */
  chart: import('./chart').ChartModel
}

export type SlideElement =
  TextElement | PictureElement | GroupElement | TableElement | ChartElement | PassthroughElement

// ── Slide / presentation ───────────────────────────────────────────────

export interface Slide {
  /** Path inside the zip, e.g. ppt/slides/slide1.xml */
  path: string
  /** Full original slideN.xml string (the base for save patches) */
  originalXml: string
  /** Body prefix / suffix (bytes outside the <p:spTree> content, kept verbatim on save) */
  bodyPrefix: string
  bodySuffix: string
  elements: SlideElement[]
  /** Structure dirty flag: set after element add/remove; save rebuilds the whole slide (complements element-level dirty) */
  structureDirty?: boolean
  /** Inheritance sources (for rendering, never written back) */
  layoutPath?: string
  masterPath?: string
  /** Background (inheritance resolved) */
  background?: Fill
  /** The slide carries its own <p:bg> override (false/absent = inherited from layout/master) */
  bgOwn?: boolean
  /** <p:sld showMasterSp="0">: master/layout background graphics hidden on this slide */
  masterSpHidden?: boolean
  /**
   * master/layout decoration layer (read-only render, never written back):
   * non-placeholder concrete shapes on the master (logos/color bars) + enabled
   * footer/slide-number/date placeholders. Drawn below elements, above background.
   */
  decorations?: SlideElement[]
}

export interface SlideSize {
  cx: number // EMU
  cy: number // EMU
}

export interface SlideDeck {
  /** Bytes of every zip entry of the original pptx (untouched entries written back byte-for-byte on save) */
  // The actual byte storage is held by zip.ts's PackageArchive; only parse results live here
  slides: Slide[]
  size: SlideSize
  /** Original SHA-256 (verified before save to ensure the file was not changed externally) */
  originalHash: string
}
