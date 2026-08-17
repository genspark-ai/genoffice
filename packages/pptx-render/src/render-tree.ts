/**
 * RenderTree — a data-driven abstract draw list (the render layer's core asset).
 *
 * Design intent (plan A): parse the Slide element tree into a pure data structure
 * of "converted pixel geometry + resolved styles + laid-out text glyph boxes".
 * Fidelity verification locks onto this layer (coordinates/metrics/wrapping are
 * unit-testable without a canvas); Konva is just a thin adapter that draws the
 * RenderTree.
 *
 * All geometry units = px (absolute target-canvas coordinates), with viewport
 * scale and nested group offsets already applied.
 */
import type { Fill, Stroke } from '@genoffice/pptx-engine'
import type { PlacedBox } from './coords'

export type RenderNodeType =
  | 'shape' // vector shape (may contain text)
  | 'picture' // image
  | 'text' // plain text box
  | 'group' // group (only carries children; its geometry is for nesting)
  | 'table' // table (cells positioned relative to the table's top-left)
  | 'chart' // chart (draw primitives positioned relative to the chart box's top-left)
  | 'placeholder-chip' // passthrough placeholder chip

export interface RenderNodeBase {
  id: string
  type: RenderNodeType
  box: PlacedBox
  /** Source Slide element id, used by the edit layer to locate write-backs */
  sourceId: string
  /** master/layout decoration node: read-only display, not selectable/draggable/snappable */
  decoration?: boolean
  /**
   * Full-page background-like element (bottom-of-z-order full-page fill/picture,
   * no text): still selectable by click, but not draggable, excluded from marquee
   * results/snapping, and marquee drags may start on top of it.
   */
  background?: boolean
}

/** Resolved render fill (RGBA / gradient stops / image dataUrl). */
export type RenderFill =
  | { kind: 'none' }
  | { kind: 'solid'; color: string }
  | {
      kind: 'gradient'
      stops: Array<{ pos: number; color: string }>
      angleDeg: number
      radial?: boolean
      /** Radial focus center as width/height fractions (from <a:fillToRect>; default 0.5/0.5) */
      center?: { x: number; y: number }
    }
  | {
      kind: 'image'
      dataUrl?: string
      mode: 'stretch' | 'tile'
      /** Translucent picture fill (alphaModFix, 0-1) */
      alpha?: number
      /** Image maps into this inset subrect of the shape (stretch fillRect, fractions) */
      fillRect?: { l: number; t: number; r: number; b: number }
      /** [dark, light] duotone colors mapped over image luminance */
      duotone?: [string, string]
      /** clrChange: pixels matching `from` become `to` (#RRGGBB or #RRGGBBAA) */
      clrChange?: { from: string; to: string }
      /** Tile grid: scale in px-per-image-px, anchor offsets in px, and the algn anchor */
      tile?: { scaleX: number; scaleY: number; txPx: number; tyPx: number; algn: string }
    }

export interface RenderStroke {
  color: string
  widthPx: number
  /** Line width (pt, viewport-independent; for property panel display/editing) */
  widthPt: number
  dash?: number[]
  /** OOXML prstDash preset name (for property panel display/editing) */
  dashPreset?: string
}

/** Outer shadow converted to px. */
/** Glow (approximated in Konva as a shadow without offset) */
export interface RenderGlow {
  color: string
  blurPx: number
}

export interface RenderShadow {
  color: string
  blurPx: number
  offsetX: number
  offsetY: number
}

/** A laid-out text glyph block (one contiguous same-format span within a line). */
export interface GlyphRun {
  text: string
  /** px coordinate relative to the text box top-left (baseline left endpoint) */
  x: number
  /** Baseline y (px relative to the text box top-left) */
  baselineY: number
  fontFamily: string
  /** The model run's raw font name (absent when the run has no explicit font — fontFamily is then
   * a layout default or missing-font substitution product; editors must commit this, never fontFamily,
   * or the display fallback gets silently written into the file as an explicit font). */
  srcFontFamily?: string
  fontSizePx: number
  color: string
  bold: boolean
  italic: boolean
  underline: boolean
  /** Strikethrough (<a:rPr strike>) */
  strike?: boolean
  /** Text highlight color (<a:rPr><a:highlight>, drawn as a background behind the run) */
  highlight?: string
  /** Raw super/subscript baseline % (positive = superscript; for editor round-trips, the shift is baked into baselineY) */
  baselinePct?: number
  /** Run width (px, including metrics and letter spacing) */
  widthPx: number
  /** Logical token order before bidi visual reordering; editable DOM uses this to preserve source text order. */
  logicalOrder?: number
  /** Letter spacing (px, appended after each char, may be negative; maps to <a:rPr spc>, fed to canvas letterSpacing by the renderer) */
  letterSpacingPx?: number
  /** Text outline (<a:rPr><a:ln>, commonly used by WordArt) */
  outline?: { color: string; widthPx: number }
  /** Run outer shadow (px), drawn via canvas shadow props */
  shadow?: { color: string; blurPx: number; offsetX: number; offsetY: number }
  /** Extra per-char spacing spread in by justify alignment (px); draw-only, the editor ignores it and doesn't store it */
  justifyExtraPx?: number
  /** Super/subscript baseline shift (px, positive = up; <a:rPr baseline>), already baked into baselineY */
  baselineShiftPx?: number
  /** Latin word in vertical text: drawn rotated 90° clockwise (x/baselineY is the rotation anchor) */
  rotate90?: boolean
  /** Bullet glyph (non-body content injected by layout; text editors should skip it) */
  isBullet?: boolean
  /** RTL direction-level run (Arabic/Hebrew): the renderer must set canvas direction=rtl so punctuation/neutral chars land on the far side */
  rtl?: boolean
  /** Source model run index (into Paragraph.runs); the editor uses it to merge fragments and trace back original formatting */
  srcRunIdx?: number
  /** Run hyperlink ("slide:N" or url, from TextRun.hyperlink); editor round-trips it, follow actions re-resolve live */
  link?: string
  /** The run font's real ascent (px); used by the renderer for baseline→top conversion (falls back to 0.8em) */
  ascentPx?: number
}

/** A laid-out line of text. */
export interface TextLine {
  runs: GlyphRun[]
  /** Line top y (px relative to the text box top-left) */
  top: number
  height: number
  /** Legacy: line advance when it exceeded height (external leading). The 1.2em
   *  PowerPoint line model folds everything into height; kept for stored decks. */
  advance?: number
  /** This line starts a model paragraph (false/absent = auto-wrap continuation of the previous one; the editor splits paragraphs on it) */
  paraStart?: boolean
  /** Trailing whitespace swallowed when wrapping (the editor re-adds a space when joining lines; hard breaks/CJK wrapping don't set it) */
  trailingSpace?: boolean
  /** The line ends with an <a:br/> soft break; value = the sentinel run's model index (the editor round-trips soft breaks with it) */
  softBreakAfter?: number
  /** Paragraph horizontal alignment (editor display) */
  align?: 'left' | 'center' | 'right' | 'justify'
  /** Paragraph indent level (editor Tab multi-level list display) */
  level?: number
  /** Paragraph left margin in px (editor display: body text starts here, not at the inset edge) */
  marLPx?: number
  /** First-line indent in px (editor display; only applies when the paragraph has no bullet) */
  indentPx?: number
  /** Baseline offset minus the line's ascent, signed (canvas baseline = top + leadAbove + ascent;
   *  positive when the line box is taller than the glyphs, negative when shorter — the editor overlay compensates) */
  leadAbove?: number
}

export interface RenderTextLayout {
  lines: TextLine[]
  /** Insets (px) */
  insets: { l: number; t: number; r: number; b: number }
  /** Vertical alignment */
  anchor: 'top' | 'middle' | 'bottom'
  /** Font scale actually applied by autofit (<1 when shrinking) */
  fontScale: number
  /** Line-spacing reduction actually applied by autofit (0-0.2, tied to the fontScale step) */
  lnSpcReduction?: number
  /** Total content height after layout (px), used for vertical alignment positioning */
  contentHeight: number
  /** Ink bottom (last baseline + descent, px): PowerPoint's basis for auto table row heights */
  inkBottom?: number
  /** bodyPr wrap (false = no wrapping, overflows the box; the editor also doesn't wrap) */
  wrap: boolean
  /** bodyPr vert: vertical column layout (lines = columns, right→left); vert/vert270/wordArtVert degrade to eaVert */
  vert?: 'eaVert' | 'vert' | 'vert270' | 'wordArtVert'
}

/** Connector/line endpoint arrow description (for rendering, sizes converted to px). */
export interface ArrowEndRender {
  /** Arrow type: triangle/stealth/diamond/oval/arrow (line arrow) */
  type: 'arrow' | 'triangle' | 'stealth' | 'diamond' | 'oval'
  /** Arrow width (px), the lateral dimension */
  widthPx: number
  /** Arrow length/depth (px), the longitudinal dimension */
  lengthPx: number
}

export interface ShapeRenderNode extends RenderNodeBase {
  type: 'shape' | 'text'
  /** Placeholder type (title/ctrTitle/subTitle/body/…); empty placeholders draw hint text on the canvas */
  placeholder?: string
  presetGeometry?: string
  /** Exact corner radius for roundRect-style geometry (px, computed from avLst adj; 50% of the min side = pill) */
  cornerRadiusPx?: number
  /** Point list of closed-polygon preset geometry (triangle/diamond/arrow…) (local px, drawn closed) */
  polygonPoints?: number[]
  /** SVG path (local px, d format): custGeom / arc-style presets, fill + stroke */
  pathData?: string
  /** Fill-only subpath (OOXML path stroke="0", e.g. an arc's pie sector) */
  fillPathData?: string
  /** Stroke-only subpath (OOXML path fill="none", e.g. arcs, brackets) */
  strokePathData?: string
  /** Connector/straight line: polyline points (local px, flips baked in) + head/tail arrow descriptions */
  line?: {
    points: number[]
    /** Curve control points (bezier, 4 points per segment [cx1,cy1,cx2,cy2,x,y], relative to the polyline start) */
    bezier?: number[]
    headEnd?: ArrowEndRender
    tailEnd?: ArrowEndRender
  }
  fill: RenderFill
  /** <a:fillOverlay>: second fill drawn over the base with multiply blending */
  fillOverlay?: RenderFill
  stroke?: RenderStroke
  shadow?: RenderShadow
  glow?: RenderGlow
  text?: RenderTextLayout
}

export interface PictureRenderNode extends RenderNodeBase {
  type: 'picture'
  dataUrl?: string
  /** Opaque backdrop behind the image (OLE previews render on a white canvas; metafiles are often transparent) */
  bgColor?: string
  /** Shape fill from the pic's own spPr, always drawn behind the (possibly translucent) image */
  fill?: RenderFill
  /** [dark, light] duotone colors applied to the picture pixels */
  duotone?: [string, string]
  /** clrChange applied to the picture pixels before duotone */
  clrChange?: { from: string; to: string }
  /** Picture shape-geometry clip (picture styles): three channels matching shape geometry; clip when any is set */
  clip?: { cornerRadiusPx?: number; polygonPoints?: number[]; pathData?: string }
  /** Source image crop ratios (0..1, how much each side is cropped) */
  srcRect?: { l: number; t: number; r: number; b: number }
  /** Whole-image opacity (0..1; default 1) */
  opacity?: number
  /** Soft-edge feather radius (px) */
  softEdgePx?: number
  /** Audio/video (the image is the poster frame): the render layer overlays a play/speaker badge */
  media?: 'video' | 'audio'
  stroke?: RenderStroke
  shadow?: RenderShadow
  glow?: RenderGlow
  /** Source element cNvPr name (the edit layer identifies its own elements, e.g. freehand ink) */
  name?: string
  /** Source element cNvPr descr payload (freehand ink vector points etc.) */
  descr?: string
}

export interface GroupRenderNode extends RenderNodeBase {
  type: 'group'
  /** children boxes are in group-local coords (relative to the group top-left); ext/chExt scaling is baked into the geometry */
  children: RenderNode[]
  /**
   * ext/chExt scaling (group box size / child coordinate system size). Already baked into
   * children boxes at build time; the render container must **not** apply it again. Kept only
   * for the edit pipeline to convert between local px and child EMU.
   */
  childScaleX?: number
  childScaleY?: number
}

export interface ChipRenderNode extends RenderNodeBase {
  type: 'placeholder-chip'
  kind: string // chart/table/smartart/...
  label: string
}

/** A positioned table cell (coordinates relative to the table top-left, px; merged placeholder cells aren't emitted). */
export interface TableCellRender {
  x: number
  y: number
  w: number
  h: number
  /** Model coordinates (rows[row][col], col is the tc index, not the logical column), for edit write-backs */
  row: number
  col: number
  /** Merge anchor span (>1 means splittable; default 1) */
  gridSpan?: number
  rowSpan?: number
  fill: RenderFill
  /** Border lines on the four sides (default none) */
  borders?: { l?: RenderStroke; r?: RenderStroke; t?: RenderStroke; b?: RenderStroke }
  text?: RenderTextLayout
}

export interface TableRenderNode extends RenderNodeBase {
  type: 'table'
  cells: TableCellRender[]
  /** Table-style <a:tblBg>: drawn under the cells (alpha band fills composite over it) */
  bgFill?: TableCellRender['fill']
  /** Grid line offsets relative to the box (px): gridX has nCols+1 entries, gridY nRows+1 */
  gridX: number[]
  gridY: number[]
  /** tblPr header-row / banded-row toggles (for the Ribbon "Table Design" display) */
  styleFlags?: { firstRow: boolean; bandRow: boolean }
}

// ── Charts (precomputed draw primitives, coordinates relative to the chart box top-left, px) ──

export interface ChartLabel {
  text: string
  /** Text top-left */
  x: number
  y: number
  fontSizePx: number
  color: string
  bold?: boolean
  /** Rotation angle (e.g. -90 for a value-axis title) */
  rotationDeg?: number
}

/** Current chart style (for the Ribbon "Chart Design" display; kind aligns with EditChartOp.kind). */
export interface ChartStyleInfo {
  kind:
    | 'bar'
    | 'bar3D'
    | 'barStacked'
    | 'line'
    | 'area'
    | 'pie'
    | 'pie3D'
    | 'doughnut'
    | 'scatter'
    | 'radar'
    | 'comboBarLine'
    | 'unknown'
  legendPos: 'b' | 't' | 'l' | 'r' | 'none'
  dataLabels: boolean
  gridlines: boolean
  title?: string
  catAxisTitle?: string
  valAxisTitle?: string
  gapWidthPct?: number
}

export interface ChartRenderNode extends RenderNodeBase {
  type: 'chart'
  /** Inserted by this app (cNvPr descr="aislides-chart"): chart editing enabled; passthrough charts lack this flag */
  appCreated?: boolean
  styleInfo?: ChartStyleInfo
  /** Whole-chart background (chartSpace spPr, e.g. picture fill) drawn under all primitives */
  bgFill?: RenderFill
  /** Whole-chart frame border (chartSpace spPr ln) drawn over all primitives */
  border?: { color: string; widthPx: number }
  /** Plot-area fill/border rectangle, drawn under the gridlines */
  plotRect?: {
    x: number
    y: number
    w: number
    h: number
    fill?: RenderFill
    borderColor?: string
    borderWidthPx?: number
  }
  /** Gridlines / axis lines */
  gridLines: Array<{
    x1: number
    y1: number
    x2: number
    y2: number
    color: string
    dash?: number[]
    widthPx?: number
  }>
  axisLines: Array<{
    x1: number
    y1: number
    x2: number
    y2: number
    color: string
    widthPx: number
  }>
  /** Tick / category / legend / axis-title text */
  labels: ChartLabel[]
  /** Bars */
  bars: Array<{ x: number; y: number; w: number; h: number; color: string }>
  /** Polylines (points is flat [x0,y0,x1,y1,...]); closed+fill for filled radar charts etc. */
  polylines: Array<{
    points: number[]
    color: string
    widthPx: number
    smooth?: boolean
    closed?: boolean
    fill?: string
    dash?: number[]
  }>
  /** Data-point markers (circles) */
  markers: Array<{ x: number; y: number; r: number; color: string }>
  /** Legend swatches */
  swatches: Array<{ x: number; y: number; w: number; h: number; color: string }>
  /** Freeform filled paths (SVG data), painter's order — pseudo-3D pie rims / bar extrusion faces */
  paths?: Array<{ d: string; fill: string; stroke?: string; dy?: number }>
  /** Pie/doughnut wedges (angles: 12 o'clock = -90°, clockwise, Konva Arc semantics) */
  wedges?: Array<{
    cx: number
    cy: number
    outerR: number
    innerR: number
    startDeg: number
    sweepDeg: number
    color: string
  }>
}

export type RenderNode =
  | ShapeRenderNode
  | PictureRenderNode
  | GroupRenderNode
  | TableRenderNode
  | ChartRenderNode
  | ChipRenderNode

export interface RenderSlide {
  widthPx: number
  heightPx: number
  /** Viewport scale (fitWidthPx / slide baseline px width) — geometry coordinates already include it */
  scale: number
  background: RenderFill
  /** The slide carries its own <p:bg> override (enables "reset background") */
  bgOwn?: boolean
  /** <p:sld showMasterSp="0">: master/layout background graphics hidden on this slide */
  bgGraphicsHidden?: boolean
  nodes: RenderNode[]
  /** Hidden slide (<p:sld show="0">): thumbnails get a badge, skipped during presentation */
  hidden?: boolean
}

// Convenience type re-exports (for internal render logic)
export type { Fill, Stroke }
