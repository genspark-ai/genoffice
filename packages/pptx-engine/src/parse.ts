/**
 * Parse one slide → Slide element tree.
 *
 * Semantic parsing uses fast-xml-parser; byte-fidelity anchors come from scanSlide
 * (one-to-one in top-level shape order). Phase 1 supports: text boxes / pictures /
 * simple shapes; everything else → passthrough.
 */
import { XMLParser } from 'fast-xml-parser'
import { scanSlide, type SpElement } from './scan'
import { tableRowGridCols } from './table-grid'
import { type Theme, resolveFontRef } from './theme'
import { resolveColorNode as resolveColorNodeShared } from './color'
import {
  resolvePlaceholderTransform,
  parseLstStyleLevels,
  placeholderStyleChain,
  mergeTextStyleChain,
  type PlaceholderMap,
  type MasterTextStyles,
  type TextStyleLevels,
  type LevelTextStyle,
} from './placeholder'
import type {
  Slide,
  SlideElement,
  TextElement,
  PictureElement,
  PassthroughElement,
  GroupElement,
  Transform,
  TextBody,
  Paragraph,
  TextRun,
  Fill,
  Stroke,
  ArrowEnd,
  ArrowEndSize,
  ShadowEffect,
  ByteAnchor,
  TableElement,
  TableCell,
  TableCellBorders,
  ChartElement,
} from './types'
import { parseChartXml } from './chart'
import { parseCustGeom } from './custgeom'
import {
  resolveTableStyle,
  cellPartStyle,
  cellStyleBorders,
  type TablePartStyle,
  type TableStyleFlags,
} from './table-style'

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  // Text fidelity: no trim (leading/trailing spaces in runs matter, e.g. "bold word " + following text),
  // no numeric coercion of tag values (otherwise <a:t>2026</a:t> becomes a number and downstream string reads lose characters)
  trimValues: false,
  parseTagValue: false,
  // Order preservation is not the point (semantic tree); keep array structure for multiple runs/paragraphs
  isArray: (name) =>
    [
      'a:p',
      'a:r',
      'a:br',
      'a:fld',
      'p:sp',
      'p:pic',
      'p:graphicFrame',
      'p:grpSp',
      'p:cxnSp',
      'a:tr',
      'a:tc',
      'a:gridCol',
    ].includes(name),
  // spTree children nested in groups also need arrays (covered above)
})

const EMU_PER_PT = 12700

export interface ParseContext {
  theme?: Theme
  /** Placeholder color for resolving style ref templates (value substituted for schemeClr val="phClr") */
  phClr?: string
  /** Media rId → zip path, for picture parsing */
  mediaRels?: Map<string, string>
  /** Hyperlink rId → resolved target: external url, or "slide:N" (0-based) for slide jumps */
  hlinkRels?: Map<string, string>
  /** Chart rId → chartN.xml content (chart part referenced by a graphicFrame) */
  chartXmls?: Map<string, string>
  /** Audio/video rId → media zip path or external URL (r:link of videoFile/audioFile) */
  avRels?: Map<string, { target: string; external?: boolean }>
  /** SmartArt: diagramData rId (dgm:relIds@r:dm) → prerendered drawing part content */
  diagramDrawings?: Map<string, string>
  /** Placeholder geometry inheritance table: from the slideLayout (read-only) */
  layoutPlaceholders?: PlaceholderMap
  /** Placeholder geometry inheritance table: from the slideMaster (read-only, fallback when the layout lacks it) */
  masterPlaceholders?: PlaceholderMap
  /** master <p:txStyles> text style defaults (title/body/other families) */
  masterTextStyles?: MasterTextStyles
  /** Full layout XML (read-only, for background inheritance) */
  layoutBg?: string
  /** Full master XML (read-only, background inheritance fallback) */
  masterBg?: string
  /** ppt/tableStyles.xml source (table style definitions, read-only) */
  tableStyles?: string
}

export interface SlideParseInput {
  path: string
  slideXml: string
  layoutPath?: string
  masterPath?: string
  ctx: ParseContext
}

let uidCounter = 0
function uid(prefix: string): string {
  return `${prefix}_${(uidCounter++).toString(36)}`
}

export function parseSlide(input: SlideParseInput): Slide {
  const { slideXml, path, layoutPath, masterPath, ctx } = input
  const scan = scanSlide(slideXml)

  // Parse each shape's XML fragment with fast-xml-parser (independent parses, naturally aligned with scan order)
  const elements: SlideElement[] = []
  scan.elements.forEach((sp, idx) => {
    const fragXml = slideXml.slice(sp.start, sp.end)
    const anchor: ByteAnchor = {
      spIndex: idx,
      originalXml: fragXml,
      range: [sp.start, sp.end],
      ...(sp.gapAfter ? { gapAfter: sp.gapAfter } : {}),
    }
    const el = parseShapeFragment(sp, fragXml, anchor, ctx)
    if (el) elements.push(el)
  })

  // Background: the slide's own <p:bg> wins, otherwise inherit layout→master (read-only).
  const background =
    parseBackground(slideXml, ctx) ??
    (ctx.layoutBg ? parseBackground(ctx.layoutBg, ctx) : undefined) ??
    (ctx.masterBg ? parseBackground(ctx.masterBg, ctx) : undefined)

  return {
    path,
    originalXml: slideXml,
    bodyPrefix: scan.bodyPrefix,
    bodySuffix: scan.bodySuffix,
    elements,
    layoutPath,
    masterPath,
    ...(background ? { background } : {}),
  }
}

/** Extract the <p:bg> background fill from slide/layout/master XML (read-only). */
function parseBackground(xml: string, ctx: ParseContext): Fill | undefined {
  // Extract only the <p:bg>…</p:bg> fragment and parse it alone, avoiding a whole-slide parse
  const m = /<p:bg\b[\s\S]*?<\/p:bg>/.exec(xml)
  if (!m) return undefined
  let doc: any
  try {
    doc = parser.parse(m[0])
  } catch {
    return undefined
  }
  const bg = doc['p:bg']
  const bgPr = bg?.['p:bgPr']
  if (bgPr) {
    // bgPr directly contains solidFill/gradFill/blipFill/pattFill
    return parseFill(bgPr, ctx)
  }
  // <p:bgRef>: references the theme fillStyleLst + a color (simplified: take its color as solid)
  const bgRef = bg?.['p:bgRef']
  if (bgRef) {
    const color = resolveColorNode(bgRef, ctx)
    if (color) return { type: 'solid', color }
  }
  return undefined
}

function parseShapeFragment(
  sp: SpElement,
  fragXml: string,
  anchor: ByteAnchor,
  ctx: ParseContext,
): SlideElement | null {
  // <a:br/> (in-paragraph soft break) → sentinel run "\n": fast-xml-parser does not
  // preserve order, so the relative position of a:br vs a:r is lost; replacing it
  // with a line-break sentinel lets the layout layer force a break. Only affects the
  // semantic tree; the byte-fidelity side's anchor.originalXml stays the original fragment.
  // <a:fld> (slide number/date) gets the same treatment: it is structurally an a:r,
  // and rewriting the tag (attributes kept, so @_type survives for run.field) keeps
  // fields in document order instead of being appended after all plain runs.
  const semanticXml = fragXml
    .replace(/<a:br\b[^>]*\/>|<a:br\b[\s\S]*?<\/a:br>/g, '<a:r><a:t>\n</a:t></a:r>')
    .replace(/<a:fld\b/g, '<a:r')
    .replace(/<\/a:fld>/g, '</a:r>')
  const doc = parser.parse(semanticXml)
  const node = doc[sp.name] ? (Array.isArray(doc[sp.name]) ? doc[sp.name][0] : doc[sp.name]) : null
  if (!node) return null

  switch (sp.name) {
    case 'p:sp':
      return parseSpShape(node, anchor, ctx)
    case 'p:pic':
      return parsePicture(node, anchor, ctx)
    case 'p:grpSp':
      return parseGroup(node, anchor, ctx)
    case 'p:graphicFrame':
      return graphicFramePassthrough(node, anchor, ctx)
    case 'p:cxnSp':
      return parseConnector(node, anchor, ctx)
    default:
      return passthrough(anchor, 'unknown', node)
  }
}

// ── p:sp (text box / shape) ──────────────────────────────────────────

function parseSpShape(
  node: any,
  anchor: ByteAnchor,
  ctx: ParseContext,
  rawXml?: string,
): TextElement | PassthroughElement {
  const spPr = node['p:spPr'] ?? {}
  const nv = node['p:nvSpPr']
  const ph = nv?.['p:nvPr']?.['p:ph']
  const phType = ph?.['@_type']
  const phIdx = ph?.['@_idx'] != null ? String(ph['@_idx']) : undefined
  const name = nv?.['p:cNvPr']?.['@_name']

  let transform = parseXfrm(spPr['a:xfrm'])
  // Phase 2 fix: when a placeholder omits <a:xfrm>, geometry is backfilled from layout/master inheritance.
  if (ph && !spPr['a:xfrm']) {
    const inherited = resolvePlaceholderTransform(
      ctx.layoutPlaceholders,
      ctx.masterPlaceholders,
      phType,
      phIdx,
    )
    if (inherited) transform = inherited
  }

  const prstGeom = spPr['a:prstGeom']
  const presetGeometry = prstGeom?.['@_prst']
  const adjust = parseAvLst(prstGeom?.['a:avLst'])
  // custGeom needs an order-preserving command stream → parse from the raw bytes (group children get their slice via rawXml)
  const customGeometry =
    spPr['a:custGeom'] != null
      ? parseCustGeom(rawXml || anchor.originalXml, transform.offset.cx, transform.offset.cy)
      : undefined
  let fill = parseFill(spPr, ctx)
  const txBody = node['p:txBody']
  // Text style inheritance chain: placeholders inherit font size/color/font defaults from layout/master
  const phChain = ph
    ? placeholderStyleChain(
        ctx.layoutPlaceholders,
        ctx.masterPlaceholders,
        ctx.masterTextStyles,
        phType,
        phIdx,
      )
    : []
  const text = txBody ? parseTextBody(txBody, ctx, phChain) : undefined

  let stroke = parseStroke(spPr, ctx)
  let shadow = parseShadow(spPr, ctx)
  let glow = parseGlow(spPr, ctx)

  // <p:style> theme style reference fallback: when spPr has no explicit value, take the
  // fmtScheme template by idx (fillStyleLst/lnStyleLst/effectStyleLst) with phClr
  // substituted by the reference color; when the theme lacks the template, fall back to
  // the reference color as solid (shape styles of SmartArt pre-rendered drawings all
  // come from here). The fontRef color is filled into runs without an explicit color.
  const style = node['p:style']
  if (style && typeof style === 'object') {
    if (fill === undefined) {
      const ref = style['a:fillRef']
      const idx = parseInt(String(ref?.['@_idx'] ?? '0'), 10) || 0
      const phClr = resolveColorNode(ref, ctx)
      if (idx > 0) {
        // idx 1..3 -> fillStyleLst; 1001..1003 -> bgFillStyleLst (background style references)
        const tpl =
          idx > 1000 ? ctx.theme?.bgFillStyles?.[idx - 1001] : ctx.theme?.fillStyles?.[idx - 1]
        const tplFill = tpl ? parseFill(tpl, { ...ctx, phClr }) : undefined
        fill = tplFill ?? (phClr ? { type: 'solid', color: phClr } : undefined)
      }
    }
    // explicit <a:ln><a:noFill/> (stroke === null) wins over the lnRef template
    if (stroke === undefined) stroke = styleRefStroke(node, ctx)
    if (!shadow && !glow) {
      const ref = style['a:effectRef']
      const idx = parseInt(String(ref?.['@_idx'] ?? '0'), 10) || 0
      const phClr = resolveColorNode(ref, ctx)
      const es = idx > 0 ? ctx.theme?.effectStyles?.[idx - 1]?.['a:effectStyle'] : undefined
      if (es) {
        const tplCtx = { ...ctx, phClr }
        shadow = parseShadow(es, tplCtx)
        glow = parseGlow(es, tplCtx)
      }
    }
    const fontColor = resolveColorNode(style['a:fontRef'], ctx)
    if (fontColor && text) {
      for (const p of text.paragraphs) {
        for (const r of p.runs) if (!r.color) r.color = fontColor
      }
    }
  }

  const el: TextElement = {
    id: uid('sp'),
    type: txBody && !presetGeometry && !customGeometry ? 'text' : 'shape',
    anchor,
    transform,
    // <p:ph> without a type (content placeholder) defaults to body per ECMA
    placeholder: ph ? (phType ?? 'body') : undefined,
    name,
    presetGeometry,
    ...(adjust ? { adjust } : {}),
    ...(customGeometry ? { customGeometry } : {}),
    fill,
    ...(stroke ? { stroke } : {}),
    ...(shadow ? { shadow } : {}),
    ...(glow ? { glow } : {}),
    text,
  }
  return el
}

/** <p:style> lnRef -> theme lnStyleLst template stroke (phClr substituted by the reference color); falls back to a 1pt stroke in the reference color when the theme lacks the template. */
function styleRefStroke(node: any, ctx: ParseContext): Stroke | undefined {
  const ref = node?.['p:style']?.['a:lnRef']
  const idx = parseInt(String(ref?.['@_idx'] ?? '0'), 10) || 0
  if (idx <= 0) return undefined
  const phClr = resolveColorNode(ref, ctx)
  const tpl = ctx.theme?.lnStyles?.[idx - 1]
  return (
    (tpl ? parseStroke(tpl, { ...ctx, phClr }) : undefined) ??
    (phClr ? { fill: { type: 'solid', color: phClr }, width: 12700 } : undefined)
  )
}

/**
 * <a:ln> stroke: fill (solid/gradient…) + width + dash + cap + arrowheads.
 * Returns undefined when a:ln is absent (nothing specified — callers may
 * inherit from p:style/lnRef) and null for an explicit <a:noFill/> (the
 * author turned the outline off — must NOT be upgraded to a theme stroke).
 */
function parseStroke(
  spPr: any,
  ctx: ParseContext,
  fallbackColor?: string,
): Stroke | null | undefined {
  const ln = spPr?.['a:ln']
  if (!ln || typeof ln !== 'object') return undefined
  if ('a:noFill' in ln) return null
  let fill = parseFill(ln, ctx)
  // <a:ln> with no explicit fill is treated as "no stroke" (a full implementation would inherit the theme lnStyleLst);
  // connectors are the exception: with no explicit fill use the caller's fallback color (a connector without a stroke is invisible)
  if (!fill || fill.type === 'none') {
    if (!fallbackColor) return undefined
    fill = { type: 'solid', color: fallbackColor }
  }
  const capMap: Record<string, Stroke['cap']> = { flat: 'flat', rnd: 'round', sq: 'square' }
  const dash = ln['a:prstDash']?.['@_val']
  const cap = ln['@_cap'] ? capMap[ln['@_cap']] : undefined
  const headEnd = parseArrowEnd(ln['a:headEnd'])
  const tailEnd = parseArrowEnd(ln['a:tailEnd'])
  return {
    fill,
    width: intOr(ln['@_w'], 12700),
    ...(dash ? { dash: String(dash) } : {}),
    ...(cap ? { cap } : {}),
    ...(headEnd ? { headEnd } : {}),
    ...(tailEnd ? { tailEnd } : {}),
  }
}

/** Parse <a:headEnd>/<a:tailEnd> → ArrowEnd (omitted when type=none). */
function parseArrowEnd(node: any): ArrowEnd | undefined {
  if (!node || typeof node !== 'object') return undefined
  const type = String(node['@_type'] ?? 'none') as ArrowEnd['type']
  if (type === 'none') return undefined
  const wRaw = node['@_w']
  const lenRaw = node['@_len']
  const sizeMap: Record<string, ArrowEndSize> = { sm: 'sm', med: 'med', lg: 'lg' }
  return {
    type,
    ...(wRaw ? { w: sizeMap[wRaw] ?? 'med' } : {}),
    ...(lenRaw ? { len: sizeMap[lenRaw] ?? 'med' } : {}),
  }
}

// ── p:cxnSp (connector) ─────────────────────────────────────────────

/**
 * Connectors: line / straightConnector / bentConnector / curvedConnector.
 * Semantically = a stroke-only shape (geometry name + adjust + stroke/arrows);
 * the start/end connections (a:stCxn/endCxn) only affect editor snapping, and
 * rendering just draws by xfrm + flip.
 */
function parseConnector(node: any, anchor: ByteAnchor, ctx: ParseContext): TextElement {
  const spPr = node['p:spPr'] ?? {}
  const nvCxn = node['p:nvCxnSpPr']
  const name = nvCxn?.['p:cNvPr']?.['@_name']
  const prstGeom = spPr['a:prstGeom']
  // Stroke priority: explicit <a:ln> (when it has no fill, complete the color from the lnRef reference color/dk1, keeping arrows and dashes)
  // -> lnRef theme template -> dk1 solid-line fallback (a connector without a stroke is effectively invisible)
  const refStroke = styleRefStroke(node, ctx)
  const fallback =
    (refStroke?.fill.type === 'solid' ? refStroke.fill.color : undefined) ??
    ctx.theme?.colors?.['dk1'] ??
    '#000000'
  const explicitStroke = parseStroke(spPr, ctx, spPr?.['a:ln'] ? fallback : undefined)
  // null = author explicitly disabled the outline; only an *absent* a:ln
  // falls back (so an unstyled connector never turns invisible)
  const stroke =
    explicitStroke === null
      ? undefined
      : (explicitStroke ??
        refStroke ??
        ({ fill: { type: 'solid', color: fallback }, width: 12700 } satisfies Stroke))
  // Attachment <a:stCxn>/<a:endCxn>: target shape cNvPr id + connection point index (for move-following)
  const cxnPr = nvCxn?.['p:cNvCxnSpPr']
  const st = cxnPr?.['a:stCxn']
  const end = cxnPr?.['a:endCxn']
  const cxnRef = (n: any) =>
    n?.['@_id'] != null ? { id: parseInt(n['@_id'], 10), idx: intOr(n['@_idx'], 0) } : undefined
  const connection =
    st || end
      ? {
          ...(cxnRef(st) ? { start: cxnRef(st)! } : {}),
          ...(cxnRef(end) ? { end: cxnRef(end)! } : {}),
        }
      : undefined
  return {
    id: uid('cxn'),
    type: 'shape',
    anchor,
    transform: parseXfrm(spPr['a:xfrm']),
    name,
    presetGeometry: prstGeom?.['@_prst'] ?? 'line',
    ...(parseAvLst(prstGeom?.['a:avLst']) ? { adjust: parseAvLst(prstGeom?.['a:avLst']) } : {}),
    ...(connection ? { connection } : {}),
    fill: { type: 'none' },
    ...(stroke ? { stroke } : {}),
  }
}

/** <a:effectLst><a:outerShdw> outer shadow. */
function parseGlow(spPr: any, ctx: ParseContext): import('./types').GlowEffect | undefined {
  const glow = spPr?.['a:effectLst']?.['a:glow']
  if (!glow || typeof glow !== 'object') return undefined
  const color = resolveColorNode(glow, ctx)
  if (!color) return undefined
  const rad = glow['@_rad'] != null ? parseInt(glow['@_rad'], 10) : 0
  return { color, radius: Number.isFinite(rad) ? rad : 0 }
}

function parseShadow(spPr: any, ctx: ParseContext): ShadowEffect | undefined {
  const shdw = spPr?.['a:effectLst']?.['a:outerShdw']
  if (!shdw || typeof shdw !== 'object') return undefined
  const color = resolveColorNode(shdw, ctx)
  if (!color) return undefined
  return {
    color,
    blurRad: intOr(shdw['@_blurRad'], 0),
    dist: intOr(shdw['@_dist'], 0),
    dirDeg: intOr(shdw['@_dir'], 0) / 60000,
  }
}

/** <a:avLst> adjust values: <a:gd name="adj" fmla="val 50000"/> → { adj: 50000 }. */
function parseAvLst(avLst: any): Record<string, number> | undefined {
  const gdRaw = avLst?.['a:gd']
  if (!gdRaw) return undefined
  const list = Array.isArray(gdRaw) ? gdRaw : [gdRaw]
  const out: Record<string, number> = {}
  for (const gd of list) {
    const name = gd?.['@_name']
    const m = /^val\s+(-?\d+)/.exec(String(gd?.['@_fmla'] ?? ''))
    if (name && m) out[name] = parseInt(m[1]!, 10)
  }
  return Object.keys(out).length ? out : undefined
}

// ── p:grpSp (group) ────────────────────────────────────────

const GROUP_CHILD_TAGS = ['p:sp', 'p:pic', 'p:grpSp', 'p:graphicFrame', 'p:cxnSp'] as const

function parseGroup(
  node: any,
  anchor: ByteAnchor,
  ctx: ParseContext,
  rawXml?: string,
): GroupElement {
  const grpSpPr = node['p:grpSpPr'] ?? {}
  const xfrm = grpSpPr['a:xfrm']
  const transform = parseXfrm(xfrm)
  const name = node['p:nvGrpSpPr']?.['p:cNvPr']?.['@_name']

  // Child coordinate system: <a:chOff>/<a:chExt> (child coords are based on it, mapped to the parent when rendering)
  const chOff = xfrm?.['a:chOff']
  const chExt = xfrm?.['a:chExt']
  const childOffset =
    chOff || chExt
      ? {
          x: chOff ? parseInt(chOff['@_x'], 10) || 0 : 0,
          y: chOff ? parseInt(chOff['@_y'], 10) || 0 : 0,
          cx: chExt ? parseInt(chExt['@_cx'], 10) || 0 : 0,
          cy: chExt ? parseInt(chExt['@_cy'], 10) || 0 : 0,
        }
      : undefined

  // Recursively parse children. Child byte anchors are group-local (only for
  // render/editor positioning; saving still uses the whole group's originalXml:
  // if any child is dirty the whole group regenerates).
  const groupXml = rawXml || anchor.originalXml
  const slices = sliceGroupChildren(groupXml)
  const byTag: Record<string, GroupChildSlice[]> = {}
  for (const s of slices) (byTag[s.name] ??= []).push(s)
  const ordered: Array<{ el: SlideElement; start: number }> = []
  for (const tag of GROUP_CHILD_TAGS) {
    const raw = node[tag]
    if (!raw) continue
    const list = Array.isArray(raw) ? raw : [raw]
    list.forEach((child, i) => {
      const slice = byTag[tag]?.[i]
      const el = parseGroupChild(tag, child, ctx, slice?.xml)
      if (el) ordered.push({ el, start: slice?.start ?? Number.MAX_SAFE_INTEGER })
    })
  }
  // fast-xml-parser batches same-name children; the slice offsets restore document order (z-order)
  ordered.sort((a, b) => a.start - b.start)
  const children = ordered.map((o) => o.el)

  return {
    id: uid('grp'),
    type: 'group',
    anchor,
    transform,
    name,
    children,
    ...(childOffset ? { childOffset } : {}),
  }
}

/** Parse a group child (uses the child node's own bytes as originalXml, only for regeneration positioning). */
function parseGroupChild(
  tag: string,
  child: any,
  ctx: ParseContext,
  rawXml?: string,
): SlideElement | null {
  // Child byte anchor: no independent byte roundtrip inside a group (whole group passes through), so use an empty anchor.
  const childAnchor: ByteAnchor = { spIndex: -1, originalXml: '', range: [0, 0] }
  let el: SlideElement | null
  switch (tag) {
    case 'p:sp':
      el = parseSpShape(child, childAnchor, ctx, rawXml)
      break
    case 'p:pic':
      el = parsePicture(child, childAnchor, ctx)
      break
    case 'p:grpSp':
      el = parseGroup(child, childAnchor, ctx, rawXml)
      break
    case 'p:graphicFrame':
      el = graphicFramePassthrough(child, childAnchor, ctx)
      break
    case 'p:cxnSp':
      el = parseConnector(child, childAnchor, ctx)
      break
    default:
      return null
  }
  const nvId = groupChildNvId(child)
  if (el && nvId != null) el.nvId = nvId
  return el
}

/** Child's <p:cNvPr id> (the nv*Pr container name varies by tag, so try each). */
function groupChildNvId(child: any): string | undefined {
  for (const key of ['p:nvSpPr', 'p:nvPicPr', 'p:nvGrpSpPr', 'p:nvGraphicFramePr', 'p:nvCxnSpPr']) {
    const id = child?.[key]?.['p:cNvPr']?.['@_id']
    if (id != null) return String(id)
  }
  return undefined
}

// Same tag matching style as scan.ts (tolerates '>' inside attribute values)
const GROUP_TAG_RE = /<\/?(?:[^<>"']|"[^"]*"|'[^']*')*>/g
const GROUP_NAME_RE = /^<\/?\s*([A-Za-z_][\w:.-]*)/

interface GroupChildSlice {
  name: string
  xml: string
  start: number
}

/**
 * Group source XML → source fragments of direct child shapes, in document order.
 * Per-tag index order matches fast-xml-parser's same-name arrays; `start` restores
 * cross-tag document order. custGeom also needs source-order command parsing.
 */
function sliceGroupChildren(xml: string): GroupChildSlice[] {
  const out: GroupChildSlice[] = []
  const tags = new Set<string>(GROUP_CHILD_TAGS)
  GROUP_TAG_RE.lastIndex = 0
  let depth = 0
  let start = -1
  let startName = ''
  let m: RegExpExecArray | null
  while ((m = GROUP_TAG_RE.exec(xml))) {
    const tag = m[0]
    if (tag.startsWith('<!--') || tag.startsWith('<![') || tag.startsWith('<?')) continue
    const closing = tag.startsWith('</')
    const self = !closing && tag.endsWith('/>')
    const name = GROUP_NAME_RE.exec(tag)?.[1] ?? ''
    if (closing) {
      depth--
      if (depth === 1 && startName) {
        out.push({ name: startName, xml: xml.slice(start, m.index + tag.length), start })
        startName = ''
      }
    } else if (self) {
      // Self-closing direct children also need a slot, keeping indices aligned with the parsed arrays
      if (depth === 1 && tags.has(name)) out.push({ name, xml: tag, start: m.index })
    } else {
      // depth 0 = the group's own open tag; depth 1 = direct children
      if (depth === 1 && tags.has(name)) {
        start = m.index
        startName = name
      }
      depth++
    }
  }
  return out
}

/** Direct child fragments of a p:grpSp in document order (depth-aware: nested groups stay one slice). */
export function sliceGroupChildXmls(grpXml: string): string[] {
  return sliceGroupChildren(grpXml).map((s) => s.xml)
}

// ── p:pic (picture) ──────────────────────────────────────────────────

/** r:embed of an <a:blip>, falling back to the Office 2016 <asvg:svgBlip> extension.
    SVG-only pictures (e.g. PowerPoint 365 vector logos) can carry a bare <a:blip>
    whose only image reference is the svgBlip inside a:extLst — without this
    fallback such pictures resolve to no media and render as a broken-image box. */
function blipEmbedId(blip: any): string | undefined {
  const direct = blip?.['@_r:embed']
  if (direct) return direct
  const exts = blip?.['a:extLst']?.['a:ext']
  for (const ext of Array.isArray(exts) ? exts : exts ? [exts] : []) {
    for (const [key, value] of Object.entries(ext as Record<string, any>)) {
      if (key === 'svgBlip' || key.endsWith(':svgBlip')) {
        const id = value?.['@_r:embed']
        if (id) return id
      }
    }
  }
  return undefined
}

function parsePicture(node: any, anchor: ByteAnchor, ctx: ParseContext): PictureElement {
  const spPr = node['p:spPr'] ?? {}
  const transform = parseXfrm(spPr['a:xfrm'])
  const blipFill = node['p:blipFill']
  const blip = blipFill?.['a:blip']
  const embedId = blipEmbedId(blip)
  const mediaRef = (embedId && ctx.mediaRels?.get(embedId)) || ''
  const name = node['p:nvPicPr']?.['p:cNvPr']?.['@_name']
  const descr = node['p:nvPicPr']?.['p:cNvPr']?.['@_descr']
  const srcRect = parseSrcRect(blipFill?.['a:srcRect'])
  // picture styles outline geometry (ellipse avatars/rounded-corner frames etc.); rect is the default and not recorded
  const picGeom = spPr['a:prstGeom']?.['@_prst']
  const picAdjust = parseAvLst(spPr['a:prstGeom']?.['a:avLst'])
  const softEdgeRad = spPr['a:effectLst']?.['a:softEdge']?.['@_rad']
  const alphaAmt = blip?.['a:alphaModFix']?.['@_amt']
  const opacity =
    alphaAmt != null ? Math.max(0, Math.min(1, parseInt(alphaAmt, 10) / 100000)) : undefined
  const stroke = parseStroke(spPr, ctx)
  const shadow = parseShadow(spPr, ctx)
  const glow = parseGlow(spPr, ctx)
  // Audio/video: a:videoFile/a:audioFile under p:nvPr; blipFill is the poster frame
  const nvPr = node['p:nvPicPr']?.['p:nvPr']
  const avNode = nvPr?.['a:videoFile'] ?? nvPr?.['a:audioFile']
  let media: PictureElement['media']
  if (avNode !== undefined) {
    const kind = nvPr?.['a:videoFile'] !== undefined ? ('video' as const) : ('audio' as const)
    const link = avNode?.['@_r:link']
    const rel = link ? ctx.avRels?.get(String(link)) : undefined
    media = {
      kind,
      ...(rel ? { target: rel.target, ...(rel.external ? { external: true } : {}) } : {}),
    }
  }
  return {
    id: uid('pic'),
    type: 'picture',
    anchor,
    transform,
    name,
    ...(descr ? { descr } : {}),
    mediaRef,
    ...(srcRect ? { srcRect } : {}),
    ...(picGeom && picGeom !== 'rect'
      ? { presetGeometry: picGeom, ...(picAdjust ? { adjust: picAdjust } : {}) }
      : {}),
    ...(opacity != null && opacity < 1 ? { opacity } : {}),
    ...(softEdgeRad != null ? { softEdge: intOr(softEdgeRad, 0) } : {}),
    ...(media ? { media } : {}),
    ...(stroke ? { stroke } : {}),
    ...(shadow ? { shadow } : {}),
    ...(glow ? { glow } : {}),
  }
}

/** <a:srcRect l/t/r/b> (1/1000 %) → 0..1 fractions; all zero → undefined. */
function parseSrcRect(sr: any): PictureElement['srcRect'] | undefined {
  if (!sr || typeof sr !== 'object') return undefined
  const f = (k: string) => intOr(sr[`@_${k}`], 0) / 100000
  const rect = { l: f('l'), t: f('t'), r: f('r'), b: f('b') }
  if (!rect.l && !rect.t && !rect.r && !rect.b) return undefined
  return rect
}

// ── p:graphicFrame (table / chart / smartart / ole) kind detection ───

function graphicFramePassthrough(node: any, anchor: ByteAnchor, ctx: ParseContext): SlideElement {
  const data = node['a:graphic']?.['a:graphicData']
  const uri: string = data?.['@_uri'] ?? ''
  // Table: semantic parsing (read-only render; save still uses the anchor's original bytes)
  if (uri.includes('/table') && data?.['a:tbl']) {
    const table = parseTable(node, data['a:tbl'], anchor, ctx)
    if (table) return table
  }
  // Chart: read the referenced chart part for semantic parsing (read-only render; save uses original bytes)
  if (uri.includes('/chart')) {
    const rid = data?.['c:chart']?.['@_r:id']
    const chartXml = rid ? ctx.chartXmls?.get(rid) : undefined
    const model = chartXml ? parseChartXml(chartXml, ctx.theme) : null
    if (model) {
      const cNvPr = node['p:nvGraphicFramePr']?.['p:cNvPr']
      const descr: string | undefined = cNvPr?.['@_descr'] || undefined
      return {
        id: uid('chart'),
        type: 'chart',
        anchor,
        transform: parseXfrm(node['p:xfrm']),
        name: cNvPr?.['@_name'],
        ...(descr ? { descr } : {}),
        chart: model,
      } satisfies ChartElement
    }
  }
  let kind: PassthroughElement['kind'] = 'unknown'
  if (uri.includes('/table')) kind = 'table'
  else if (uri.includes('/chart')) kind = 'chart'
  else if (uri.includes('/diagram') || uri.includes('SmartArt')) kind = 'smartart'
  else if (uri.includes('/ole')) kind = 'ole'
  const transform = parseXfrm(node['p:xfrm'])
  const el: PassthroughElement = {
    id: uid('gf'),
    type: 'passthrough',
    anchor,
    transform,
    kind,
  }
  // SmartArt read-only preview: dgm:relIds@r:dm → prerendered drawing part (assembled in index.ts)
  if (kind === 'smartart') {
    const dm = data?.['dgm:relIds']?.['@_r:dm']
    const drawingXml = dm ? ctx.diagramDrawings?.get(String(dm)) : undefined
    if (drawingXml) {
      const shapes = parseDiagramDrawing(drawingXml, ctx)
      if (shapes.length) el.previewShapes = shapes
    }
  }
  // OLE read-only preview: p:pic embedded in graphicData (often inside mc:AlternateContent/mc:Fallback)
  if (kind === 'ole') {
    const pic = findDescendantPic(data)
    if (pic) el.previewPicture = parsePicture(pic, anchor, ctx)
  }
  return el
}

/**
 * SmartArt prerendered drawing part (diagrams/drawingN.xml, dsp namespace) →
 * read-only shapes. dsp:sp/dsp:spPr/dsp:txBody/dsp:style are structurally
 * isomorphic to the p: prefix, so after a prefix swap the p:sp parser is reused
 * directly (shape colors fall back to dsp:style's fillRef/lnRef/fontRef theme
 * references). Coordinate system: the diagram canvas (origin 0,0, size ≈
 * graphicFrame ext).
 */
function parseDiagramDrawing(drawingXml: string, ctx: ParseContext): SlideElement[] {
  const xml = drawingXml.replace(/<(\/?)dsp:/g, '<$1p:')
  let doc: any
  try {
    doc = parser.parse(xml)
  } catch {
    return []
  }
  const spTree = doc['p:drawing']?.['p:spTree']
  if (!spTree) return []
  const spsRaw = spTree['p:sp']
  const sps: any[] = Array.isArray(spsRaw) ? spsRaw : spsRaw ? [spsRaw] : []
  const out: SlideElement[] = []
  for (const sp of sps) {
    // The preview layer has no byte fidelity (never written back); the anchor is a placeholder
    const anchor: ByteAnchor = { spIndex: -1, originalXml: '', range: [0, 0] }
    const el = parseSpShape(sp, anchor, ctx)
    if (el.type !== 'passthrough') out.push(el)
  }
  return out
}

/** Find the first p:pic in the graphicData subtree (piercing wrappers like mc:AlternateContent). */
function findDescendantPic(node: any, depth = 0): any | undefined {
  if (!node || typeof node !== 'object' || depth > 6) return undefined
  const pics = node['p:pic']
  if (Array.isArray(pics) && pics.length) return pics[0]
  for (const [k, v] of Object.entries(node)) {
    if (k.startsWith('@_')) continue
    for (const child of Array.isArray(v) ? v : [v]) {
      const found = findDescendantPic(child, depth + 1)
      if (found) return found
    }
  }
  return undefined
}

// ── Table (a:tbl) ───────────────────────────────────────────────────

function parseTable(
  node: any,
  tbl: any,
  anchor: ByteAnchor,
  ctx: ParseContext,
): TableElement | null {
  const gridRaw = tbl['a:tblGrid']?.['a:gridCol']
  const gridCols: any[] = Array.isArray(gridRaw) ? gridRaw : gridRaw ? [gridRaw] : []
  const colWidths = gridCols.map((g) => intOr(g['@_w'], 0))
  const trsRaw = tbl['a:tr']
  const trs: any[] = Array.isArray(trsRaw) ? trsRaw : trsRaw ? [trsRaw] : []
  if (!colWidths.length || !trs.length) return null

  // Table style: the styleId referenced by tblPr (embedded definition or PowerPoint built-in style)
  const tblPr = tbl['a:tblPr'] ?? {}
  const styleIdRaw = tblPr['a:tableStyleId']
  const styleId = typeof styleIdRaw === 'string' ? styleIdRaw : styleIdRaw?.['#text']
  const styleDef = resolveTableStyle(styleId, ctx.tableStyles, ctx.theme)
  const flags: TableStyleFlags = {
    firstRow: tblPr['@_firstRow'] === '1',
    lastRow: tblPr['@_lastRow'] === '1',
    firstCol: tblPr['@_firstCol'] === '1',
    lastCol: tblPr['@_lastCol'] === '1',
    bandRow: tblPr['@_bandRow'] === '1',
    bandCol: tblPr['@_bandCol'] === '1',
  }

  const nRows = trs.length
  const nCols = colWidths.length
  const rowHeights = trs.map((tr) => intOr(tr['@_h'], 0))
  const rows: TableCell[][] = trs.map((tr, r) => {
    const tcsRaw = tr['a:tc']
    const tcs: any[] = Array.isArray(tcsRaw) ? tcsRaw : tcsRaw ? [tcsRaw] : []
    const gridCols = tableRowGridCols(
      tcs.map((tc) => ({
        gridSpan: tc['@_gridSpan'] ? parseInt(tc['@_gridSpan'], 10) || 1 : 1,
        merged: tc['@_hMerge'] === '1' || tc['@_vMerge'] === '1',
      })),
    )
    return tcs.map((tc, i) => {
      const c = gridCols[i]!
      const part = styleDef ? cellPartStyle(styleDef, flags, r, c, nRows, nCols) : undefined
      const inside = styleDef ? cellStyleBorders(styleDef, flags, r, c, nRows, nCols) : undefined
      return parseTableCell(tc, ctx, part, inside)
    })
  })

  return {
    id: uid('tbl'),
    type: 'table',
    anchor,
    transform: parseXfrm(node['p:xfrm']),
    name: node['p:nvGraphicFramePr']?.['p:cNvPr']?.['@_name'],
    colWidths,
    rowHeights,
    rows,
    styleFlags: { firstRow: flags.firstRow, bandRow: flags.bandRow },
  }
}

function parseTableCell(
  tc: any,
  ctx: ParseContext,
  part?: TablePartStyle,
  inside?: { l?: Stroke; r?: Stroke; t?: Stroke; b?: Stroke },
): TableCell {
  const tcPr = tc['a:tcPr'] ?? {}
  const cell: TableCell = {}

  if (tc['a:txBody']) {
    // Table-style text defaults (bold white header text etc.) injected at the end of the inheritance chain
    const styleChain: TextStyleLevels[] =
      part && (part.bold !== undefined || part.textColor)
        ? [
            {
              levels: [
                {
                  ...(part.bold !== undefined ? { bold: part.bold } : {}),
                  ...(part.textColor ? { color: part.textColor } : {}),
                },
              ],
            },
          ]
        : []
    const text = parseTextBody(tc['a:txBody'], ctx, styleChain)
    // Cell vertical alignment and insets come from tcPr (bodyPr is usually empty in tables)
    const anchorMap: Record<string, TextBody['anchor']> = { t: 'top', ctr: 'middle', b: 'bottom' }
    if (tcPr['@_anchor']) text.anchor = anchorMap[tcPr['@_anchor']]
    text.insets = {
      l: intOr(tcPr['@_marL'], 91440),
      r: intOr(tcPr['@_marR'], 91440),
      t: intOr(tcPr['@_marT'], 45720),
      b: intOr(tcPr['@_marB'], 45720),
    }
    cell.text = text
  }

  // Fill: explicit tcPr fill > table-style region fill
  const fill = parseFill(tcPr, ctx)
  if (fill && fill.type !== 'none') cell.fill = fill
  else if (part?.fill) cell.fill = part.fill

  // Borders on four edges: a:lnL/R/T/B share a:ln's structure, so reuse parseStroke; style inside-borders as fallback
  const borders: TableCellBorders = {}
  for (const [key, tag] of [
    ['l', 'a:lnL'],
    ['r', 'a:lnR'],
    ['t', 'a:lnT'],
    ['b', 'a:lnB'],
  ] as const) {
    const ln = tcPr[tag]
    if (!ln || typeof ln !== 'object') continue
    const stroke = parseStroke({ 'a:ln': ln }, ctx)
    if (stroke) borders[key] = stroke
  }
  for (const k of ['l', 'r', 't', 'b'] as const) {
    if (inside?.[k] && !borders[k]) borders[k] = inside[k]
  }
  if (Object.keys(borders).length) cell.borders = borders

  const gridSpan = tc['@_gridSpan'] ? parseInt(tc['@_gridSpan'], 10) : undefined
  const rowSpan = tc['@_rowSpan'] ? parseInt(tc['@_rowSpan'], 10) : undefined
  if (gridSpan && gridSpan > 1) cell.gridSpan = gridSpan
  if (rowSpan && rowSpan > 1) cell.rowSpan = rowSpan
  if (tc['@_hMerge'] === '1' || tc['@_vMerge'] === '1') cell.merged = true

  return cell
}

function passthrough(
  anchor: ByteAnchor,
  kind: PassthroughElement['kind'],
  node: any,
): PassthroughElement {
  const spPr = node?.['p:spPr'] ?? node?.['p:grpSpPr']
  const transform = parseXfrm(spPr?.['a:xfrm'])
  return { id: uid('pt'), type: 'passthrough', anchor, transform, kind }
}

// ── Geometry ─────────────────────────────────────────────────────────

function parseXfrm(xfrm: any): Transform {
  const zero: Transform = {
    offset: { x: 0, y: 0, cx: 0, cy: 0 },
    rot: 0,
    flipH: false,
    flipV: false,
  }
  if (!xfrm) return zero
  const off = xfrm['a:off']
  const ext = xfrm['a:ext']
  return {
    offset: {
      x: off ? parseInt(off['@_x'], 10) || 0 : 0,
      y: off ? parseInt(off['@_y'], 10) || 0 : 0,
      cx: ext ? parseInt(ext['@_cx'], 10) || 0 : 0,
      cy: ext ? parseInt(ext['@_cy'], 10) || 0 : 0,
    },
    rot: xfrm['@_rot'] ? parseInt(xfrm['@_rot'], 10) || 0 : 0,
    flipH: xfrm['@_flipH'] === '1' || xfrm['@_flipH'] === 'true',
    flipV: xfrm['@_flipV'] === '1' || xfrm['@_flipV'] === 'true',
  }
}

// ── Fill ─────────────────────────────────────────────────────────────

function parseFill(spPr: any, ctx: ParseContext): Fill | undefined {
  if (!spPr) return undefined
  if ('a:noFill' in spPr) return { type: 'none' }

  const solid = spPr['a:solidFill']
  if (solid) {
    const color = resolveColorNode(solid, ctx)
    if (color) return { type: 'solid', color }
  }

  const grad = spPr['a:gradFill']
  if (grad) return parseGradFill(grad, ctx)

  const blip = spPr['a:blipFill']
  if (blip) {
    const embedId = blipEmbedId(blip['a:blip'])
    const mediaRef = (embedId && ctx.mediaRels?.get(embedId)) || ''
    if (mediaRef) return { type: 'image', mediaRef, mode: 'a:tile' in blip ? 'tile' : 'stretch' }
  }

  const pat = spPr['a:pattFill']
  if (pat) {
    const fg = resolveColorNode(pat['a:fgClr'], ctx) ?? '#000000'
    const bg = resolveColorNode(pat['a:bgClr'], ctx) ?? '#FFFFFF'
    return { type: 'pattern', fg, bg, preset: String(pat['@_prst'] ?? 'pct50') }
  }

  return undefined
}

/** Parse a gradient fill <a:gradFill>. */
function parseGradFill(grad: any, ctx: ParseContext): Fill | undefined {
  const gsLst = grad['a:gsLst']?.['a:gs']
  const list = gsLst ? (Array.isArray(gsLst) ? gsLst : [gsLst]) : []
  const stops = list
    .map((gs: any) => {
      const pos = (parseInt(gs['@_pos'], 10) || 0) / 100000 // 0..100000 → 0..1
      const color = resolveColorNode(gs, ctx)
      return color ? { pos, color } : null
    })
    .filter((s: any): s is { pos: number; color: string } => !!s)
  if (!stops.length) return undefined
  // Linear gradient angle: <a:lin ang=""> (unit 1/60000 degree); radial etc. get the default angle for now
  const ang = grad['a:lin']?.['@_ang']
  const angle = ang != null ? parseInt(ang, 10) || 0 : undefined
  const pathType = grad['a:path']?.['@_path']
  return {
    type: 'gradient',
    stops,
    ...(angle != null ? { angle } : {}),
    ...(pathType === 'circle' || pathType === 'rect' || pathType === 'shape'
      ? { path: pathType }
      : {}),
  }
}

/** Color resolution lives in color.ts (shared with placeholder style inheritance); this is a thin wrapper taking ctx.theme. */
function resolveColorNode(node: any, ctx: ParseContext): string | undefined {
  return resolveColorNodeShared(node, ctx.theme, ctx.phClr)
}

// ── Text ─────────────────────────────────────────────────────────────

function parseTextBody(txBody: any, ctx: ParseContext, phChain: TextStyleLevels[] = []): TextBody {
  const bodyPrRaw = txBody['a:bodyPr']
  const bodyPr = bodyPrRaw && typeof bodyPrRaw === 'object' ? bodyPrRaw : {}
  const anchorMap: Record<string, TextBody['anchor']> = { t: 'top', ctr: 'middle', b: 'bottom' }
  const paras = txBody['a:p']
    ? Array.isArray(txBody['a:p'])
      ? txBody['a:p']
      : [txBody['a:p']]
    : []
  // Inheritance chain: the shape's own <a:lstStyle> first, then the placeholder layout/master chain
  const ownStyle = parseLstStyleLevels(txBody['a:lstStyle'], ctx.theme)
  const chain: Array<TextStyleLevels | undefined> = [ownStyle, ...phChain]
  const paragraphs: Paragraph[] = paras.map((p: any) => parseParagraph(p, ctx, chain))

  let autofit: TextBody['autofit'] = 'none'
  if ('a:normAutofit' in bodyPr) autofit = 'shrink'
  else if ('a:spAutoFit' in bodyPr) autofit = 'resize'
  // Font-shrink/line-spacing-reduction ratios precomputed by PowerPoint (1/1000 %):
  // used directly for rendering; otherwise files with shrunk text would display too large per our own metrics
  const naf = bodyPr['a:normAutofit']
  const nafAttr = (k: string): number | undefined => {
    const v = naf && typeof naf === 'object' ? naf[k] : undefined
    const n = v != null ? parseInt(String(v), 10) : NaN
    return Number.isFinite(n) && n > 0 ? n / 100000 : undefined
  }
  const fontScale = nafAttr('@_fontScale')
  const lnSpcReduction = nafAttr('@_lnSpcReduction')
  const vertRaw = bodyPr['@_vert']
  const vert: TextBody['vert'] =
    vertRaw === 'eaVert' || vertRaw === 'vert' || vertRaw === 'vert270' || vertRaw === 'wordArtVert'
      ? vertRaw
      : undefined

  return {
    paragraphs,
    anchor: bodyPr['@_anchor'] ? anchorMap[bodyPr['@_anchor']] : undefined,
    insets: {
      l: intOr(bodyPr['@_lIns'], 91440),
      t: intOr(bodyPr['@_tIns'], 45720),
      r: intOr(bodyPr['@_rIns'], 91440),
      b: intOr(bodyPr['@_bIns'], 45720),
    },
    autofit,
    ...(fontScale != null ? { fontScale } : {}),
    ...(lnSpcReduction != null ? { lnSpcReduction } : {}),
    wrap: bodyPr['@_wrap'] !== 'none',
    ...(vert ? { vert } : {}),
  }
}

/** <a:spcPct val="150000"/> → 150 (%). */
function spcPct(node: any): number | undefined {
  const v = node?.['a:spcPct']?.['@_val']
  return v != null ? (parseInt(v, 10) || 0) / 1000 : undefined
}

/** <a:spcPts val="2400"/> → 24 (pt). */
function spcPts(node: any): number | undefined {
  const v = node?.['a:spcPts']?.['@_val']
  return v != null ? (parseInt(v, 10) || 0) / 100 : undefined
}

function parseParagraph(
  p: any,
  ctx: ParseContext,
  chain: Array<TextStyleLevels | undefined> = [],
): Paragraph {
  const pPr = p['a:pPr'] ?? {}
  const alignMap: Record<string, Paragraph['align']> = {
    l: 'left',
    ctr: 'center',
    r: 'right',
    just: 'justify',
  }
  const level = pPr['@_lvl'] ? parseInt(pPr['@_lvl'], 10) : undefined
  // Inherited default style for this level (shape lstStyle → layout ph → master ph → master txStyles)
  const dflt = mergeTextStyleChain(chain, level ?? 0)
  const runsRaw = p['a:r'] ? (Array.isArray(p['a:r']) ? p['a:r'] : [p['a:r']]) : []
  const runs: TextRun[] = runsRaw.map((r: any) => {
    const run = parseRun(r, ctx, dflt)
    // a:fld rewritten to a:r by parseShapeFragment (a genuine a:r never carries @_type)
    if (r?.['@_type']) run.field = String(r['@_type'])
    return run
  })
  // <a:fld> reaching here in its original form (parse paths without the fragment
  // rewrite, e.g. master footers): order relative to a:r is lost, appending is the
  // legacy fallback — a footer fld usually owns its paragraph.
  const fldsRaw = p['a:fld'] ? (Array.isArray(p['a:fld']) ? p['a:fld'] : [p['a:fld']]) : []
  for (const f of fldsRaw) {
    const run = parseRun(f, ctx, dflt)
    if (f?.['@_type']) run.field = String(f['@_type'])
    runs.push(run)
  }

  // Line spacing: spcPct (%) or spcPts (absolute pt); space before/after: spcPts / spcPct (as % of single line height).
  // An explicit node overrides wholesale; otherwise inherited from the lstStyle/placeholder/master txStyles chain
  const lnSpcNode = pPr['a:lnSpc']
  const lineHeight = lnSpcNode ? spcPct(lnSpcNode) : dflt?.lineHeight
  const lineExact = lnSpcNode ? spcPts(lnSpcNode) : dflt?.lineExact
  const befNode = pPr['a:spcBef']
  const spaceBefore = befNode ? spcPts(befNode) : dflt?.spaceBefore
  const spaceBeforePct = befNode ? spcPct(befNode) : dflt?.spaceBeforePct
  const aftNode = pPr['a:spcAft']
  const spaceAfter = aftNode ? spcPts(aftNode) : dflt?.spaceAfter
  const spaceAfterPct = aftNode ? spcPct(aftNode) : dflt?.spaceAfterPct

  // Bullets: buNone / buChar / buAutoNum (color from buClr, defaults to the run text color)
  let bullet: Paragraph['bullet']
  if (pPr['a:buNone'] !== undefined) bullet = { type: 'none' }
  else if (pPr['a:buChar']?.['@_char'] != null) {
    bullet = { type: 'char', char: decodeCharRefs(String(pPr['a:buChar']['@_char'])) }
  } else if (pPr['a:buAutoNum']) {
    bullet = { type: 'number' }
    if (pPr['a:buAutoNum']['@_type']) bullet.numType = String(pPr['a:buAutoNum']['@_type'])
  }
  if (bullet && bullet.type !== 'none') {
    if (pPr['a:buClr']) {
      const c = resolveColorNode(pPr['a:buClr'], ctx)
      if (c) bullet.color = c
    }
    if (pPr['a:buFont']?.['@_typeface']) bullet.font = String(pPr['a:buFont']['@_typeface'])
    if (pPr['a:buSzPct']?.['@_val']) {
      const v = parseInt(pPr['a:buSzPct']['@_val'], 10)
      if (Number.isFinite(v)) bullet.sizePct = v / 1000
    }
  }

  const marLRaw = pPr['@_marL'] != null ? parseInt(pPr['@_marL'], 10) : undefined
  const indentRaw = pPr['@_indent'] != null ? parseInt(pPr['@_indent'], 10) : undefined
  // Inheritance fallback: explicit pPr wins, level defaults from the placeholder/lstStyle chain fill gaps
  // (the master bodyStyle's buChar/marL/indent is where classic-template body bullets come from)
  const effBullet = bullet ?? dflt?.bullet
  const hasMarL = marLRaw != null && !Number.isNaN(marLRaw)
  const hasIndent = indentRaw != null && !Number.isNaN(indentRaw)
  const marL = hasMarL ? marLRaw : dflt?.marL
  const indent = hasIndent ? indentRaw : dflt?.indent

  // Record which properties come from an explicit pPr (the rebuild path writes only explicit items; inherited values are not baked in)
  const pPrExplicit: NonNullable<Paragraph['pPrExplicit']> = {
    ...(pPr['@_algn'] ? { align: true } : {}),
    ...(lnSpcNode ? { lnSpc: true } : {}),
    ...(befNode ? { spcBef: true } : {}),
    ...(aftNode ? { spcAft: true } : {}),
    ...(bullet ? { bullet: true } : {}),
    ...(hasMarL ? { marL: true } : {}),
    ...(hasIndent ? { indent: true } : {}),
  }

  return {
    runs,
    align: pPr['@_algn'] ? alignMap[pPr['@_algn']] : dflt?.align,
    level,
    pPrExplicit,
    ...(lineHeight != null ? { lineHeight } : {}),
    ...(lineExact != null ? { lineExact } : {}),
    ...(spaceBefore != null ? { spaceBefore } : {}),
    ...(spaceAfter != null ? { spaceAfter } : {}),
    ...(spaceBeforePct != null ? { spaceBeforePct } : {}),
    ...(spaceAfterPct != null ? { spaceAfterPct } : {}),
    ...(effBullet ? { bullet: effBullet } : {}),
    ...(marL != null ? { marL } : {}),
    ...(indent != null ? { indent } : {}),
  }
}

/** fast-xml-parser does not decode numeric character references in attributes (&#x2022; etc.); done here. */
function decodeCharRefs(s: string): string {
  return s
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(parseInt(d, 10)))
}

// East Asian (OOXML a:ea bucket): Chinese/Japanese + Hangul (jamo/syllables), matching the EAW fullwidth ranges in metrics
const CJK_RE =
  /[\u1100-\u11ff\u2e80-\u303e\u3041-\u33ff\u3400-\u9fff\ua960-\ua97f\uac00-\ud7a3\uf900-\ufaff\ufe30-\ufe4f\uff00-\uffef]/
// Complex Script (OOXML a:cs bucket): Hebrew/Arabic/Indic/Thai/Lao/Myanmar/Khmer + presentation forms
const CS_RE =
  /[\u0590-\u07bf\u08a0-\u08ff\u0900-\u0dff\u0e00-\u0eff\u1000-\u109f\u1780-\u17ff\ufb1d-\ufdff\ufe70-\ufeff]/

function parseRun(r: any, ctx: ParseContext, dflt?: LevelTextStyle): TextRun {
  const rPr = r['a:rPr'] ?? {}
  const rawT = r['a:t']
  const text = decodeCharRefs(
    typeof rawT === 'string'
      ? rawT
      : rawT == null
        ? ''
        : typeof rawT === 'object'
          ? String(rawT['#text'] ?? '')
          : String(rawT),
  )
  const hlink = rPr['a:hlinkClick']
  const hlinkTarget = hlink?.['@_r:id'] ? ctx.hlinkRels?.get(String(hlink['@_r:id'])) : undefined
  const fill = rPr['a:solidFill']
  // PowerPoint styles linked runs with the theme hlink color unless the run has an explicit fill
  const color =
    (fill ? resolveColorNode(fill, ctx) : undefined) ??
    (hlinkTarget ? ctx.theme?.colors?.hlink : undefined) ??
    dflt?.color
  // Whether the color is display-only: from schemeClr/inheritance (not an explicit run srgbClr).
  // The patch path uses this to avoid baking theme colors into srgbClr (theme switches must stay linked)
  const colorFollowsTheme = color != null && !(fill && fill['a:srgbClr'])
  const colorInherited = color != null && !fill
  // Font: run explicit (incl. +mj/+mn theme refs) → inherited default → theme font
  const latin = resolveFontRef(rPr['a:latin']?.['@_typeface'], ctx.theme) ?? dflt?.latinFont
  const ea = resolveFontRef(rPr['a:ea']?.['@_typeface'], ctx.theme) ?? dflt?.eaFont
  const cs = resolveFontRef(rPr['a:cs']?.['@_typeface'], ctx.theme) ?? dflt?.csFont
  // Pick the bucket by script: complex script → a:cs, CJK → a:ea, otherwise → a:latin; fall back through buckets when missing
  const fontFamily =
    (CS_RE.test(text) ? (cs ?? latin ?? ea) : CJK_RE.test(text) ? (ea ?? latin) : (latin ?? ea)) ??
    ctx.theme?.minorFont
  const bAttr = rPr['@_b']
  const iAttr = rPr['@_i']
  // Text outline <a:rPr><a:ln> (WordArt): only solid-color outlines are modeled, kept by the rebuild path
  let outline: TextRun['outline']
  const lnNode = rPr['a:ln']
  if (lnNode && typeof lnNode === 'object' && lnNode['a:solidFill']) {
    const lnColor = resolveColorNode(lnNode['a:solidFill'], ctx)
    if (lnColor) {
      const w = lnNode['@_w'] != null ? parseInt(lnNode['@_w'], 10) : 9525
      outline = { color: lnColor, widthEmu: Number.isFinite(w) ? w : 9525 }
    }
  }
  const uAttr = rPr['@_u']
  const strikeAttr = rPr['@_strike']
  const hasStrike = strikeAttr !== undefined && strikeAttr !== 'noStrike'
  const latinRaw = rPr['a:latin']?.['@_typeface']
  const eaRaw = rPr['a:ea']?.['@_typeface']
  const csRaw = rPr['a:cs']?.['@_typeface']
  // Linked runs underline by default (PowerPoint hlink styling) unless u is explicit
  const linkUnderline = hlinkTarget != null && uAttr === undefined
  return {
    text,
    bold: bAttr != null ? bAttr === '1' || bAttr === 'true' : !!dflt?.bold,
    italic: iAttr != null ? iAttr === '1' || iAttr === 'true' : !!dflt?.italic,
    underline: (uAttr !== undefined && uAttr !== 'none') || linkUnderline,
    ...(uAttr !== undefined && uAttr !== 'none' ? { underlineStyle: String(uAttr) } : {}),
    ...(linkUnderline ? { underlineImplicit: true } : {}),
    ...(hasStrike ? { strike: true, strikeStyle: String(strikeAttr) } : {}),
    ...(latinRaw ? { latinFont: String(latinRaw) } : {}),
    ...(eaRaw ? { eaFont: String(eaRaw) } : {}),
    ...(csRaw ? { csFont: String(csRaw) } : {}),
    ...(!latinRaw && !eaRaw ? { fontImplicit: true } : {}),
    fontSize: rPr['@_sz'] ? parseInt(rPr['@_sz'], 10) / 100 : dflt?.fontSize,
    ...(rPr['@_sz'] ? {} : { fontSizeImplicit: true }),
    ...(rPr['@_spc'] ? { letterSpacing: parseInt(rPr['@_spc'], 10) / 100 } : {}),
    ...(rPr['@_baseline'] ? { baseline: parseInt(rPr['@_baseline'], 10) / 1000 } : {}),
    fontFamily,
    color,
    ...(colorFollowsTheme ? { colorFollowsTheme } : {}),
    ...(colorInherited ? { colorInherited } : {}),
    ...(outline ? { outline } : {}),
    ...(hlink?.['@_r:id']
      ? {
          hyperlinkRId: String(hlink['@_r:id']),
          ...(hlinkTarget ? { hyperlink: hlinkTarget } : {}),
          ...(hlink['@_action'] ? { hyperlinkAction: String(hlink['@_action']) } : {}),
          ...(hlink['@_tooltip'] ? { hyperlinkTooltip: String(hlink['@_tooltip']) } : {}),
        }
      : {}),
  }
}

// ── master/layout decoration layer ───────────────────────────────────

export interface DecorationOptions {
  /**
   * Footer-family placeholder types allowed to render (subset of ftr/sldNum/dt).
   * Such placeholders on the master show only when <p:hf> hasn't disabled them
   * and the slide has no placeholder of the same type.
   */
  hfTypes?: Set<string>
  /** Actual value of the slide-number field <a:fld type="slidenum"> (replaces the cached text) */
  slideNum?: number
}

/**
 * Parse decoration-layer elements from layout/master XML (read-only render, not saved):
 * - Non-placeholder concrete shapes (logos/color bars/decor images/connectors/groups) are all kept;
 * - Placeholders keep only the footer family specified by opts.hfTypes (ftr/sldNum/dt);
 *   the rest (title/body/pic etc. are content carriers, overridden by the slide) are skipped.
 */
export function parseDecorations(
  xml: string,
  ctx: ParseContext,
  opts: DecorationOptions = {},
): SlideElement[] {
  let scan: ReturnType<typeof scanSlide>
  try {
    scan = scanSlide(xml)
  } catch {
    return []
  }
  const out: SlideElement[] = []
  scan.elements.forEach((sp, idx) => {
    const fragXml = xml.slice(sp.start, sp.end)
    // The decoration layer has no byte fidelity (never written back); the anchor is a placeholder
    const anchor: ByteAnchor = { spIndex: -(idx + 1), originalXml: '', range: [0, 0] }
    const el = parseShapeFragment(sp, fragXml, anchor, ctx)
    if (!el || el.type === 'passthrough') return
    const ph = (el as { placeholder?: string }).placeholder
    if (ph !== undefined) {
      if (!opts.hfTypes?.has(ph)) return
    } else if (/<p:ph[\s/>]/.test(fragXml)) {
      // Untyped <p:ph idx="…"/> (body family) or picture/table placeholders: content carriers, skip
      return
    }
    if (opts.slideNum != null) substituteSlideNum(el, opts.slideNum)
    out.push(el)
  })
  return out
}

/** Recursively replace slidenum fields in element text with the actual slide number. */
function substituteSlideNum(el: SlideElement, num: number): void {
  if (el.type === 'group') {
    for (const c of (el as GroupElement).children) substituteSlideNum(c, num)
    return
  }
  const text = (el as TextElement).text
  if (!text) return
  for (const p of text.paragraphs) {
    for (const r of p.runs) {
      if (r.field === 'slidenum') r.text = String(num)
    }
  }
}

function intOr(v: any, dflt: number): number {
  if (v === undefined || v === null) return dflt
  const n = parseInt(v, 10)
  return Number.isNaN(n) ? dflt : n
}

export { EMU_PER_PT }
