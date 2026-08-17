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
import { type Theme, resolveFontRef, resolveSchemeColor } from './theme'
import { resolveColorNode as resolveColorNodeShared } from './color'
import {
  resolvePlaceholderTransform,
  resolvePlaceholderAnchor,
  resolvePlaceholderFillSpPr,
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
import { parseChartExXml } from './chartex'
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
  /** Effective fill of the enclosing group (<a:grpFill/> in a child resolves to this) */
  groupFill?: Fill
  /** Placeholder color for resolving style ref templates (value substituted for schemeClr val="phClr") */
  phClr?: string
  /** Media rId → zip path, for picture parsing */
  mediaRels?: Map<string, string>
  /** Hyperlink rId → resolved target: external url, or "slide:N" (0-based) for slide jumps */
  hlinkRels?: Map<string, string>
  /** Chart rId → chartN.xml content (chart part referenced by a graphicFrame) */
  chartXmls?: Map<string, string>
  /** Chart rId → chartUserShapes drawing content (user-drawn overlays on the chart) */
  chartUserShapes?: Map<string, string>
  /** Audio/video rId → media zip path or external URL (r:link of videoFile/audioFile) */
  avRels?: Map<string, { target: string; external?: boolean }>
  /** SmartArt: diagramData rId (dgm:relIds@r:dm) → data part content (fallback layout without a drawing part) */
  diagramDatas?: Map<string, string>
  /** SmartArt: diagramData rId (dgm:relIds@r:dm) → prerendered drawing part content */
  diagramDrawings?: Map<string, string>
  /** SmartArt: diagramLayout rId (dgm:relIds@r:lo) → layout definition part content */
  diagramLayouts?: Map<string, string>
  /** SmartArt: diagramColors rId (dgm:relIds@r:cs) → color definition part content */
  diagramColors?: Map<string, string>
  /** Placeholder geometry inheritance table: from the slideLayout (read-only) */
  layoutPlaceholders?: PlaceholderMap
  /** Placeholder geometry inheritance table: from the slideMaster (read-only, fallback when the layout lacks it) */
  masterPlaceholders?: PlaceholderMap
  /** master <p:txStyles> text style defaults (title/body/other families) */
  masterTextStyles?: MasterTextStyles
  /** presentation.xml <p:defaultTextStyle>: base defaults for non-placeholder text boxes */
  defaultTextStyle?: TextStyleLevels
  /** Full layout XML (read-only, for background inheritance) */
  layoutBg?: string
  /** Full master XML (read-only, background inheritance fallback) */
  masterBg?: string
  /** Layout/master part image rels (blip rIds in inherited backgrounds live in those parts) */
  layoutMediaRels?: Map<string, string>
  masterMediaRels?: Map<string, string>
  /** Theme part image rels (blip rIds inside fmtScheme fill templates live in the theme part) */
  themeMediaRels?: Map<string, string>
  /** Chart rId → that chart part's own image rels (for chart background picture fills) */
  chartMediaRels?: Map<string, Map<string, string>>
  /** Diagram data rId → the drawing part's own image rels (SmartArt picture fills) */
  diagramMediaRels?: Map<string, Map<string, string>>
  /** ppt/tableStyles.xml source (table style definitions, read-only) */
  tableStyles?: string
  /** Legacy VML previews: oleObj spid → preview image zip path (v:shape/v:imagedata) */
  vmlPreviews?: Map<string, string>
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
  // Inherited backgrounds resolve blip rIds against their own part's rels, not the slide's.
  const ownBackground = parseBackground(slideXml, ctx)
  // With no <p:bg> on any layer, PowerPoint paints the bg1 scheme color — visible when a
  // clrMapOvr remaps bg1 away from white (layout-clrmap-override)
  const defaultBg1 = resolveSchemeColor('bg1', ctx.theme)
  const background =
    ownBackground ??
    (ctx.layoutBg
      ? parseBackground(ctx.layoutBg, { ...ctx, mediaRels: ctx.layoutMediaRels ?? ctx.mediaRels })
      : undefined) ??
    (ctx.masterBg
      ? parseBackground(ctx.masterBg, { ...ctx, mediaRels: ctx.masterMediaRels ?? ctx.mediaRels })
      : undefined) ??
    (defaultBg1 && defaultBg1.toUpperCase() !== '#FFFFFF'
      ? { type: 'solid' as const, color: defaultBg1 }
      : undefined)
  // Only real slides carry showMasterSp (<p:sldLayout> has "sldLayout" so \b won't match)
  const masterSpHidden = /<p:sld\b[^>]*\bshowMasterSp="(?:0|false)"/.test(slideXml)

  return {
    path,
    originalXml: slideXml,
    bodyPrefix: scan.bodyPrefix,
    bodySuffix: scan.bodySuffix,
    elements,
    layoutPath,
    masterPath,
    ...(background ? { background } : {}),
    ...(ownBackground || /<p:bg[\s>]/.test(slideXml) ? { bgOwn: true } : {}),
    ...(masterSpHidden ? { masterSpHidden: true } : {}),
  }
}

/** Extract the <p:bg> background fill from slide/layout/master XML (read-only). */
export function parseBackground(xml: string, ctx: ParseContext): Fill | undefined {
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
  // <p:bgRef idx>: theme fill template (1..3 → fillStyleLst, 1001..1003 → bgFillStyleLst)
  // instantiated with the referenced color as phClr; color-only fallback when unresolvable.
  const bgRef = bg?.['p:bgRef']
  if (bgRef) {
    const color = resolveColorNode(bgRef, ctx)
    const idx = parseInt(String(bgRef['@_idx'] ?? ''), 10)
    const tpl =
      idx >= 1001
        ? ctx.theme?.bgFillStyles?.[idx - 1001]
        : idx >= 1
          ? ctx.theme?.fillStyles?.[idx - 1]
          : undefined
    if (tpl) {
      const fill = parseFill(tpl, {
        ...ctx,
        phClr: color,
        mediaRels: ctx.themeMediaRels ?? ctx.mediaRels,
      })
      if (fill) return fill
    }
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
      // fragXml explicitly: decoration anchors carry an empty originalXml, and custGeom
      // parses from raw bytes — without it master/layout freeforms degrade to rects
      return parseSpShape(node, anchor, ctx, fragXml)
    case 'p:pic':
      return parsePicture(node, anchor, ctx)
    case 'p:grpSp':
      return parseGroup(node, anchor, ctx, fragXml)
    case 'p:graphicFrame':
      return graphicFramePassthrough(node, anchor, ctx)
    case 'p:cxnSp':
      return parseConnector(node, anchor, ctx)
    case 'mc:AlternateContent': {
      // Prefer a Choice we can render (chartEx frames live here); otherwise the Fallback
      // picture preview; otherwise keep the block byte-only (no chip, save replays it verbatim)
      const choicesRaw = node['mc:Choice']
      const choices: any[] = Array.isArray(choicesRaw) ? choicesRaw : choicesRaw ? [choicesRaw] : []
      for (const ch of choices) {
        const gfRaw = ch?.['p:graphicFrame']
        const gf = Array.isArray(gfRaw) ? gfRaw[0] : gfRaw
        if (!gf) continue
        const el = graphicFramePassthrough(gf, anchor, ctx)
        if (el && el.type !== 'passthrough') return el
      }
      const fb = node['mc:Fallback']
      const picRaw = fb?.['p:pic']
      const pic = Array.isArray(picRaw) ? picRaw[0] : picRaw
      if (pic) return parsePicture(pic, anchor, ctx)
      const spRaw = fb?.['p:sp']
      const sp2 = Array.isArray(spRaw) ? spRaw[0] : spRaw
      if (sp2) return parseSpShape(sp2, anchor, ctx)
      const silent = passthrough(anchor, 'unknown', node)
      silent.noChip = true
      return silent
    }
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
    : ctx.defaultTextStyle && node['p:nvSpPr']?.['p:cNvSpPr']?.['@_txBox'] === '1'
      ? [ctx.defaultTextStyle]
      : []
  // <p:style> fontRef color ranks between the shape's own lstStyle and the
  // layout/master defaults (a styled placeholder shows the style color, not the
  // master txStyles color — PowerPoint behavior, bnc904423)
  const fontRefColor = resolveColorNode(node['p:style']?.['a:fontRef'], ctx)
  const chainLayers = fontRefColor ? [{ levels: [{ color: fontRefColor }] }, ...phChain] : phChain
  const text = txBody ? parseTextBody(txBody, ctx, chainLayers) : undefined
  // bodyPr anchor inherits along the placeholder chain (e.g. master titles anchor="ctr")
  if (ph && text && !text.anchor) {
    const inherited = resolvePlaceholderAnchor(
      ctx.layoutPlaceholders,
      ctx.masterPlaceholders,
      phType,
      phIdx,
    )
    if (inherited) text.anchor = inherited
  }

  let stroke = parseStroke(spPr, ctx)
  let shadow = parseShadow(spPr, ctx)
  let glow = parseGlow(spPr, ctx)
  // <a:fillOverlay> holds a second fill element directly (a:gradFill/…), so parseFill reads it like an spPr
  const overlayNode = spPr?.['a:effectLst']?.['a:fillOverlay']
  const fillOverlay = overlayNode ? parseFill(overlayNode, ctx) : undefined

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
        const tplFill = tpl
          ? parseFill(tpl, { ...ctx, phClr, mediaRels: ctx.themeMediaRels ?? ctx.mediaRels })
          : undefined
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

  // Placeholder fill inheritance: layout, then master spPr fill (blip rIds resolve in that part's rels)
  if (fill === undefined && ph) {
    const inh = resolvePlaceholderFillSpPr(
      ctx.layoutPlaceholders,
      ctx.masterPlaceholders,
      phType,
      phIdx,
    )
    if (inh) {
      fill = parseFill(inh.spPr, {
        ...ctx,
        mediaRels:
          (inh.layer === 'layout' ? ctx.layoutMediaRels : ctx.masterMediaRels) ?? ctx.mediaRels,
      })
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
    ...(node['@_useBgFill'] === '1' || node['@_useBgFill'] === 'true' ? { useBgFill: true } : {}),
    ...(fillOverlay && fillOverlay.type !== 'none' ? { fillOverlay } : {}),
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
  // grpSpPr fill: children with <a:grpFill/> inherit it (a nested grpFill defers to the outer group)
  const groupFill =
    'a:grpFill' in grpSpPr ? ctx.groupFill : (parseFill(grpSpPr, ctx) ?? ctx.groupFill)
  const childCtx = groupFill === ctx.groupFill ? ctx : { ...ctx, groupFill }

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
      const el = parseGroupChild(tag, child, childCtx, slice?.xml)
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
  let transform = parseXfrm(spPr['a:xfrm'])
  // Pictures dropped into a placeholder may omit <a:xfrm> entirely; geometry comes from layout/master
  const picPh = node['p:nvPicPr']?.['p:nvPr']?.['p:ph']
  if (picPh && !spPr['a:xfrm']) {
    const inherited = resolvePlaceholderTransform(
      ctx.layoutPlaceholders,
      ctx.masterPlaceholders,
      picPh['@_type'],
      picPh['@_idx'] != null ? String(picPh['@_idx']) : undefined,
    )
    if (inherited) transform = inherited
  }
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
  // Pic's own spPr fill: PowerPoint draws it as a backdrop behind the (possibly translucent) blip
  const fill = parseFill(spPr, ctx)
  const duotone = parseDuotone(blip, ctx)
  const clrChange = parseClrChange(blip, ctx)
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
    ...(fill ? { fill } : {}),
    ...(duotone ? { duotone } : {}),
    ...(clrChange ? { clrChange } : {}),
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

/**
 * chartUserShapes overlays, straight lines only (cdr:relSizeAnchor from/to are
 * fractions of the chart frame). Other overlay shapes are rare and skipped.
 */
function parseChartUserLines(
  usXml: string,
  ctx: ParseContext,
): NonNullable<import('./chart').ChartModel['userLines']> {
  let doc: any
  try {
    doc = parser.parse(usXml)
  } catch {
    return []
  }
  const anchorsRaw = doc['c:userShapes']?.['cdr:relSizeAnchor']
  const anchors: any[] = Array.isArray(anchorsRaw) ? anchorsRaw : anchorsRaw ? [anchorsRaw] : []
  const out: NonNullable<import('./chart').ChartModel['userLines']> = []
  for (const a of anchors) {
    const spRaw = a?.['cdr:sp']
    const sp = Array.isArray(spRaw) ? spRaw[0] : spRaw
    const spPr = sp?.['cdr:spPr']
    const prst = spPr?.['a:prstGeom']?.['@_prst']
    if (prst !== 'line' && prst !== 'straightConnector1') continue
    const color = resolveColorNode(spPr?.['a:ln']?.['a:solidFill'], ctx)
    if (!color) continue
    const frac = (n: any) => {
      const v = parseFloat(String(n ?? ''))
      return Number.isFinite(v) ? v : 0
    }
    const w = parseInt(spPr?.['a:ln']?.['@_w'], 10)
    out.push({
      x1: frac(a?.['cdr:from']?.['cdr:x']),
      y1: frac(a?.['cdr:from']?.['cdr:y']),
      x2: frac(a?.['cdr:to']?.['cdr:x']),
      y2: frac(a?.['cdr:to']?.['cdr:y']),
      color,
      widthEmu: Number.isFinite(w) && w > 0 ? w : 9525,
    })
  }
  return out
}

function graphicFramePassthrough(node: any, anchor: ByteAnchor, ctx: ParseContext): SlideElement {
  const data = node['a:graphic']?.['a:graphicData']
  const uri: string = data?.['@_uri'] ?? ''
  // Table: semantic parsing (read-only render; save still uses the anchor's original bytes)
  if (uri.includes('/table') && data?.['a:tbl']) {
    const table = parseTable(node, data['a:tbl'], anchor, ctx)
    if (table) return table
  }
  // chartEx (cx namespace, e.g. funnel/sunburst): same read-only semantics as classic charts
  if (uri.includes('/chartex')) {
    const rid = data?.['cx:chart']?.['@_r:id']
    const chartXml = rid ? ctx.chartXmls?.get(String(rid)) : undefined
    const model = chartXml ? parseChartExXml(chartXml, ctx.theme) : null
    if (model) {
      const cNvPr = node['p:nvGraphicFramePr']?.['p:cNvPr']
      return {
        id: uid('chart'),
        type: 'chart',
        anchor,
        transform: parseXfrm(node['p:xfrm']),
        name: cNvPr?.['@_name'],
        chart: model,
      } satisfies ChartElement
    }
  }
  // Chart: read the referenced chart part for semantic parsing (read-only render; save uses original bytes)
  if (uri.includes('/chart')) {
    const rid = data?.['c:chart']?.['@_r:id']
    const chartXml = rid ? ctx.chartXmls?.get(rid) : undefined
    // Fill resolver bound to the chart part's own rels (blip rIds live there, not on the slide)
    const chartFillCtx: ParseContext = { ...ctx, mediaRels: ctx.chartMediaRels?.get(String(rid)) }
    const model = chartXml
      ? parseChartXml(chartXml, ctx.theme, (spPr) => parseFill(spPr, chartFillCtx))
      : null
    if (model) {
      const usXml = rid ? ctx.chartUserShapes?.get(String(rid)) : undefined
      if (usXml) {
        const lines = parseChartUserLines(usXml, ctx)
        if (lines.length) model.userLines = lines
      }
    }
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
      // Picture-fill rIds inside the drawing resolve against the drawing part's own rels
      const drawingRels = dm ? ctx.diagramMediaRels?.get(String(dm)) : undefined
      const shapes = parseDiagramDrawing(
        drawingXml,
        drawingRels ? { ...ctx, mediaRels: drawingRels } : ctx,
      )
      // A single text-less shape is a stub (writers emit just a background rect);
      // fall through to the layout fallback instead of drawing one giant block
      const meaningful =
        shapes.length > 1 ||
        shapes.some((sh: any) =>
          sh.text?.paragraphs?.some((par: any) => par.runs?.some((r: any) => r.text?.trim())),
        )
      if (shapes.length && meaningful) el.previewShapes = shapes
    }
    if (!el.previewShapes?.length && transform) {
      // No prerendered drawing — or an empty one (unfinished writers): lay the
      // node tree out ourselves, dispatched on the layout part's uniqueId
      const dataXml = dm ? ctx.diagramDatas?.get(String(dm)) : undefined
      if (dataXml) {
        const lo = data?.['dgm:relIds']?.['@_r:lo']
        const cs = data?.['dgm:relIds']?.['@_r:cs']
        const layoutXml = lo ? ctx.diagramLayouts?.get(String(lo)) : undefined
        const uniqueId = layoutXml
          ? /\buniqueId="([^"]+)"/.exec(layoutXml)?.[1]?.split('/').pop()
          : undefined
        const colorsXml = cs ? ctx.diagramColors?.get(String(cs)) : undefined
        const shapes = layoutDiagramFallback(
          dataXml,
          ctx,
          transform.offset.cx,
          transform.offset.cy,
          uniqueId,
          colorsXml,
        )
        if (shapes.length) el.previewShapes = shapes
      }
    }
  }
  // OLE read-only preview: p:pic embedded in graphicData (often inside mc:AlternateContent/mc:Fallback),
  // or — in legacy converted decks — a VML shape keyed by the oleObj spid (v:imagedata, usually WMF/EMF)
  if (kind === 'ole') {
    // A substitute pic without a blip (Aspose writes an empty one) must not
    // short-circuit the VML spid preview path
    const pic = findDescendantPic(data)
    const hasBlip = pic?.['p:blipFill']?.['a:blip']?.['@_r:embed'] != null
    if (pic && hasBlip) el.previewPicture = parsePicture(pic, anchor, ctx)
    else if (transform) {
      const oleObj =
        data?.['p:oleObj'] ??
        data?.['mc:AlternateContent']?.['mc:Choice']?.['p:oleObj'] ??
        data?.['mc:AlternateContent']?.['mc:Fallback']?.['p:oleObj']
      const spid = oleObj?.['@_spid']
      const mediaRef = spid != null ? ctx.vmlPreviews?.get(String(spid)) : undefined
      if (mediaRef) {
        el.previewPicture = { id: uid('olepic'), type: 'picture', anchor, transform, mediaRef }
      }
    }
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
    // dsp:txXfrm gives the text its own frame; split text off the shape so both
    // render with their proper transforms (text rotation is shape rot + txXfrm rot)
    const txXfrm = sp['p:txXfrm']
    const txBody = sp['p:txBody']
    if (txXfrm && typeof txXfrm === 'object' && txBody) {
      const shapeOnly = { ...sp }
      delete shapeOnly['p:txBody']
      const shapeEl = parseSpShape(shapeOnly, anchor, ctx)
      if (shapeEl.type !== 'passthrough') out.push(shapeEl)
      const spRot = parseInt(sp['p:spPr']?.['a:xfrm']?.['@_rot'] ?? '0', 10) || 0
      const txRot = parseInt(txXfrm['@_rot'] ?? '0', 10) || 0
      const textXfrm = { ...txXfrm, '@_rot': String(spRot + txRot) }
      // Keep fontRef for the text color only. Copying the whole p:style still
      // applies effectRef — empty a:effectLst does not block the !shadow && !glow
      // fallback — so the split text frame would redraw the shape's shadow/glow.
      const fontRef = sp['p:style']?.['a:fontRef']
      const textSp = {
        'p:nvSpPr': sp['p:nvSpPr'],
        'p:spPr': {
          'a:xfrm': textXfrm,
          'a:prstGeom': { '@_prst': 'rect' },
          'a:noFill': {},
          'a:ln': { 'a:noFill': {} },
        },
        ...(fontRef ? { 'p:style': { 'a:fontRef': fontRef } } : {}),
        'p:txBody': txBody,
      }
      const textEl = parseSpShape(textSp, anchor, ctx)
      if (textEl.type !== 'passthrough') out.push(textEl)
      continue
    }
    const el = parseSpShape(sp, anchor, ctx)
    if (el.type !== 'passthrough') out.push(el)
  }
  return out
}

/** Find the first p:pic in the graphicData subtree (piercing wrappers like mc:AlternateContent). */
/**
 * SmartArt fallback layout when the pptx carries no prerendered drawing part
 * (Aspose-generated decks): only flat node lists under the doc root are laid out,
 * as the basic-block-list snake (child h = 0.6w, gap ≈ 0.115w, last row centered,
 * blocks max-fit to the frame, grid anchored top-left; measured against PowerPoint).
 * Hierarchies (org charts etc.) still fall back to the chip.
 */
/** One dgm data node: own text paragraphs, optional explicit dgm:spPr, children in srcOrd order. */
interface DgmTreeNode {
  id: string
  texts: string[]
  spPr?: any
  children: DgmTreeNode[]
}

/** Depth-first bullet lines of a node's descendants (lvl 1 = direct child). */
function dgmBulletLines(node: DgmTreeNode, lvl = 1): Array<{ text: string; lvl: number }> {
  const out: Array<{ text: string; lvl: number }> = []
  for (const c of node.children) {
    for (const t of c.texts.length ? c.texts : ['']) if (t) out.push({ text: t, lvl })
    out.push(...dgmBulletLines(c, lvl + 1))
  }
  return out
}

/** colors1.xml node1 fillClrLst cycle (accent scheme colors), resolved against the theme. */
function diagramCycleColors(colorsXml: string | undefined, ctx: ParseContext): string[] {
  const fallback = resolveColorNode({ 'a:schemeClr': { '@_val': 'accent1' } }, ctx) ?? '#4472C4'
  if (!colorsXml) return [fallback]
  const lbl =
    /<dgm:styleLbl name="node1">([\s\S]*?)<\/dgm:styleLbl>/.exec(colorsXml)?.[1] ??
    /<dgm:styleLbl name="node0">([\s\S]*?)<\/dgm:styleLbl>/.exec(colorsXml)?.[1]
  const lst = lbl ? /<dgm:fillClrLst[^>]*>([\s\S]*?)<\/dgm:fillClrLst>/.exec(lbl)?.[1] : undefined
  if (!lst) return [fallback]
  const out: string[] = []
  for (const m of lst.matchAll(
    /<a:schemeClr val="([^"]+)"\s*\/>|<a:srgbClr val="([^"]+)"\s*\/>/g,
  )) {
    const c = m[1]
      ? resolveColorNode({ 'a:schemeClr': { '@_val': m[1] } }, ctx)
      : '#' + String(m[2]).toUpperCase()
    if (c) out.push(c)
  }
  return out.length ? out : [fallback]
}

/** Synthetic sp node for the diagram fallback: box in EMU, solid fill, optional text lines. */
function dgmSp(
  box: { x: number; y: number; cx: number; cy: number },
  fill: string | { spPr: any },
  lines: Array<{ text: string; lvl: number; sizePt: number; bold?: boolean }>,
  opts: {
    prst?: string
    textColor?: string
    align?: 'l' | 'ctr'
    anchor?: 't' | 'ctr'
    stroke?: string
    noFill?: boolean
    adj?: number
    /** xfrm rotation (60000ths of a degree) */
    rot?: number
  } = {},
): any {
  const fillNode = opts.noFill
    ? { 'a:noFill': {} }
    : typeof fill === 'string'
      ? { 'a:solidFill': { 'a:srgbClr': { '@_val': fill.replace('#', '') } } }
      : { 'a:solidFill': fill.spPr?.['a:solidFill'] }
  return {
    'p:spPr': {
      'a:xfrm': {
        ...(opts.rot ? { '@_rot': String(opts.rot) } : {}),
        'a:off': { '@_x': String(Math.round(box.x)), '@_y': String(Math.round(box.y)) },
        'a:ext': { '@_cx': String(Math.round(box.cx)), '@_cy': String(Math.round(box.cy)) },
      },
      'a:prstGeom': {
        '@_prst': opts.prst ?? 'rect',
        ...(opts.adj != null
          ? { 'a:avLst': { 'a:gd': { '@_name': 'adj', '@_fmla': 'val ' + Math.round(opts.adj) } } }
          : {}),
      },
      ...fillNode,
      ...(opts.stroke
        ? {
            'a:ln': {
              '@_w': '9525',
              'a:solidFill': { 'a:srgbClr': { '@_val': opts.stroke.replace('#', '') } },
            },
          }
        : {}),
    },
    ...(lines.length
      ? {
          'p:txBody': {
            'a:bodyPr': { '@_anchor': opts.anchor ?? 'ctr' },
            'a:p': lines.map((l) => ({
              'a:pPr': {
                '@_algn': opts.align ?? 'ctr',
                ...(l.lvl > 0
                  ? {
                      '@_marL': String(228600 * l.lvl),
                      '@_indent': '-114300',
                      'a:buChar': { '@_char': '\u2022' },
                    }
                  : { 'a:buNone': {} }),
              },
              'a:r': {
                'a:rPr': {
                  '@_sz': String(Math.round(l.sizePt * 100)),
                  ...(l.bold ? { '@_b': '1' } : {}),
                  'a:solidFill': {
                    'a:srgbClr': { '@_val': (opts.textColor ?? '#FFFFFF').replace('#', '') },
                  },
                },
                'a:t': l.text,
              },
            })),
          },
        }
      : {}),
  }
}

/** Tint toward white keeping pct of the color (fallback body panels). */
function dgmTint(hex: string, pct: number): string {
  const h = hex.replace('#', '')
  const mix = (i: number) =>
    Math.round(255 * (1 - pct) + parseInt(h.slice(i, i + 2), 16) * pct)
      .toString(16)
      .toUpperCase()
      .padStart(2, '0')
  return '#' + mix(0) + mix(2) + mix(4)
}

/**
 * SmartArt without a usable prerendered drawing: lay the data tree out ourselves.
 * Flat lists keep the PowerPoint-measured tile grid; hierarchical data dispatches on
 * the layout part's uniqueId to a small family of hand-written algorithms
 * (columns / table list / stacked list / hierarchy), defaulting to top-node tiles
 * with descendant bullet text.
 */
export function layoutDiagramFallback(
  dataXml: string,
  ctx: ParseContext,
  frameCx: number,
  frameCy: number,
  layoutId?: string,
  colorsXml?: string,
): SlideElement[] {
  let doc: any
  try {
    doc = parser.parse(dataXml)
  } catch {
    return []
  }
  const model = doc['dgm:dataModel']
  const ptsRaw = model?.['dgm:ptLst']?.['dgm:pt']
  const pts: any[] = Array.isArray(ptsRaw) ? ptsRaw : ptsRaw ? [ptsRaw] : []
  const cxnsRaw = model?.['dgm:cxnLst']?.['dgm:cxn']
  const cxns: any[] = Array.isArray(cxnsRaw) ? cxnsRaw : cxnsRaw ? [cxnsRaw] : []
  const docId = pts.find((p) => p?.['@_type'] === 'doc')?.['@_modelId']
  if (docId == null) return []
  // dgm:pt type defaults to 'node'; some writers emit it explicitly
  const nodePts = new Map(
    pts
      .filter((p) => p?.['@_type'] == null || p?.['@_type'] === 'node')
      .map((p) => [String(p['@_modelId']), p]),
  )
  // Plain parent-child connections (type absent or parOf), grouped by parent, ordered by srcOrd
  const bySrc = new Map<string, any[]>()
  for (const c of cxns) {
    const t = c?.['@_type']
    if (t != null && t !== 'parOf') continue
    if (!nodePts.has(String(c['@_destId']))) continue
    const k = String(c['@_srcId'])
    if (!bySrc.has(k)) bySrc.set(k, [])
    bySrc.get(k)!.push(c)
  }
  for (const arr of bySrc.values())
    arr.sort((a, b) => (parseInt(a['@_srcOrd'], 10) || 0) - (parseInt(b['@_srcOrd'], 10) || 0))
  const seen = new Set<string>()
  const build = (id: string): DgmTreeNode[] =>
    (bySrc.get(id) ?? [])
      .map((c) => String(c['@_destId']))
      .filter((d) => !seen.has(d) && (seen.add(d), true))
      .map((d) => {
        const pt = nodePts.get(d)
        return {
          id: d,
          texts: collectDgmTexts(pt),
          ...(pt?.['dgm:spPr']?.['a:solidFill'] ? { spPr: pt['dgm:spPr'] } : {}),
          children: build(d),
        }
      })
  const roots = build(String(docId))
  if (!roots.length) return []

  const colors = diagramCycleColors(colorsXml, ctx)
  const colorOf = (node: DgmTreeNode, i: number): string | { spPr: any } =>
    node.spPr ? { spPr: node.spPr } : colors[i % colors.length]!
  const hasHierarchy = roots.some((r) => r.children.length)
  const sps: any[] = []
  // Text sizes scale with the box and shrink with line count (SmartArt autofit, coarse)
  const fitSize = (boxCyEmu: number, nLines: number, cap = 26) => {
    const boxPt = boxCyEmu / 12700
    return Math.max(8, Math.min(cap, (boxPt * 0.82) / Math.max(nLines, 1) / 1.35))
  }
  // Width-aware variant: also shrink until the longest word fits the box width and
  // the wrapped line count fits the height (coarse 0.62em average char advance)
  const fitSizeW = (boxCyEmu: number, boxCxEmu: number, texts: string[], cap = 26) => {
    let s = fitSize(boxCyEmu, Math.max(texts.length, 1), cap)
    const boxWPt = (boxCxEmu / 12700) * 0.92
    const boxHPt = (boxCyEmu / 12700) * 0.9
    const longest = Math.max(0, ...texts.flatMap((t) => t.split(/\s+/).map((w) => w.length)))
    if (longest) s = Math.min(s, boxWPt / (longest * 0.62))
    for (let i = 0; i < 3; i++) {
      const lines = texts.reduce(
        (acc, t) => acc + Math.max(1, Math.ceil((t.length * 0.62 * s) / boxWPt)),
        0,
      )
      const need = lines * 1.35 * s
      if (need <= boxHPt) break
      s *= Math.sqrt(boxHPt / need)
    }
    return Math.max(6, s)
  }
  // Layouts our families still model worse than the neutral placeholder chip
  if (layoutId === 'cycle4') return []
  const byLayout =
    layoutId === 'hList1' || layoutId === 'hList2'
      ? 'columns'
      : layoutId === 'hList3'
        ? 'tableList'
        : layoutId != null && /^pList/.test(layoutId)
          ? 'pictureList'
          : layoutId != null && /^list/.test(layoutId)
            ? 'boxList'
            : layoutId === 'vList5' || (layoutId != null && /^Bracket/.test(layoutId))
              ? 'sideList'
              : layoutId === 'vProcess5'
                ? 'stepped'
                : layoutId != null && /^(vList|vProcess)/.test(layoutId)
                  ? 'stacked'
                  : layoutId != null && /^bList/.test(layoutId)
                    ? 'cards'
                    : layoutId != null && /^(process|hProcess|bProcess)/.test(layoutId)
                      ? 'procCards'
                      : layoutId != null && /^equation/.test(layoutId)
                        ? 'equation'
                        : layoutId != null && /^pyramid/.test(layoutId)
                          ? 'pyramid'
                          : layoutId != null && /^Picture/.test(layoutId)
                            ? 'strips'
                            : layoutId === 'chevron2'
                              ? 'chevronList'
                              : layoutId != null && /^chevron/.test(layoutId)
                                ? 'chevronRow'
                                : layoutId != null && /^(cycle[127]|radial)/.test(layoutId)
                                  ? 'cycle'
                                  : layoutId != null && /orgchart/i.test(layoutId)
                                    ? 'orgChart'
                                    : layoutId != null && /^hierarchy/.test(layoutId)
                                      ? 'hierarchy'
                                      : 'blocks'
  const family = byLayout === 'blocks' && !hasHierarchy ? 'flatGrid' : byLayout

  if (family === 'flatGrid') {
    // Flat list: tile grid (PowerPoint-measured GAP/aspect)
    const n = roots.length
    const GAP = 0.115
    const ASPECT = 0.6
    const availCy = frameCy * 0.98
    let cols = 1
    let best = 0
    for (let c = 1; c <= n; c++) {
      const r = Math.ceil(n / c)
      const w = Math.min(frameCx / (c + (c - 1) * GAP), availCy / (r * ASPECT + (r - 1) * GAP))
      if (w > best) {
        best = w
        cols = c
      }
    }
    const rows = Math.ceil(n / cols)
    const bw = best
    const bh = bw * ASPECT
    const gap = bw * GAP
    const gridW = cols * bw + (cols - 1) * gap
    const gridH = rows * bh + (rows - 1) * gap
    // PowerPoint centers the tile grid inside the frame on both axes (napierone 0005
    // p7 measured: width-bound 3x2 grid sits with equal ~10% top/bottom margins)
    const xOff = (frameCx - gridW) / 2
    const yOff = (frameCy - gridH) / 2
    roots.forEach((node, i) => {
      const row = Math.floor(i / cols)
      const inRow = row === rows - 1 ? n - (rows - 1) * cols : cols
      const col = i - row * cols
      const rowW = inRow * bw + (inRow - 1) * gap
      const x = xOff + (gridW - rowW) / 2 + col * (bw + gap)
      const y = yOff + row * (bh + gap)
      // PowerPoint's tile autofit leaves generous vertical padding: wrapped text fits
      // about half the tile height (napierone 0005 measured ~17pt in a 110pt tile);
      // short single-line text keeps the normal cap
      const t = fitSizeW(bh * 0.5, bw, node.texts)
      sps.push(
        dgmSp(
          { x, y, cx: bw, cy: bh },
          colorOf(node, i),
          node.texts.map((tx) => ({ text: tx, lvl: 0, sizePt: t })),
        ),
      )
    })
  } else if (family === 'equation') {
    // equation1: operand circles joined by + and = operator glyphs (a + b = c)
    const n = roots.length
    const OP = 0.42
    const d = Math.min(frameCy * 0.92, frameCx / (n + (n - 1) * OP))
    const opW = d * OP
    const totalW = n * d + (n - 1) * opW
    const x0 = (frameCx - totalW) / 2
    const yC = (frameCy - d) / 2
    const opColor = typeof colors[0] === 'string' ? colors[0] : '#4472C4'
    roots.forEach((node, i) => {
      const x = x0 + i * (d + opW)
      sps.push(
        dgmSp(
          { x, y: yC, cx: d, cy: d },
          colorOf(node, i),
          node.texts.map((tx) => ({
            text: tx,
            lvl: 0,
            sizePt: fitSizeW(d * 0.72, d * 0.78, node.texts),
          })),
          { prst: 'ellipse' },
        ),
      )
      if (i < n - 1) {
        const g = d * 0.3
        sps.push(
          dgmSp(
            { x: x + d + (opW - g) / 2, y: yC + (d - g) / 2, cx: g, cy: g },
            dgmTint(opColor, 0.25),
            [],
            { prst: i === n - 2 ? 'mathEqual' : 'mathPlus' },
          ),
        )
      }
    })
  } else if (family === 'columns') {
    // Per top node: colored header (title) + tinted body (descendant bullets);
    // column width caps at 0.62 x column height and the group centers (PPT hList look)
    const n = roots.length
    const gap = frameCy * 0.055
    const availCy = frameCy * 0.92
    const y0 = frameCy * 0.04
    const cw = Math.min((frameCx - gap * (n - 1)) / n, availCy * 0.52)
    const x0 = (frameCx - (cw * n + gap * (n - 1))) / 2
    const headCy = availCy * 0.18
    roots.forEach((node, i) => {
      const x = x0 + i * (cw + gap)
      const base = colorOf(node, i)
      const baseHex = typeof base === 'string' ? base : undefined
      const t = fitSize(headCy, Math.max(node.texts.length, 1), 20)
      sps.push(
        dgmSp(
          { x, y: y0, cx: cw, cy: headCy },
          base,
          node.texts.map((tx) => ({ text: tx, lvl: 0, sizePt: t })),
        ),
      )
      const bullets = dgmBulletLines(node)
      const bodyCy = availCy - headCy - frameCy * 0.01
      const b = fitSize(bodyCy, Math.max(bullets.length, 1), 16)
      sps.push(
        dgmSp(
          { x, y: y0 + headCy + frameCy * 0.01, cx: cw, cy: bodyCy },
          baseHex ? dgmTint(baseHex, 0.2) : base,
          bullets.map((l) => ({ text: l.text, lvl: l.lvl, sizePt: b })),
          { textColor: '#333333', align: 'l', anchor: 't' },
        ),
      )
    })
  } else if (family === 'tableList') {
    // Single parent banner on top, children as equal cells below
    const parent = roots[0]!
    const headCy = frameCy * 0.28
    sps.push(
      dgmSp(
        { x: 0, y: 0, cx: frameCx, cy: headCy },
        colorOf(parent, 0),
        parent.texts.map((t) => ({
          text: t,
          lvl: 0,
          sizePt: fitSize(headCy, Math.max(parent.texts.length, 1)),
        })),
      ),
    )
    const kids = parent.children.length ? parent.children : roots.slice(1)
    const n = Math.max(kids.length, 1)
    const gap = frameCx * 0.012
    const cw = (frameCx - gap * (n - 1)) / n
    kids.forEach((k, i) => {
      sps.push(
        dgmSp(
          { x: i * (cw + gap), y: headCy + frameCy * 0.012, cx: cw, cy: frameCy * 0.62 },
          colorOf(k, i + 1),
          k.texts.map((t) => ({
            text: t,
            lvl: 0,
            sizePt: fitSize(frameCy * 0.62, Math.max(k.texts.length, 1)),
          })),
        ),
      )
    })
    sps.push(
      dgmSp({ x: 0, y: frameCy * 0.945, cx: frameCx, cy: frameCy * 0.055 }, colorOf(parent, 0), []),
    )
  } else if (family === 'stacked') {
    // Rounded blocks stacked vertically: title + descendant bullets inside each
    const n = roots.length
    const gap = frameCy * 0.06
    const bh = (frameCy - gap * (n - 1)) / n
    roots.forEach((node, i) => {
      const bullets = dgmBulletLines(node)
      const nLines = node.texts.length + bullets.length
      const t = fitSize(bh, Math.max(nLines, 1))
      const lines = [
        ...node.texts.map((tx) => ({ text: tx, lvl: 0, sizePt: t, bold: true })),
        ...bullets.map((l) => ({ text: l.text, lvl: l.lvl, sizePt: t * 0.8 })),
      ]
      sps.push(
        dgmSp({ x: 0, y: i * (bh + gap), cx: frameCx, cy: bh }, colorOf(node, i), lines, {
          prst: 'roundRect',
          align: 'l',
          anchor: 'ctr',
        }),
      )
    })
  } else if (family === 'hierarchy') {
    // Parent tiles in a row; each parent's children side by side beneath it
    const n = roots.length
    const gap = frameCx * 0.03
    const colW = (frameCx - gap * (n - 1)) / n
    const parentCy = frameCy * 0.42
    const childCy = frameCy * 0.48
    roots.forEach((node, i) => {
      const x = i * (colW + gap)
      sps.push(
        dgmSp(
          { x: x + colW * 0.06, y: 0, cx: colW * 0.88, cy: parentCy },
          colorOf(node, i),
          node.texts.map((t) => ({
            text: t,
            lvl: 0,
            sizePt: fitSize(parentCy, Math.max(node.texts.length, 1)),
          })),
          { prst: 'roundRect' },
        ),
      )
      const kids = node.children
      if (!kids.length) return
      const kgap = colW * 0.04
      const kw = (colW - kgap * (kids.length - 1)) / kids.length
      kids.forEach((k, j) => {
        sps.push(
          dgmSp(
            { x: x + j * (kw + kgap), y: frameCy - childCy, cx: kw, cy: childCy },
            colorOf(k, i),
            k.texts.map((t) => ({
              text: t,
              lvl: 0,
              sizePt: fitSize(childCy, Math.max(k.texts.length, 1), 20),
            })),
            { prst: 'roundRect' },
          ),
        )
      })
    })
  } else if (family === 'sideList') {
    // Vertical Block List: compact rows top-aligned — accent title box left, light bullet panel right
    const n = roots.length
    const gap = frameCy * 0.045
    const ih = Math.min((frameCy * 0.98 - gap * (n - 1)) / n, frameCy * 0.2)
    roots.forEach((node, i) => {
      const y = frameCy * 0.02 + i * (ih + gap)
      const base = colorOf(node, i)
      const baseHex = typeof base === 'string' ? base : undefined
      const bullets = dgmBulletLines(node)
      if (bullets.length) {
        sps.push(
          dgmSp(
            { x: frameCx * 0.2, y, cx: frameCx * 0.38, cy: ih },
            baseHex ? dgmTint(baseHex, 0.16) : base,
            bullets.map((l) => ({
              text: l.text,
              lvl: l.lvl,
              sizePt: fitSize(ih, Math.max(bullets.length, 1), 15),
            })),
            { textColor: '#333333', align: 'l' },
          ),
        )
      }
      sps.push(
        dgmSp(
          { x: 0, y, cx: frameCx * 0.205, cy: ih },
          base,
          node.texts.map((tx) => ({
            text: tx,
            lvl: 0,
            sizePt: fitSize(ih, Math.max(node.texts.length, 1), 20),
          })),
          { prst: 'roundRect' },
        ),
      )
    })
  } else if (family === 'stepped') {
    // vProcess5 (step-down process): boxes stagger right as they descend, a small down
    // arrow tucked into the gap under each box's trailing edge
    const n = roots.length
    const gapY = frameCy * 0.055
    const bh = (frameCy - gapY * (n - 1)) / n
    const bw = frameCx * 0.62
    const stepX = n > 1 ? (frameCx - bw) / (n - 1) : 0
    roots.forEach((node, i) => {
      const x = i * stepX
      const y = i * (bh + gapY)
      const texts = [...node.texts, ...dgmBulletLines(node).map((l) => l.text)]
      const t = fitSizeW(bh, bw, texts, 16)
      sps.push(
        dgmSp(
          { x, y, cx: bw, cy: bh },
          colorOf(node, i),
          texts.map((tx) => ({ text: tx, lvl: 0, sizePt: t })),
          { align: 'l' },
        ),
      )
      if (i < n - 1) {
        const ah = gapY * 0.95
        const aw = ah * 1.1
        const base = colorOf(node, i)
        sps.push(
          dgmSp(
            { x: x + bw * 0.78 - aw / 2, y: y + bh + gapY * 0.025, cx: aw, cy: ah },
            typeof base === 'string' ? dgmTint(base, 0.45) : '#BFBFBF',
            [],
            { prst: 'downArrow' },
          ),
        )
      }
    })
  } else if (family === 'chevronList') {
    // chevron2 (vertical chevron list): a down-pointing chevron accent per row on the
    // left, the node's bullets in an outlined rounded card on the right
    const n = roots.length
    const gap = frameCy * 0.045
    const ih = (frameCy - gap * (n - 1)) / n
    const chW = Math.min(frameCx * 0.16, ih * 0.75)
    roots.forEach((node, i) => {
      const y = i * (ih + gap)
      const base = colorOf(node, i)
      const baseHex = typeof base === 'string' ? base : undefined
      // rotated 90deg: swap the box around the accent cell's center so it points down
      sps.push(
        dgmSp({ x: chW / 2 - ih / 2, y: y + ih / 2 - chW / 2, cx: ih, cy: chW }, base, [], {
          prst: 'chevron',
          rot: 5400000,
        }),
      )
      // label overlay stays horizontal (text inside the rotated shape would rotate)
      const t = fitSizeW(ih * 0.9, chW * 0.85, node.texts, 13)
      sps.push(
        dgmSp(
          { x: 0, y, cx: chW, cy: ih },
          '#FFFFFF',
          node.texts.map((tx) => ({ text: tx, lvl: 0, sizePt: t })),
          { noFill: true },
        ),
      )
      const bullets = dgmBulletLines(node)
      const b = fitSizeW(
        ih * 0.92,
        (frameCx - chW * 1.2) * 0.94,
        bullets.map((l) => l.text),
        13,
      )
      sps.push(
        dgmSp(
          { x: chW * 1.2, y, cx: frameCx - chW * 1.2, cy: ih },
          '#FFFFFF',
          bullets.map((l) => ({ text: l.text, lvl: l.lvl, sizePt: b })),
          {
            prst: 'roundRect',
            textColor: '#333333',
            align: 'l',
            stroke: baseHex ?? '#999999',
            adj: 10000,
          },
        ),
      )
    })
  } else if (family === 'chevronRow') {
    // One horizontal band of chevron arrows, vertically centered (chevron1)
    const flat = roots.flatMap((r) => [r, ...r.children])
    const n = flat.length
    const overlap = 0.18 // chevrons tuck into each other
    const ch = Math.min(frameCy * 0.34, (frameCx / (n - (n - 1) * overlap)) * 0.42)
    const cw = ch / 0.42
    const step = cw * (1 - overlap)
    const rowW = cw + step * (n - 1)
    const x0 = (frameCx - rowW) / 2
    const y = (frameCy - ch) / 2
    flat.forEach((node, i) => {
      sps.push(
        dgmSp(
          { x: x0 + i * step, y, cx: cw, cy: ch },
          colorOf(node, i),
          node.texts.map((tx) => ({
            text: tx,
            lvl: 0,
            sizePt: fitSize(ch, Math.max(node.texts.length, 1), 20),
          })),
          { prst: 'chevron' },
        ),
      )
    })
  } else if (family === 'cycle') {
    // Nodes on a circle; radial* = big center circle with satellite circles touching it
    const central = layoutId != null && /^radial/.test(layoutId)
    const ring = central ? (roots[0]!.children.length ? roots[0]!.children : roots.slice(1)) : roots
    const n = Math.max(ring.length, 1)
    const cxr = frameCx / 2
    const cyr = frameCy / 2
    if (central) {
      const minDim = Math.min(frameCx, frameCy)
      const R = minDim * 0.28
      const r = minDim * 0.17
      const base = colors[0] ?? '#4472C4'
      const light = dgmTint(base, 0.45)
      const circle = (node: DgmTreeNode, cx0: number, cy0: number, rad: number, cap: number) =>
        dgmSp(
          { x: cx0 - rad, y: cy0 - rad, cx: rad * 2, cy: rad * 2 },
          light,
          node.texts.map((tx) => ({
            text: tx,
            lvl: 0,
            sizePt: fitSize(rad * 2, Math.max(node.texts.length, 1) * 1.6, cap),
          })),
          { prst: 'ellipse', textColor: '#333333' },
        )
      ring.forEach((node, i) => {
        const ang = -Math.PI / 2 + (i * 2 * Math.PI) / n
        sps.push(
          circle(
            node,
            cxr + Math.cos(ang) * (R + r * 0.82),
            cyr + Math.sin(ang) * (R + r * 0.82),
            r,
            20,
          ),
        )
      })
      sps.push(circle(roots[0]!, cxr, cyr, R, 26))
    } else {
      const bw = Math.min(frameCx / 3.6, frameCy / 2.6)
      const bh = bw * 0.6
      const rx = frameCx / 2 - bw / 2
      const ry = frameCy / 2 - bh / 2
      ring.forEach((node, i) => {
        const ang = -Math.PI / 2 + (i * 2 * Math.PI) / n
        const x = cxr + Math.cos(ang) * rx
        const y = cyr + Math.sin(ang) * ry
        sps.push(
          dgmSp(
            { x: x - bw / 2, y: y - bh / 2, cx: bw, cy: bh },
            colorOf(node, i),
            node.texts.map((tx) => ({
              text: tx,
              lvl: 0,
              sizePt: fitSize(bh, Math.max(node.texts.length, 1), 18),
            })),
            { prst: 'roundRect' },
          ),
        )
      })
    }
  } else if (family === 'orgChart') {
    // Tidy-ish org tree: leaf-count spans, one row per depth
    // hierBranch 'init': PowerPoint hangs a branch when all its children are leaves (>=2)
    const hangs = (node: DgmTreeNode): boolean =>
      node.children.length >= 2 && node.children.every((c) => !c.children.length)
    const leaves = (node: DgmTreeNode): number =>
      hangs(node) || !node.children.length
        ? node.children.length
          ? 1.5
          : 1
        : node.children.reduce((a, c) => a + leaves(c), 0)
    const depthOf = (node: DgmTreeNode): number =>
      hangs(node)
        ? 1 + node.children.length
        : node.children.length
          ? 1 + Math.max(...node.children.map(depthOf))
          : 1
    const totalLeaves = roots.reduce((a, r) => a + leaves(r), 0)
    const depth = Math.max(...roots.map(depthOf))
    const rowCy = frameCy / depth
    const bh = Math.min(rowCy * 0.66, frameCy * 0.2)
    const slotW = frameCx / Math.max(totalLeaves, 1)
    const bw = Math.min(slotW * 0.94, frameCx / 5.5)
    const lineW = Math.max(frameCx * 0.0012, 9525)
    const lineColor = typeof colors[0] === 'string' ? colors[0]! : '#4472C4'
    const rowY = (d: number) => d * rowCy + (rowCy - bh) / 2
    const vline = (x: number, y1: number, y2: number) =>
      sps.push(dgmSp({ x: x - lineW / 2, y: y1, cx: lineW, cy: y2 - y1 }, lineColor, []))
    const hline = (x1: number, x2: number, y: number) =>
      sps.push(
        dgmSp(
          {
            x: Math.min(x1, x2) - lineW / 2,
            y: y - lineW / 2,
            cx: Math.abs(x2 - x1) + lineW,
            cy: lineW,
          },
          lineColor,
          [],
        ),
      )
    const box = (node: DgmTreeNode, cx: number, d: number) =>
      sps.push(
        dgmSp(
          { x: cx - bw / 2, y: rowY(d), cx: bw, cy: bh },
          colorOf(node, 0),
          node.texts.map((tx) => ({
            text: tx,
            lvl: 0,
            sizePt: fitSize(bh, Math.max(node.texts.length, 1), 20),
          })),
        ),
      )
    const place = (node: DgmTreeNode, slot0: number, d: number): number => {
      const span = leaves(node)
      const cx = (slot0 + span / 2) * slotW
      box(node, cx, d)
      if (hangs(node)) {
        // Hanging branch: children stack in rows below, trunk at the parent's left third
        const trunkX = cx - bw * 0.3
        vline(trunkX, rowY(d) + bh, rowY(d + node.children.length) + bh / 2)
        node.children.forEach((c, j) => {
          const cy0 = rowY(d + 1 + j)
          const ccx = cx + bw * 0.14
          hline(trunkX, ccx - bw / 2, cy0 + bh / 2)
          box(c, ccx, d + 1 + j)
        })
      } else {
        let sNext = slot0
        for (const c of node.children) {
          const cSpan = leaves(c)
          const ccx = (sNext + cSpan / 2) * slotW
          const midY = (rowY(d) + bh + rowY(d + 1)) / 2
          vline(cx, rowY(d) + bh, midY)
          hline(cx, ccx, midY)
          vline(ccx, midY, rowY(d + 1))
          sNext += place(c, sNext, d + 1)
        }
      }
      return span
    }
    let s0 = 0
    for (const r of roots) s0 += place(r, s0, 0)
  } else if (family === 'cards') {
    // bList cards: outlined white card with child bullets, accent footer strip with the parent title
    const n = roots.length
    const GAP = 0.14
    const ASPECT = 1.05
    const availCy = frameCy * 0.96
    let cols = 1
    let best = 0
    for (let c = 1; c <= n; c++) {
      const r = Math.ceil(n / c)
      const w = Math.min(frameCx / (c + (c - 1) * GAP), availCy / (r * ASPECT + (r - 1) * GAP))
      if (w > best) {
        best = w
        cols = c
      }
    }
    const rows = Math.ceil(n / cols)
    const bw = best
    const bh = bw * ASPECT
    const gap = bw * GAP
    const gridW = cols * bw + (cols - 1) * gap
    const gridH = rows * bh + (rows - 1) * gap
    const y00 = (frameCy - gridH) / 2
    roots.forEach((node, i) => {
      const row = Math.floor(i / cols)
      const col = i - row * cols
      const x = (frameCx - gridW) / 2 + col * (bw + gap)
      const y = y00 + row * (bh + gap)
      const base = colorOf(node, i)
      const baseHex = typeof base === 'string' ? base : '#4472C4'
      const bullets = dgmBulletLines(node)
      const footCy = bh * 0.28
      sps.push(
        dgmSp(
          { x, y, cx: bw, cy: bh - footCy },
          base,
          bullets.map((l) => ({
            text: l.text,
            lvl: l.lvl,
            sizePt: fitSize(bh - footCy, Math.max(bullets.length, 1) * 2, 14),
          })),
          { textColor: '#333333', align: 'l', anchor: 't', noFill: true, stroke: baseHex },
        ),
      )
      sps.push(
        dgmSp(
          { x, y: y + bh - footCy, cx: bw, cy: footCy },
          base,
          node.texts.map((tx) => ({
            text: tx,
            lvl: 0,
            sizePt: fitSize(footCy, Math.max(node.texts.length, 1), 12),
          })),
          { align: 'l' },
        ),
      )
      const r = footCy * 0.55
      sps.push(
        dgmSp(
          { x: x + bw - r * 1.6, y: y + bh - footCy - r * 0.45, cx: r * 2, cy: r * 2 },
          baseHex ? dgmTint(baseHex, 0.35) : base,
          [],
          { prst: 'ellipse' },
        ),
      )
    })
  } else if (family === 'procCards') {
    // process cards: items in a row — accent title box, outlined child panel offset
    // below-right, small arrow between items
    const n = roots.length
    const gapX = frameCx * 0.1
    const iw = (frameCx - gapX * (n - 1)) / n
    const ih = Math.min(frameCy * 0.5, iw * 1.1)
    const y00 = (frameCy - ih) / 2
    roots.forEach((node, i) => {
      const x = i * (iw + gapX)
      const base = colorOf(node, i)
      const baseHex = typeof base === 'string' ? base : '#4472C4'
      const titleCy = ih * 0.42
      sps.push(
        dgmSp(
          { x, y: y00, cx: iw * 0.62, cy: titleCy },
          base,
          node.texts.map((tx) => ({
            text: tx,
            lvl: 0,
            sizePt: fitSize(titleCy, Math.max(node.texts.length, 1), 18),
          })),
          { prst: 'roundRect', align: 'l' },
        ),
      )
      const bullets = dgmBulletLines(node)
      sps.push(
        dgmSp(
          { x: x + iw * 0.14, y: y00 + titleCy * 0.62, cx: iw * 0.66, cy: ih - titleCy * 0.62 },
          base,
          bullets.map((l) => ({
            text: l.text,
            lvl: l.lvl,
            sizePt: fitSize(ih - titleCy, Math.max(bullets.length, 1) * 1.6, 15),
          })),
          {
            prst: 'roundRect',
            textColor: '#333333',
            align: 'l',
            anchor: 't',
            noFill: true,
            stroke: baseHex,
          },
        ),
      )
      if (i < n - 1) {
        const aw = gapX * 0.55
        sps.push(
          dgmSp(
            { x: x + iw + (gapX - aw) / 2, y: y00 + titleCy * 0.28, cx: aw, cy: titleCy * 0.45 },
            baseHex ? dgmTint(baseHex, 0.45) : base,
            [],
            { prst: 'rightArrow' },
          ),
        )
      }
    })
  } else if (family === 'pyramid') {
    // Pyramid: single node = full triangle; several = triangle apex + widening trapezoid rows
    const items = roots
    const n = items.length
    const w0 = Math.min(frameCx * 0.62, frameCy * 1.05)
    const x0 = (frameCx - w0) / 2
    const rowH = frameCy / n
    // Side edges stay collinear: each trapezoid's per-side inset is w0/(2n),
    // expressed through the preset's adj (inset = min(w,h) * adj / 100000)
    items.forEach((node, i) => {
      const botW = (w0 * (i + 1)) / n
      const lines = node.texts.map((tx) => ({
        text: tx,
        lvl: 0,
        sizePt: fitSize(rowH, Math.max(node.texts.length, 1), 22),
      }))
      if (i === 0) {
        sps.push(
          dgmSp({ x: x0 + (w0 - botW) / 2, y: 0, cx: botW, cy: rowH }, colorOf(node, i), lines, {
            prst: 'triangle',
            textColor: '#333333',
          }),
        )
      } else {
        const inset = w0 / (2 * n)
        const adj = (inset / Math.min(botW, rowH)) * 100000
        sps.push(
          dgmSp(
            { x: x0 + (w0 - botW) / 2, y: i * rowH, cx: botW, cy: rowH },
            colorOf(node, i),
            lines,
            { prst: 'trapezoid', textColor: '#333333', adj },
          ),
        )
      }
    })
  } else if (family === 'strips') {
    // Picture strips: stacked outlined rows with a large dark label
    const n = roots.length
    const gap = frameCy * 0.05
    const ih = (frameCy - gap * (n - 1)) / n
    const x = frameCx * 0.2
    const w = frameCx * 0.45
    roots.forEach((node, i) => {
      const base = colorOf(node, i)
      const baseHex = typeof base === 'string' ? base : '#4472C4'
      sps.push(
        dgmSp(
          { x, y: i * (ih + gap), cx: w, cy: ih },
          base,
          node.texts.map((tx) => ({
            text: tx,
            lvl: 0,
            sizePt: fitSize(ih, Math.max(node.texts.length, 1) * 1.4, 24),
          })),
          { textColor: '#333333', align: 'l', noFill: true, stroke: baseHex },
        ),
      )
    })
  } else if (family === 'pictureList') {
    // Picture List: a light container with per-item picture boxes on top,
    // per-item text columns (title + child bullets) below
    const n = roots.length
    const contCy = frameCy * 0.44
    const cont = colors[0] ? dgmTint(colors[0], 0.25) : '#DCE3F2'
    sps.push(dgmSp({ x: 0, y: 0, cx: frameCx, cy: contCy }, cont, [], { prst: 'roundRect' }))
    const gap = frameCx * 0.03
    const cw = (frameCx - gap * (n - 1) - frameCx * 0.04) / n
    roots.forEach((node, i) => {
      const x = frameCx * 0.02 + i * (cw + gap)
      sps.push(
        dgmSp(
          { x, y: contCy * 0.12, cx: cw, cy: contCy * 0.76 },
          colors[0] ? dgmTint(colors[0], 0.12) : '#E8ECF7',
          [],
          { prst: 'roundRect' },
        ),
      )
      const bullets = dgmBulletLines(node)
      const bodyCy = frameCy - contCy - frameCy * 0.02
      const t = fitSize(bodyCy, Math.max(node.texts.length + bullets.length, 1), 18)
      sps.push(
        dgmSp(
          { x, y: contCy + frameCy * 0.02, cx: cw, cy: bodyCy },
          colorOf(node, i),
          [
            ...node.texts.map((tx) => ({ text: tx, lvl: 0, sizePt: t })),
            ...bullets.map((l) => ({ text: l.text, lvl: l.lvl, sizePt: t * 0.85 })),
          ],
          { prst: 'roundRect', align: 'l', anchor: 't' },
        ),
      )
    })
  } else if (family === 'boxList') {
    // Vertical Box List: per item a rounded title box upper-left + a light body panel below
    const n = roots.length
    const gap = frameCy * 0.05
    const ih = (frameCy - gap * (n - 1)) / n
    roots.forEach((node, i) => {
      const y = i * (ih + gap)
      const base = colorOf(node, i)
      const baseHex = typeof base === 'string' ? base : undefined
      const titleCy = ih * 0.45
      sps.push(
        dgmSp(
          { x: frameCx * 0.03, y, cx: frameCx * 0.55, cy: titleCy },
          base,
          node.texts.map((tx) => ({
            text: tx,
            lvl: 0,
            sizePt: fitSize(titleCy, Math.max(node.texts.length, 1), 20),
          })),
          { prst: 'roundRect', align: 'l' },
        ),
      )
      const bullets = dgmBulletLines(node)
      const bodyCy = ih - titleCy * 0.55
      sps.push(
        dgmSp(
          { x: 0, y: y + titleCy * 0.55, cx: frameCx * 0.72, cy: bodyCy },
          base,
          bullets.map((l) => ({
            text: l.text,
            lvl: l.lvl,
            sizePt: fitSize(bodyCy, Math.max(bullets.length, 1), 16),
          })),
          {
            textColor: '#333333',
            align: 'l',
            anchor: 't',
            noFill: true,
            stroke: baseHex ?? '#4472C4',
          },
        ),
      )
    })
  } else {
    // blocks: top-node tile grid with descendant bullet text inside each tile
    const n = roots.length
    const GAP = 0.115
    const ASPECT = 0.62
    const availCy = frameCy * 0.98
    let cols = 1
    let best = 0
    for (let c = 1; c <= n; c++) {
      const r = Math.ceil(n / c)
      const w = Math.min(frameCx / (c + (c - 1) * GAP), availCy / (r * ASPECT + (r - 1) * GAP))
      if (w > best) {
        best = w
        cols = c
      }
    }
    const rows = Math.ceil(n / cols)
    const bw = best
    const bh = bw * ASPECT
    const gap = bw * GAP
    const gridW = cols * bw + (cols - 1) * gap
    const gridH = rows * bh + (rows - 1) * gap
    // PowerPoint centers the tile grid inside the frame on both axes (napierone 0005
    // p7 measured: width-bound 3x2 grid sits with equal ~10% top/bottom margins)
    const xOff = (frameCx - gridW) / 2
    const yOff = (frameCy - gridH) / 2
    roots.forEach((node, i) => {
      const row = Math.floor(i / cols)
      const inRow = row === rows - 1 ? n - (rows - 1) * cols : cols
      const col = i - row * cols
      const rowW = inRow * bw + (inRow - 1) * gap
      const x = xOff + (gridW - rowW) / 2 + col * (bw + gap)
      const y = yOff + row * (bh + gap)
      const bullets = dgmBulletLines(node)
      const t = fitSizeW(bh, bw, [...node.texts, ...bullets.map((l) => l.text)])
      const lines = [
        ...node.texts.map((tx) => ({ text: tx, lvl: 0, sizePt: t })),
        ...bullets.map((l) => ({ text: l.text, lvl: l.lvl, sizePt: t * 0.8 })),
      ]
      sps.push(
        dgmSp({ x, y, cx: bw, cy: bh }, colorOf(node, i), lines, { align: 'l', anchor: 'ctr' }),
      )
    })
  }

  const anchor: ByteAnchor = { spIndex: -1, originalXml: '', range: [0, 0] }
  const out: SlideElement[] = []
  for (const sp of sps) {
    const el = parseSpShape(sp, anchor, ctx)
    if (el.type !== 'passthrough') out.push(el)
  }
  return out
}

/** Text of one dgm:pt (paragraph strings of dgm:t). */
function collectDgmTexts(pt: any): string[] {
  const body = pt?.['dgm:t']
  if (!body || typeof body !== 'object') return []
  const paras: any[] = Array.isArray(body['a:p']) ? body['a:p'] : body['a:p'] ? [body['a:p']] : []
  const out: string[] = []
  for (const p of paras) {
    const runs: any[] = Array.isArray(p?.['a:r']) ? p['a:r'] : p?.['a:r'] ? [p['a:r']] : []
    const t = runs
      .map((r) => {
        const v = r?.['a:t']
        return typeof v === 'string' ? v : String(v?.['#text'] ?? '')
      })
      .join('')
    if (t) out.push(t)
  }
  return out
}

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
  // <a:tblBg>: direct fill, or a fillRef instantiated from the theme fill styles
  let bgFill = styleDef?.tblBg
  if (!bgFill && styleDef?.tblBgRef) {
    const { idx, phClr } = styleDef.tblBgRef
    const tpl =
      idx > 1000 ? ctx.theme?.bgFillStyles?.[idx - 1001] : ctx.theme?.fillStyles?.[idx - 1]
    bgFill = tpl
      ? parseFill(tpl, { ...ctx, phClr, mediaRels: ctx.themeMediaRels ?? ctx.mediaRels })
      : undefined
    if (!bgFill && phClr) bgFill = { type: 'solid', color: phClr }
  }
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
    ...(bgFill && bgFill.type !== 'none' ? { bgFill } : {}),
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

/** <a:duotone>: two colors mapping image luminance dark→light (theme texture backgrounds). */
function parseDuotone(blipNode: any, ctx: ParseContext): [string, string] | undefined {
  const duoRaw = blipNode?.['a:duotone']
  if (!duoRaw) return undefined
  const clrs: Array<{ c: string; tag: string }> = []
  for (const tag of ['a:schemeClr', 'a:srgbClr', 'a:prstClr', 'a:sysClr']) {
    const raw = duoRaw[tag]
    const arr: any[] = Array.isArray(raw) ? raw : raw ? [raw] : []
    for (const c of arr) {
      const resolved = resolveColorNode({ [tag]: c }, ctx)
      if (resolved) clrs.push({ c: resolved, tag })
    }
  }
  if (clrs.length < 2) return undefined
  let pair = clrs.slice(0, 2)
  // The parser keeps order within one tag type but loses it across types, and
  // Office's standard picture duotone mixes types (<a:prstClr black/> +
  // <a:schemeClr accent/>): restore dark-to-light by luminance in that case
  if (pair[0]!.tag !== pair[1]!.tag) {
    const lum = (h: string) =>
      0.299 * parseInt(h.slice(1, 3), 16) +
      0.587 * parseInt(h.slice(3, 5), 16) +
      0.114 * parseInt(h.slice(5, 7), 16)
    pair = [...pair].sort((a, b) => lum(a.c) - lum(b.c))
  }
  return [pair[0]!.c, pair[1]!.c]
}

/** <a:clrChange>: replace clrFrom pixels with clrTo (alpha 0 clrTo = color-to-transparent, common on metafiles). */
function parseClrChange(
  blipNode: any,
  ctx: ParseContext,
): { from: string; to: string } | undefined {
  const cc = blipNode?.['a:clrChange']
  if (!cc) return undefined
  const from = resolveColorNode(cc['a:clrFrom'], ctx)
  const to = resolveColorNode(cc['a:clrTo'], ctx)
  return from && to ? { from, to } : undefined
}

function parseFill(spPr: any, ctx: ParseContext): Fill | undefined {
  if (!spPr) return undefined
  if ('a:noFill' in spPr) return { type: 'none' }
  // <a:grpFill/>: inherit the enclosing group's grpSpPr fill
  if ('a:grpFill' in spPr) return ctx.groupFill

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
    if (mediaRef) {
      const alphaAmt = blip['a:blip']?.['a:alphaModFix']?.['@_amt']
      const alpha =
        alphaAmt != null ? Math.max(0, Math.min(1, parseInt(alphaAmt, 10) / 100000)) : undefined
      const duotone = parseDuotone(blip['a:blip'], ctx)
      const clrChange = parseClrChange(blip['a:blip'], ctx)
      const fr = blip['a:stretch']?.['a:fillRect']
      const pct = (v: unknown) => (v != null ? (parseInt(String(v), 10) || 0) / 100000 : 0)
      const fillRect =
        fr && (fr['@_l'] != null || fr['@_t'] != null || fr['@_r'] != null || fr['@_b'] != null)
          ? { l: pct(fr['@_l']), t: pct(fr['@_t']), r: pct(fr['@_r']), b: pct(fr['@_b']) }
          : undefined
      const tl = blip['a:tile']
      const tile =
        tl && typeof tl === 'object'
          ? {
              tx: intOr(tl['@_tx'], 0),
              ty: intOr(tl['@_ty'], 0),
              sx: intOr(tl['@_sx'], 100000) / 100000,
              sy: intOr(tl['@_sy'], 100000) / 100000,
              algn: String(tl['@_algn'] ?? 'tl'),
            }
          : undefined
      return {
        type: 'image',
        mediaRef,
        mode: 'a:tile' in blip ? 'tile' : 'stretch',
        ...(alpha != null && alpha < 1 ? { alpha } : {}),
        ...(fillRect ? { fillRect } : {}),
        ...(duotone ? { duotone } : {}),
        ...(clrChange ? { clrChange } : {}),
        ...(tile ? { tile } : {}),
      }
    }
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
  const ftr = grad['a:path']?.['a:fillToRect']
  // Omitted fillToRect attributes default to 0 (whole tile rect), not to a centered inset
  const frac = (v: unknown) => (v != null ? (parseInt(String(v), 10) || 0) / 100000 : 0)
  return {
    type: 'gradient',
    stops,
    ...(angle != null ? { angle } : {}),
    ...(pathType === 'circle' || pathType === 'rect' || pathType === 'shape'
      ? { path: pathType }
      : {}),
    ...(ftr
      ? {
          fillTo: {
            l: frac(ftr['@_l']),
            t: frac(ftr['@_t']),
            r: frac(ftr['@_r']),
            b: frac(ftr['@_b']),
          },
        }
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
    ...(intOr(bodyPr['@_numCol'], 1) > 1
      ? { numCol: intOr(bodyPr['@_numCol'], 1), spcCol: intOr(bodyPr['@_spcCol'], 0) }
      : {}),
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
    const startAt = parseInt(pPr['a:buAutoNum']['@_startAt'], 10)
    if (Number.isFinite(startAt) && startAt > 1) bullet.startAt = startAt
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
  // (the master bodyStyle's buChar/marL/indent is where classic-template body bullets come from).
  // No field-wise merge: a paragraph redefining its bullet resets unspecified buClr/buSzPct/buFont
  // to follow the text (buClrTx/buSzTx/buFontTx semantics), not the chain's values.
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
  // Text highlight <a:highlight> (PowerPoint draws it as a background behind the run)
  const highlightNode = rPr['a:highlight']
  const highlight = highlightNode ? resolveColorNode(highlightNode, ctx) : undefined
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
  const runShadow = parseShadow(rPr, ctx) ?? dflt?.shadow
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
    ...(() => {
      const cap = rPr['@_cap'] != null ? String(rPr['@_cap']) : dflt?.cap
      return cap && cap !== 'none' ? { cap } : {}
    })(),
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
    ...(highlight ? { highlight } : {}),
    ...(outline ? { outline } : {}),
    ...(runShadow ? { shadow: runShadow } : {}),
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
  /** Skip non-placeholder shapes (showMasterSp="0"), keeping only the hf placeholders */
  hideShapes?: boolean
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
    } else if (/<p:ph[\s/>]/.test(fragXml) || opts.hideShapes) {
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
