/**
 * Single resolver for effective paragraph / run properties (ECMA-376 §17.7.2):
 * builtin -> docDefaults -> table style (whole table, then conditions) ->
 * paragraph style chain -> numbering level pPr -> direct. Runs continue with
 * docDefaults -> table-style rPr -> paragraph style -> character style -> direct,
 * toggle properties XOR across the style layers, the Cs twins selected by the
 * rtl chain. The raw tri-state readers (extractParaFormat, styleDisplayOf)
 * stay where they are; this module only composes them.
 */
import { attrsOf, findChild, nameOf, xmlParser, type XNode } from './xml-utils'
import { autoColorOf, colorFrom, lineTwipsOf, onOffOf, stripHash } from './parse-xml-text'
import {
  JC_ALIGN,
  activeCharIndents,
  autoSpaceFlagsOf,
  autoSpaceMode,
  charIndentsOf,
  indentTwipsOf,
  mergeCharIndents,
  paraBorderSidesOf,
  paraBordersOf,
  resolveCharIndents,
  shdDisplayFill,
  tabStopsOf,
  themedRFonts,
  type CharUnits,
} from './parse-props'
import { runBorderOf, styleRunFormat } from './parse-styles'
import type {
  CharIndents,
  DocDefaults,
  NumberingDef,
  NumberingLevel,
  ParaAlign,
  ParaBorderSides,
  ParaFormat,
  Run,
  StyleDisplay,
  StyleInfo,
  TabStop,
  TableCondFormat,
  ThemeColors,
  ThemeFonts,
} from './types'

export type PropSource =
  'direct' | 'numbering' | 'paraStyle' | 'charStyle' | 'tableStyle' | 'docDefaults' | 'builtin'

export type TableCond = 'firstRow' | 'lastRow' | 'firstCol' | 'lastCol'

export interface ResolveCompat {
  /** settings.xml compatibilityMode (0 = legacy layout) */
  mode: number
  /** w:doNotUseHTMLParagraphAutoSpacing: before/after spacing adds instead of collapsing */
  additiveParaSpacing: boolean
  /** w:doNotUseIndentAsNumberingTabStop */
  indentNotNumberingTabStop: boolean
}

export interface ResolveCtx {
  styles: Map<string, StyleInfo>
  defaultParaStyle?: StyleInfo
  defaultCharStyle?: StyleInfo
  docDefaults?: DocDefaults
  numbering: Map<string, NumberingDef>
  themeFonts?: ThemeFonts | null
  themeColors?: ThemeColors | null
  compat: ResolveCompat
  /** the table a cell paragraph sits in (conditions in ascending precedence: the last one wins) */
  table?: { styleId?: string; cond: TableCond[] }
  /** per-ctx caches (filled lazily) */
  cache?: WeakMap<XNode, { key: string; eff: EffectivePPr }>
  styleMemo?: Map<string, PLayer>
}

export interface LineSpacing {
  rule: 'auto' | 'atLeast' | 'exact'
  raw: number
}

export interface EffectivePPr {
  styleId?: string
  align?: ParaAlign
  bidi: boolean
  indent: { left: number; right: number; firstLine: number; chars?: CharIndents }
  spacing: {
    before: number
    after: number
    beforeAuto: boolean
    afterAuto: boolean
    line?: LineSpacing
  }
  keepNext: boolean
  keepLines: boolean
  pageBreakBefore: boolean
  widowControl: boolean
  contextualSpacing: boolean
  suppressLineNumbers: boolean
  snapToGrid: boolean
  autoSpaceDE: boolean
  autoSpaceDN: boolean
  wordWrap: boolean
  overflowPunct: boolean
  tabStops: TabStop[]
  borders?: ParaBorderSides
  shading?: string
  eastAsiaLang?: string
  list?: { numId: string; ilvl: number; level?: NumberingLevel }
  /** paragraph-mark rPr resolved through the same chain */
  mark: EffectiveRPr
  /** the mark is hidden (direct / style / default-style w:vanish, specVanish excluded) */
  hidden: boolean
  source: Record<string, PropSource>
}

export interface EffectiveRPr {
  styleId?: string
  cs: boolean
  bold: boolean
  italic: boolean
  sizeHalfPoints: number
  fonts: { ascii?: string; eastAsia?: string; cs?: string; eaSlotEmpty?: boolean }
  color?: string
  underline: boolean
  strike: boolean
  caps: 'none' | 'all' | 'small'
  vanish: boolean
  charSpacingTwips: number
  kernHalfPoints?: number
  vertAlign?: 'superscript' | 'subscript'
  bdr?: Run['bdr']
  shading?: string
  eastAsiaLang?: string
  source: Record<string, PropSource>
}

export interface ResolveParaOptions {
  /** twips per character for *Chars indents; absent = character indents stay unresolved */
  charUnits?: CharUnits
}

/** one layer of paragraph properties, every field tri-state */
export interface PLayer {
  align?: ParaAlign
  bidi?: boolean
  indentLeft?: number
  indentRight?: number
  indentFirstLine?: number
  indentChars?: CharIndents
  spaceBefore?: number
  spaceAfter?: number
  spaceBeforeAuto?: boolean
  spaceAfterAuto?: boolean
  line?: LineSpacing
  keepNext?: boolean
  keepLines?: boolean
  pageBreakBefore?: boolean
  widowControl?: boolean
  contextualSpacing?: boolean
  suppressLineNumbers?: boolean
  snapToGrid?: boolean
  autoSpaceDE?: boolean
  autoSpaceDN?: boolean
  wordWrap?: boolean
  overflowPunct?: boolean
  tabStops?: TabStop[]
  borders?: ParaBorderSides
  shading?: string
  eastAsiaLang?: string
}

/** one layer of run properties, every field tri-state */
interface RLayer {
  bold?: boolean
  italic?: boolean
  sizeHalfPoints?: number
  rtl?: boolean
  ascii?: string
  eastAsia?: string
  cs?: string
  eaSlotEmpty?: boolean
  color?: string
  underline?: boolean
  strike?: boolean
  caps?: 'none' | 'all' | 'small'
  vanish?: boolean
  charSpacingTwips?: number
  kernHalfPoints?: number
  vertAlign?: 'superscript' | 'subscript' | 'baseline'
  bdr?: Run['bdr'] | null
  shading?: string
  eastAsiaLang?: string
}

const WORD_DEFAULT_SIZE_HALF_POINTS = 20
const TOGGLES = ['bold', 'italic', 'strike', 'vanish'] as const

// ---------------------------------------------------------------- layers

function docDefaultsPLayer(d: DocDefaults | undefined): PLayer {
  if (!d) return {}
  const out: PLayer = {}
  if (d.widowControl !== undefined) out.widowControl = d.widowControl
  if (d.bidi !== undefined) out.bidi = d.bidi
  if (d.indentLeftTwips !== undefined) out.indentLeft = d.indentLeftTwips
  if (d.indentRightTwips !== undefined) out.indentRight = d.indentRightTwips
  if (d.indentFirstLineTwips !== undefined) out.indentFirstLine = d.indentFirstLineTwips
  if (d.indentChars) out.indentChars = d.indentChars
  if (d.spaceBeforeTwips !== undefined) out.spaceBefore = d.spaceBeforeTwips
  if (d.spaceAfterTwips !== undefined) out.spaceAfter = d.spaceAfterTwips
  if (d.spaceBeforeAuto !== undefined) out.spaceBeforeAuto = d.spaceBeforeAuto
  if (d.spaceAfterAuto !== undefined) out.spaceAfterAuto = d.spaceAfterAuto
  const line = lineOf(d.lineRule, d.lineRawTwips, d.lineSpacing)
  if (line) out.line = line
  if (d.borderSides) out.borders = d.borderSides
  if (d.eastAsiaLang) out.eastAsiaLang = d.eastAsiaLang
  return out
}

function lineOf(
  rule: 'auto' | 'atLeast' | 'exact' | undefined,
  raw: number | undefined,
  multiple: number | undefined,
): LineSpacing | undefined {
  if (raw !== undefined) return { rule: rule ?? 'auto', raw }
  if (multiple !== undefined) return { rule: 'auto', raw: Math.round(multiple * 240) }
  return undefined
}

function tableStylePLayer(ctx: ResolveCtx): PLayer {
  const t = tableStyleOf(ctx)?.tableDisplay
  if (!t) return {}
  const out: PLayer = {}
  const apply = (c: Pick<TableCondFormat, 'paraSpacing' | 'paraJc'> | undefined) => {
    if (!c) return
    const ps = c.paraSpacing
    if (ps) {
      if (ps.beforeTwips !== undefined) out.spaceBefore = ps.beforeTwips
      if (ps.afterTwips !== undefined) out.spaceAfter = ps.afterTwips
      const line = lineOf(ps.lineRule, ps.lineRawTwips, ps.lineSpacing)
      if (line) out.line = line
    }
    if (c.paraJc && JC_ALIGN[c.paraJc]) out.align = JC_ALIGN[c.paraJc]
  }
  apply(t)
  for (const cond of ctx.table?.cond ?? []) apply(t[cond])
  return out
}

function levelPLayer(level: NumberingLevel | undefined): PLayer {
  if (!level) return {}
  const out: PLayer = {}
  if (level.indentLeft !== undefined) out.indentLeft = level.indentLeft
  if (level.indentRight !== undefined) out.indentRight = level.indentRight
  if (level.hanging !== undefined && level.hanging > 0) out.indentFirstLine = -level.hanging
  else if (level.firstLine !== undefined && level.firstLine > 0)
    out.indentFirstLine = level.firstLine
  else if (level.hanging !== undefined || level.firstLine !== undefined) out.indentFirstLine = 0
  if (level.bidi !== undefined) out.bidi = level.bidi
  return out
}

function stylePLayer(ctx: ResolveCtx, info: StyleInfo | undefined): PLayer {
  if (!info) return {}
  const memo = (ctx.styleMemo ??= new Map())
  const hit = memo.get(info.styleId)
  if (hit) return hit
  const d = info.display
  const out: PLayer = {}
  if (d) {
    if (d.align) out.align = d.align
    if (d.bidi !== undefined) out.bidi = d.bidi
    if (d.indentLeftTwips !== undefined) out.indentLeft = d.indentLeftTwips
    if (d.indentRightTwips !== undefined) out.indentRight = d.indentRightTwips
    if (d.indentFirstLineTwips !== undefined) out.indentFirstLine = d.indentFirstLineTwips
    if (d.indentChars) out.indentChars = d.indentChars
    if (d.spaceBeforeTwips !== undefined) out.spaceBefore = d.spaceBeforeTwips
    if (d.spaceAfterTwips !== undefined) out.spaceAfter = d.spaceAfterTwips
    if (d.spaceBeforeAuto !== undefined) out.spaceBeforeAuto = d.spaceBeforeAuto
    if (d.spaceAfterAuto !== undefined) out.spaceAfterAuto = d.spaceAfterAuto
    const line = lineOf(d.lineRule, d.lineRawTwips, d.lineSpacing)
    if (line) out.line = line
    for (const k of [
      'keepNext',
      'keepLines',
      'pageBreakBefore',
      'widowControl',
      'contextualSpacing',
      'suppressLineNumbers',
      'autoSpaceDE',
      'autoSpaceDN',
      'wordWrap',
      'overflowPunct',
    ] as const) {
      if (d[k] !== undefined) out[k] = d[k]
    }
    if (d.tabStops) out.tabStops = d.tabStops
    if (d.borderSides) out.borders = d.borderSides
    if (d.shadingFill) out.shading = d.shadingFill
    if (d.eastAsiaLang) out.eastAsiaLang = d.eastAsiaLang
  }
  memo.set(info.styleId, out)
  return out
}

/** the direct w:pPr as a layer (extractParaFormat's reading, jc kept logical) */
function directPLayer(pPr: XNode | undefined, theme: ThemeColors | null | undefined): PLayer {
  if (!pPr) return {}
  const out: PLayer = {}
  const bidi = onOffOf(pPr, 'w:bidi')
  if (bidi !== undefined) out.bidi = bidi
  const jc = attrsOf(findChild(pPr, 'w:jc') ?? {})['w:val']
  if (jc && JC_ALIGN[jc]) out.align = JC_ALIGN[jc]
  const spacing = findChild(pPr, 'w:spacing')
  if (spacing) {
    const a = attrsOf(spacing)
    const rule = (a['w:lineRule'] ?? 'auto') as LineSpacing['rule']
    const line = lineTwipsOf(a['w:line'])
    if (line > 0 || (line === 0 && rule === 'atLeast')) out.line = { rule, raw: line }
    const autoOf = (v: string | undefined) =>
      v === undefined ? undefined : v === '1' || v === 'true'
    const ba = autoOf(a['w:beforeAutospacing'])
    if (ba !== undefined) out.spaceBeforeAuto = ba
    const aa = autoOf(a['w:afterAutospacing'])
    if (aa !== undefined) out.spaceAfterAuto = aa
    const before = parseInt(a['w:before'] ?? '', 10)
    if (before >= 0) out.spaceBefore = before
    const after = parseInt(a['w:after'] ?? '', 10)
    if (after >= 0) out.spaceAfter = after
  }
  const ind = findChild(pPr, 'w:ind')
  if (ind) {
    const { left, right, firstLine } = indentTwipsOf(ind)
    if (left !== undefined) out.indentLeft = left
    if (right !== undefined) out.indentRight = right
    if (firstLine !== undefined) out.indentFirstLine = firstLine
    const chars = charIndentsOf(ind)
    if (chars) out.indentChars = chars
  }
  for (const [k, tag] of [
    ['pageBreakBefore', 'w:pageBreakBefore'],
    ['keepNext', 'w:keepNext'],
    ['keepLines', 'w:keepLines'],
    ['suppressLineNumbers', 'w:suppressLineNumbers'],
    ['widowControl', 'w:widowControl'],
    ['contextualSpacing', 'w:contextualSpacing'],
    ['snapToGrid', 'w:snapToGrid'],
    ['wordWrap', 'w:wordWrap'],
    ['overflowPunct', 'w:overflowPunct'],
  ] as const) {
    const v = onOffOf(pPr, tag)
    if (v !== undefined) out[k] = v
  }
  const auto = autoSpaceFlagsOf(pPr)
  if (auto.de !== undefined) out.autoSpaceDE = auto.de
  if (auto.dn !== undefined) out.autoSpaceDN = auto.dn
  const shd = findChild(pPr, 'w:shd')
  if (shd) {
    const fill = attrsOf(shd)['w:fill']
    const display = shdDisplayFill(shd, theme)
    out.shading = display ?? (fill && fill !== 'auto' ? stripHash(fill) : '')
  }
  const sides = paraBorderSidesOf(pPr, theme)
  if (sides) out.borders = sides
  const stops = tabStopsOf(pPr)
  if (stops) out.tabStops = stops
  return out
}

// ---------------------------------------------------------------- merge

function mergeTabStops(base: TabStop[], over: TabStop[]): TabStop[] {
  const kept = base.filter((s) => !over.some((o) => o.pos === s.pos))
  const merged = [...kept, ...over.filter((s) => s.val !== 'clear')]
  return merged.sort((a, b) => a.pos - b.pos)
}

function mergeBorders(base: ParaBorderSides | undefined, over: ParaBorderSides): ParaBorderSides {
  const out: ParaBorderSides = { ...(base ?? {}) }
  for (const ch of ['t', 'b', 'l', 'r'] as const) {
    if (over[ch] !== undefined) out[ch] = over[ch]
  }
  return out
}

const PLAIN_PKEYS = [
  'align',
  'bidi',
  'indentLeft',
  'indentRight',
  'indentFirstLine',
  'spaceBefore',
  'spaceAfter',
  'spaceBeforeAuto',
  'spaceAfterAuto',
  'line',
  'keepNext',
  'keepLines',
  'pageBreakBefore',
  'widowControl',
  'contextualSpacing',
  'suppressLineNumbers',
  'snapToGrid',
  'autoSpaceDE',
  'autoSpaceDN',
  'wordWrap',
  'overflowPunct',
  'shading',
  'eastAsiaLang',
] as const

function applyPLayer(
  acc: PLayer,
  layer: PLayer,
  src: PropSource,
  source: Record<string, PropSource>,
): void {
  for (const k of PLAIN_PKEYS) {
    const v = layer[k]
    if (v === undefined) continue
    ;(acc as Record<string, unknown>)[k] = v
    source[k] = src
  }
  if (layer.indentChars) {
    acc.indentChars = mergeCharIndents(acc.indentChars, layer.indentChars)
    const c = layer.indentChars
    if (c.left !== undefined) source.indentLeftChars = src
    if (c.right !== undefined) source.indentRightChars = src
    if (c.firstLine !== undefined || c.hanging !== undefined) source.indentSpecialChars = src
  }
  if (layer.tabStops) {
    acc.tabStops = mergeTabStops(acc.tabStops ?? [], layer.tabStops)
    source.tabStops = src
  }
  if (layer.borders) {
    acc.borders = mergeBorders(acc.borders, layer.borders)
    source.borders = src
  }
}

// ---------------------------------------------------------------- paragraphs

function listRefOf(
  ctx: ResolveCtx,
  pPr: XNode | undefined,
  style: StyleInfo | undefined,
): { numId: string; ilvl: number } | undefined {
  const numPr = pPr ? findChild(pPr, 'w:numPr') : undefined
  const directNumId = numPr ? attrsOf(findChild(numPr, 'w:numId') ?? {})['w:val'] : undefined
  const directIlvl = numPr ? attrsOf(findChild(numPr, 'w:ilvl') ?? {})['w:val'] : undefined
  if (directNumId === '0') return undefined
  const styleNum = style?.numPr
  const styleNumPr = styleNum === 'none' ? undefined : styleNum
  const numId = directNumId ?? styleNumPr?.numId
  if (!numId) return undefined
  const ilvl = directIlvl !== undefined ? parseInt(directIlvl, 10) || 0 : (styleNumPr?.ilvl ?? 0)
  return { numId, ilvl }
}

function tableStyleOf(ctx: ResolveCtx): StyleInfo | undefined {
  return ctx.table?.styleId ? ctx.styles.get(ctx.table.styleId) : undefined
}

function cacheKey(ctx: ResolveCtx, opts: ResolveParaOptions | undefined): string {
  const t = ctx.table ? `${ctx.table.styleId ?? ''}:${ctx.table.cond.join(',')}` : ''
  const u = opts?.charUnits ? `${opts.charUnits.run}/${opts.charUnits.normal}` : ''
  return `${t}|${u}`
}

export function resolveParagraphProps(
  pPr: XNode | undefined,
  ctx: ResolveCtx,
  opts?: ResolveParaOptions,
): EffectivePPr {
  const key = cacheKey(ctx, opts)
  const cache = (ctx.cache ??= new WeakMap())
  if (pPr) {
    const hit = cache.get(pPr)
    if (hit && hit.key === key) return hit.eff
  }
  const eff = resolveParagraphPropsUncached(pPr, ctx, opts)
  if (pPr) cache.set(pPr, { key, eff })
  return eff
}

function resolveParagraphPropsUncached(
  pPr: XNode | undefined,
  ctx: ResolveCtx,
  opts: ResolveParaOptions | undefined,
): EffectivePPr {
  const styleId = pPr ? attrsOf(findChild(pPr, 'w:pStyle') ?? {})['w:val'] : undefined
  const style = (styleId ? ctx.styles.get(styleId) : undefined) ?? ctx.defaultParaStyle
  const listRef = listRefOf(ctx, pPr, style)
  const level = listRef ? ctx.numbering.get(listRef.numId)?.levels[listRef.ilvl] : undefined
  const source: Record<string, PropSource> = {}
  const acc: PLayer = {
    bidi: false,
    indentLeft: 0,
    indentRight: 0,
    indentFirstLine: 0,
    spaceBefore: 0,
    spaceAfter: 0,
    spaceBeforeAuto: false,
    spaceAfterAuto: false,
    keepNext: false,
    keepLines: false,
    pageBreakBefore: false,
    widowControl: true,
    contextualSpacing: false,
    suppressLineNumbers: false,
    snapToGrid: true,
    autoSpaceDE: true,
    autoSpaceDN: true,
    wordWrap: true,
    overflowPunct: true,
  }
  for (const k of Object.keys(acc)) source[k] = 'builtin'
  applyPLayer(acc, docDefaultsPLayer(ctx.docDefaults), 'docDefaults', source)
  applyPLayer(acc, tableStylePLayer(ctx), 'tableStyle', source)
  const styleLayer = stylePLayer(ctx, style)
  // numbering owns a list item's indent: the style chain's character indents stay off it
  applyPLayer(
    acc,
    listRef ? { ...styleLayer, indentChars: undefined } : styleLayer,
    'paraStyle',
    source,
  )
  // Word lays the level's w:ind over the paragraph style's (probe 2026-10-20), unlike the §17.7.2 order
  applyPLayer(acc, levelPLayer(level), 'numbering', source)
  applyPLayer(acc, directPLayer(pPr, ctx.themeColors), 'direct', source)

  // Word quirk: w:jc left/right are logical in bidi paragraphs; the model stores the visual side
  let align = acc.align
  if (acc.bidi && (align === 'left' || align === 'right')) {
    align = align === 'left' ? 'right' : 'left'
  }

  let indent = { left: acc.indentLeft!, right: acc.indentRight!, firstLine: acc.indentFirstLine! }
  const chars = activeCharIndents(acc.indentChars)
  if (chars && opts?.charUnits) {
    const f = resolveCharIndents(
      { indentLeft: indent.left, indentRight: indent.right, indentFirstLine: indent.firstLine },
      chars,
      opts.charUnits,
    )
    indent = {
      left: f.indentLeft ?? 0,
      right: f.indentRight ?? 0,
      firstLine: f.indentFirstLine ?? 0,
    }
    // each twips key takes the source of the character component that moved it
    const by: Array<[keyof typeof f, PropSource | undefined]> = [
      ['indentLeft', chars.left !== undefined ? source.indentLeftChars : source.indentSpecialChars],
      ['indentRight', source.indentRightChars],
      ['indentFirstLine', source.indentSpecialChars],
    ]
    for (const [k, src] of by) {
      if (src && (acc as Record<string, unknown>)[k] !== f[k]) source[k] = src
    }
  }

  const para: EffectivePPr = {
    ...(styleId ? { styleId } : {}),
    ...(align ? { align } : {}),
    bidi: acc.bidi!,
    indent: { ...indent, ...(chars ? { chars } : {}) },
    spacing: {
      before: acc.spaceBefore!,
      after: acc.spaceAfter!,
      beforeAuto: acc.spaceBeforeAuto!,
      afterAuto: acc.spaceAfterAuto!,
      ...(acc.line ? { line: acc.line } : {}),
    },
    keepNext: acc.keepNext!,
    keepLines: acc.keepLines!,
    pageBreakBefore: acc.pageBreakBefore!,
    widowControl: acc.widowControl!,
    contextualSpacing: acc.contextualSpacing!,
    suppressLineNumbers: acc.suppressLineNumbers!,
    snapToGrid: acc.snapToGrid!,
    autoSpaceDE: acc.autoSpaceDE!,
    autoSpaceDN: acc.autoSpaceDN!,
    wordWrap: acc.wordWrap!,
    overflowPunct: acc.overflowPunct!,
    tabStops: acc.tabStops ?? [],
    ...(acc.borders ? { borders: acc.borders } : {}),
    ...(acc.shading ? { shading: acc.shading } : {}),
    ...(acc.eastAsiaLang ? { eastAsiaLang: acc.eastAsiaLang } : {}),
    ...(listRef ? { list: { ...listRef, ...(level ? { level } : {}) } } : {}),
    mark: undefined as unknown as EffectiveRPr,
    hidden: false,
    source,
  }
  const markRPr = pPr ? findChild(pPr, 'w:rPr') : undefined
  para.mark = resolveRunProps(markRPr, para, ctx)
  // a w:specVanish mark is a style separator, never a hidden paragraph
  para.hidden = para.mark.vanish && !(markRPr && onOffOf(markRPr, 'w:specVanish') === true)
  return para
}

/** resolveParagraphProps for callers holding a raw <w:p> / <w:pPr> XML slice */
export function resolveParagraphPropsXml(
  xml: string,
  ctx: ResolveCtx,
  opts?: ResolveParaOptions,
): EffectivePPr {
  const m = /<w:pPr(?:\s[^>]*)?>[\s\S]*?<\/w:pPr>|<w:pPr\s*\/>/.exec(xml)
  if (!m) return resolveParagraphProps(undefined, ctx, opts)
  let pPr: XNode | undefined
  try {
    const parsed = xmlParser.parse(m[0]) as XNode[]
    pPr = parsed.find((n) => nameOf(n) === 'w:pPr')
  } catch {
    pPr = undefined
  }
  return resolveParagraphProps(pPr, ctx, opts)
}

// ---------------------------------------------------------------- runs

function styleRLayer(d: StyleDisplay | undefined, cs: boolean): RLayer {
  if (!d) return {}
  const out: RLayer = {}
  const f = styleRunFormat(d, cs)
  if (f.bold !== undefined) out.bold = f.bold
  if (f.italic !== undefined) out.italic = f.italic
  if (f.sizeHalfPoints !== undefined) out.sizeHalfPoints = f.sizeHalfPoints
  if (d.rtl !== undefined) out.rtl = d.rtl
  if (d.fontAscii ?? d.font) out.ascii = d.fontAscii ?? d.font
  if (d.eastAsiaFont) out.eastAsia = d.eastAsiaFont
  if (d.eaSlotEmpty !== undefined) out.eaSlotEmpty = d.eaSlotEmpty
  if (d.csFont) out.cs = d.csFont
  if (d.color) out.color = d.color
  if (d.underline !== undefined) out.underline = d.underline
  if (d.strike !== undefined) out.strike = d.strike
  if (d.caps) out.caps = d.caps
  if (d.vanish !== undefined) out.vanish = d.vanish
  if (d.charSpacingTwips !== undefined) out.charSpacingTwips = d.charSpacingTwips
  if (d.kernHalfPoints !== undefined) out.kernHalfPoints = d.kernHalfPoints
  if (d.bdr) out.bdr = d.bdr
  if (d.shading) out.shading = d.shading
  if (d.eastAsiaLang) out.eastAsiaLang = d.eastAsiaLang
  return out
}

function docDefaultsRLayer(d: DocDefaults | undefined, cs: boolean): RLayer {
  const out: RLayer = {}
  if (!d) return out
  const bold = cs ? d.boldCs : d.bold
  const italic = cs ? d.italicCs : d.italic
  const size = cs ? d.sizeCsHalfPoints : d.sizeHalfPoints
  if (bold !== undefined) out.bold = bold
  if (italic !== undefined) out.italic = italic
  if (size !== undefined) out.sizeHalfPoints = size
  if (d.rtl !== undefined) out.rtl = d.rtl
  if (d.asciiFont) out.ascii = d.asciiFont
  if (d.eastAsiaFont) out.eastAsia = d.eastAsiaFont
  if (d.eaSlotEmpty !== undefined) out.eaSlotEmpty = d.eaSlotEmpty
  if (d.csFont) out.cs = d.csFont
  if (d.color) out.color = d.color
  if (d.kernHalfPoints !== undefined) out.kernHalfPoints = d.kernHalfPoints
  if (d.eastAsiaLang) out.eastAsiaLang = d.eastAsiaLang
  return out
}

type CondLike = Partial<TableCondFormat> & { kernHalfPoints?: number }

function condRLayer(c: CondLike | undefined): RLayer {
  const out: RLayer = {}
  if (!c) return out
  if (c.bold !== undefined) out.bold = c.bold
  if (c.italic !== undefined) out.italic = c.italic
  if (c.sizeHalfPoints !== undefined) out.sizeHalfPoints = c.sizeHalfPoints
  if (c.color) out.color = c.color
  if (c.caps) out.caps = c.caps
  if (c.fontAscii) out.ascii = c.fontAscii
  if (c.charSpacingTwips !== undefined) out.charSpacingTwips = c.charSpacingTwips
  if (c.kernHalfPoints !== undefined) out.kernHalfPoints = c.kernHalfPoints
  return out
}

function tableStyleRLayer(ctx: ResolveCtx): RLayer {
  const t = tableStyleOf(ctx)?.tableDisplay
  if (!t) return {}
  const out = condRLayer(t.wholeTable)
  for (const cond of ctx.table?.cond ?? []) Object.assign(out, condRLayer(t[cond]))
  return out
}

function directRLayer(
  rPr: XNode | undefined,
  cs: boolean,
  themeColors: ThemeColors | null | undefined,
  themeFonts: ThemeFonts | null | undefined,
): RLayer {
  const out: RLayer = {}
  if (!rPr) return out
  const bold = onOffOf(rPr, cs ? 'w:bCs' : 'w:b')
  if (bold !== undefined) out.bold = bold
  const italic = onOffOf(rPr, cs ? 'w:iCs' : 'w:i')
  if (italic !== undefined) out.italic = italic
  const sz = parseInt(attrsOf(findChild(rPr, cs ? 'w:szCs' : 'w:sz') ?? {})['w:val'] ?? '', 10)
  if (Number.isFinite(sz)) out.sizeHalfPoints = Math.max(1, sz)
  const rtl = onOffOf(rPr, 'w:rtl')
  if (rtl !== undefined) out.rtl = rtl
  const rfNode = findChild(rPr, 'w:rFonts')
  if (rfNode) {
    const rf = themedRFonts(attrsOf(rfNode), themeFonts)
    if (rf.ascii ?? rf.hAnsi) out.ascii = rf.ascii ?? rf.hAnsi
    if (rf.eastAsia) out.eastAsia = rf.eastAsia
    if (rf.eaSlotEmpty) out.eaSlotEmpty = true
    if (rf.cs) out.cs = rf.cs
  }
  const color = colorFrom(rPr, themeColors) ?? autoColorOf(rPr)
  if (color) out.color = color
  // a w:u without w:val is not an underline (Word; python-docx notes agree)
  const u = attrsOf(findChild(rPr, 'w:u') ?? {})['w:val']
  if (u !== undefined) out.underline = u !== 'none'
  const strike = onOffOf(rPr, 'w:strike')
  if (strike !== undefined) out.strike = strike
  const caps = onOffOf(rPr, 'w:caps')
  const small = onOffOf(rPr, 'w:smallCaps')
  if (caps === true) out.caps = 'all'
  else if (small === true) out.caps = 'small'
  else if (caps === false || small === false) out.caps = 'none'
  if (onOffOf(rPr, 'w:specVanish') !== true) {
    const vanish = onOffOf(rPr, 'w:vanish')
    if (vanish !== undefined) out.vanish = vanish
  }
  const spacing = parseInt(attrsOf(findChild(rPr, 'w:spacing') ?? {})['w:val'] ?? '', 10)
  if (Number.isFinite(spacing)) out.charSpacingTwips = spacing
  const kern = attrsOf(findChild(rPr, 'w:kern') ?? {})['w:val']
  if (kern !== undefined) out.kernHalfPoints = parseInt(kern, 10) || 0
  const va = attrsOf(findChild(rPr, 'w:vertAlign') ?? {})['w:val']
  if (va === 'superscript' || va === 'subscript' || va === 'baseline') out.vertAlign = va
  const bdrNode = findChild(rPr, 'w:bdr')
  if (bdrNode) out.bdr = runBorderOf(bdrNode) ?? null
  const shd = findChild(rPr, 'w:shd')
  if (shd) {
    const display = shdDisplayFill(shd, themeColors)
    const fill = attrsOf(shd)['w:fill']
    out.shading = display ?? (fill && fill !== 'auto' ? stripHash(fill) : '')
  }
  const lang = attrsOf(findChild(rPr, 'w:lang') ?? {})['w:eastAsia']
  if (lang) out.eastAsiaLang = lang
  return out
}

const PLAIN_RKEYS = [
  'sizeHalfPoints',
  'ascii',
  'eastAsia',
  'cs',
  'eaSlotEmpty',
  'color',
  'underline',
  'caps',
  'charSpacingTwips',
  'kernHalfPoints',
  'vertAlign',
  'bdr',
  'shading',
  'eastAsiaLang',
] as const

export function resolveRunProps(
  rPr: XNode | undefined,
  para: EffectivePPr,
  ctx: ResolveCtx,
): EffectiveRPr {
  const rStyleId = rPr ? attrsOf(findChild(rPr, 'w:rStyle') ?? {})['w:val'] : undefined
  const charStyle = rStyleId ? ctx.styles.get(rStyleId) : undefined
  const paraStyle = para.styleId ? ctx.styles.get(para.styleId) : ctx.defaultParaStyle
  const cs =
    ((rPr ? onOffOf(rPr, 'w:rtl') : undefined) ??
      charStyle?.display?.rtl ??
      paraStyle?.display?.rtl ??
      ctx.docDefaults?.rtl) === true

  const source: Record<string, PropSource> = {}
  const acc: RLayer = {
    bold: false,
    italic: false,
    sizeHalfPoints: WORD_DEFAULT_SIZE_HALF_POINTS,
    underline: false,
    strike: false,
    caps: 'none',
    vanish: false,
    charSpacingTwips: 0,
  }
  for (const k of Object.keys(acc)) source[k] = 'builtin'
  const styleLayers: Array<[RLayer, PropSource]> = [
    [tableStyleRLayer(ctx), 'tableStyle'],
    [styleRLayer(paraStyle?.display, cs), 'paraStyle'],
    [styleRLayer(charStyle?.display, cs), 'charStyle'],
  ]
  const apply = (layer: RLayer, src: PropSource) => {
    for (const k of PLAIN_RKEYS) {
      const v = layer[k]
      if (v === undefined) continue
      ;(acc as Record<string, unknown>)[k] = v
      source[k] = src
    }
  }
  const defaults = docDefaultsRLayer(ctx.docDefaults, cs)
  apply(defaults, 'docDefaults')
  for (const t of TOGGLES) {
    if (defaults[t] !== undefined) {
      acc[t] = defaults[t]
      source[t] = 'docDefaults'
    }
  }
  for (const [layer, src] of styleLayers) {
    apply(layer, src)
    // toggle properties XOR across the style layers (ECMA-376 §17.7.3)
    for (const t of TOGGLES) {
      const v = layer[t]
      if (v === undefined) continue
      const base = source[t] === 'builtin' || source[t] === 'docDefaults' ? false : acc[t]!
      acc[t] = base !== v
      source[t] = src
    }
  }
  const direct = directRLayer(rPr, cs, ctx.themeColors, ctx.themeFonts)
  apply(direct, 'direct')
  for (const t of TOGGLES) {
    if (direct[t] !== undefined) {
      acc[t] = direct[t]
      source[t] = 'direct'
    }
  }

  const fonts: EffectiveRPr['fonts'] = {}
  if (acc.ascii) fonts.ascii = acc.ascii
  if (acc.eastAsia) fonts.eastAsia = acc.eastAsia
  if (acc.cs) fonts.cs = acc.cs
  if (acc.eaSlotEmpty) fonts.eaSlotEmpty = true
  return {
    ...(rStyleId ? { styleId: rStyleId } : {}),
    cs,
    bold: acc.bold!,
    italic: acc.italic!,
    sizeHalfPoints: acc.sizeHalfPoints!,
    fonts,
    ...(acc.color ? { color: acc.color } : {}),
    underline: acc.underline!,
    strike: acc.strike!,
    caps: acc.caps!,
    vanish: acc.vanish!,
    charSpacingTwips: acc.charSpacingTwips!,
    ...(acc.kernHalfPoints !== undefined ? { kernHalfPoints: acc.kernHalfPoints } : {}),
    ...(acc.vertAlign && acc.vertAlign !== 'baseline' ? { vertAlign: acc.vertAlign } : {}),
    ...(acc.bdr ? { bdr: acc.bdr } : {}),
    ...(acc.shading ? { shading: acc.shading } : {}),
    ...(acc.eastAsiaLang ? { eastAsiaLang: acc.eastAsiaLang } : {}),
    source,
  }
}

// ---------------------------------------------------------------- direct view

/**
 * The paragraph's own w:pPr as the direct-only ParaFormat the model saves:
 * every key the resolver sourced from the direct layer, in extractParaFormat's
 * shape. w:bidi inherited from the chain surfaces as bidiInherited (the
 * generator needs it to write the logical jc Word expects).
 */
export function directParaFormat(
  eff: EffectivePPr,
): Pick<
  ParaFormat,
  | 'align'
  | 'bidi'
  | 'bidiInherited'
  | 'lineSpacing'
  | 'lineRule'
  | 'lineRawTwips'
  | 'indentLeft'
  | 'indentRight'
  | 'indentFirstLine'
  | 'spaceBefore'
  | 'spaceAfter'
  | 'spaceBeforeAuto'
  | 'spaceAfterAuto'
  | 'pageBreakBefore'
  | 'keepNext'
  | 'keepLines'
  | 'suppressLineNumbers'
  | 'widowControl'
  | 'snapToGrid'
  | 'autoSpace'
  | 'autoSpaceDE'
  | 'autoSpaceDN'
  | 'wordWrap'
  | 'overflowPunct'
  | 'contextualSpacing'
> {
  const direct = (k: string) => eff.source[k] === 'direct'
  const f: ReturnType<typeof directParaFormat> = {}
  if (eff.bidi) {
    if (direct('bidi')) f.bidi = true
    else f.bidiInherited = true
  }
  if (direct('align') && eff.align) f.align = eff.align
  if (direct('line')) Object.assign(f, effectiveLineFormat(eff.spacing.line))
  if (direct('spaceBeforeAuto')) f.spaceBeforeAuto = eff.spacing.beforeAuto
  if (direct('spaceAfterAuto')) f.spaceAfterAuto = eff.spacing.afterAuto
  if (direct('spaceBefore')) f.spaceBefore = eff.spacing.before
  if (direct('spaceAfter')) f.spaceAfter = eff.spacing.after
  if (direct('indentLeft')) f.indentLeft = eff.indent.left
  if (direct('indentRight')) f.indentRight = eff.indent.right
  if (direct('indentFirstLine')) f.indentFirstLine = eff.indent.firstLine
  for (const k of [
    'pageBreakBefore',
    'keepNext',
    'keepLines',
    'suppressLineNumbers',
    'widowControl',
    'contextualSpacing',
    'wordWrap',
    'overflowPunct',
  ] as const) {
    if (direct(k)) f[k] = eff[k]
  }
  if (direct('snapToGrid') && !eff.snapToGrid) f.snapToGrid = false
  const de = direct('autoSpaceDE') ? eff.autoSpaceDE : undefined
  const dn = direct('autoSpaceDN') ? eff.autoSpaceDN : undefined
  if (de !== undefined) f.autoSpaceDE = de
  if (dn !== undefined) f.autoSpaceDN = dn
  const auto = autoSpaceMode(de, dn)
  if (auto !== undefined) f.autoSpace = auto
  return f
}

/**
 * The run's own w:rPr as the direct-only Run fields the model saves: every key
 * the resolver sourced from the direct layer, in extractRun's shape (font is
 * the eastAsia-first display face; rtl, colour and shading keep their own
 * readers).
 */
export function directRunFormat(
  eff: EffectiveRPr,
): Pick<
  Run,
  | 'bold'
  | 'italic'
  | 'sizeHalfPoints'
  | 'underline'
  | 'strike'
  | 'caps'
  | 'charSpacingTwips'
  | 'kernHalfPoints'
  | 'vertAlign'
  | 'font'
  | 'fontAscii'
  | 'eastAsiaFont'
  | 'eaSlotEmpty'
  | 'csFont'
> {
  const direct = (k: string) => eff.source[k] === 'direct'
  const f: ReturnType<typeof directRunFormat> = {}
  if (direct('bold')) f.bold = eff.bold
  if (direct('italic')) f.italic = eff.italic
  if (direct('underline')) f.underline = eff.underline
  if (direct('strike')) f.strike = eff.strike
  if (direct('sizeHalfPoints')) f.sizeHalfPoints = eff.sizeHalfPoints
  const eastAsia = direct('eastAsia') ? eff.fonts.eastAsia : undefined
  const ascii = direct('ascii') ? eff.fonts.ascii : undefined
  const font = eastAsia ?? ascii
  if (font) f.font = font
  if (eastAsia) f.eastAsiaFont = eastAsia
  if (eastAsia && direct('eaSlotEmpty') && eff.fonts.eaSlotEmpty) f.eaSlotEmpty = true
  if (ascii) f.fontAscii = ascii
  if (direct('cs') && eff.fonts.cs) f.csFont = eff.fonts.cs
  if (direct('charSpacingTwips')) f.charSpacingTwips = eff.charSpacingTwips
  if (direct('kernHalfPoints')) f.kernHalfPoints = eff.kernHalfPoints
  if (direct('caps')) f.caps = eff.caps
  if (direct('vertAlign') && eff.vertAlign) f.vertAlign = eff.vertAlign
  return f
}

// ---------------------------------------------------------------- shadow

export interface ResolveDiff {
  /** which parse surface produced the paragraph */
  surface?: 'body' | 'cell' | 'run'
  key: string
  old: unknown
  eff: unknown
  text?: string
}

/** key-by-key comparison of a legacy direct reader's output with the resolver's */
export function shadowCompare(
  legacy: object | undefined,
  resolved: object | undefined,
  keys?: readonly string[],
): ResolveDiff[] {
  const out: ResolveDiff[] = []
  const a = legacy ?? {}
  const b = resolved ?? {}
  for (const key of keys ?? new Set([...Object.keys(a), ...Object.keys(b)])) {
    if (key === 'effective') continue
    const o = (a as Record<string, unknown>)[key]
    const e = (b as Record<string, unknown>)[key]
    if (JSON.stringify(o) !== JSON.stringify(e)) out.push({ key, old: o, eff: e })
  }
  return out
}

/** ResolveCtx over a parsed document (tests, tools, renderer-side consumers) */
export function resolveCtxOf(doc: {
  styles: Map<string, StyleInfo>
  docDefaults?: DocDefaults
  numbering: Map<string, NumberingDef>
  themeFonts?: ThemeFonts | null
  themeColors?: ThemeColors | null
  compatibilityMode?: number
  doNotUseHTMLParagraphAutoSpacing?: boolean
  indentNotNumberingTabStop?: boolean
}): ResolveCtx {
  const defaultParaStyle = [...doc.styles.values()].find(
    (s) => s.type === 'paragraph' && s.isDefault,
  )
  return {
    styles: doc.styles,
    ...(defaultParaStyle ? { defaultParaStyle } : {}),
    ...(doc.docDefaults ? { docDefaults: doc.docDefaults } : {}),
    numbering: doc.numbering,
    themeFonts: doc.themeFonts,
    themeColors: doc.themeColors,
    compat: {
      mode: doc.compatibilityMode ?? 0,
      additiveParaSpacing: doc.doNotUseHTMLParagraphAutoSpacing === true,
      indentNotNumberingTabStop: doc.indentNotNumberingTabStop === true,
    },
  }
}

/** Word's HTML auto paragraph spacing (w:beforeAutospacing / w:afterAutospacing), twips */
export const AUTO_SPACING_TWIPS = 280

/** w:*Autospacing resolved for a document: Word's 14pt, 0 under doNotUseHTMLParagraphAutoSpacing */
export function autoSpacingTwips(compat: ResolveCompat): number {
  return compat.additiveParaSpacing ? 0 : AUTO_SPACING_TWIPS
}

export interface EffectiveSpacingTwips {
  before: number
  after: number
}

/** before / after spacing with the auto flags folded in (twips) */
export function effectiveSpacingTwips(
  eff: EffectivePPr,
  compat: ResolveCompat,
): EffectiveSpacingTwips {
  const auto = autoSpacingTwips(compat)
  return {
    before: eff.spacing.beforeAuto ? auto : eff.spacing.before,
    after: eff.spacing.afterAuto ? auto : eff.spacing.after,
  }
}

/** line spacing in the ParaFormat triple (lineRule / lineRawTwips / lineSpacing multiple) */
export function effectiveLineFormat(
  line: LineSpacing | undefined,
): Pick<ParaFormat, 'lineRule' | 'lineRawTwips' | 'lineSpacing'> {
  if (!line) return {}
  return {
    lineRule: line.rule,
    lineRawTwips: line.raw,
    ...(line.rule === 'auto' ? { lineSpacing: Math.round((line.raw / 240) * 100) / 100 } : {}),
  }
}

function effectiveLineSource(eff: EffectivePPr, sources: readonly PropSource[] | undefined) {
  return !sources || sources.includes(eff.source.line)
}

/** the layers whose line spacing a header/footer surface carries inline: the
 *  strip base already lays out Normal / docDefaults, a named style and the
 *  direct pPr (plus the table style for cells) come on the paragraph */
export function hfLineSources(eff: EffectivePPr, inTable: boolean): PropSource[] {
  const out: PropSource[] = ['direct']
  if (inTable) out.push('tableStyle')
  if (eff.styleId) out.push('paraStyle')
  return out
}

/**
 * The resolved paragraph as display-only ParaFormat keys (non-editable
 * surfaces: header/footer strips, cells, notes). Zero geometry is omitted so
 * the consumers' "absent = none" checks keep working.
 */
export interface EffectiveParaFormatOptions {
  /** layers whose line spacing travels inline; absent = every layer. Surfaces
   *  whose base line already carries Normal / docDefaults pass the style and
   *  direct layers only */
  lineSources?: readonly PropSource[]
}

export function effectiveParaFormat(
  eff: EffectivePPr,
  compat: ResolveCompat,
  opts: EffectiveParaFormatOptions = {},
): Pick<
  ParaFormat,
  | 'align'
  | 'bidi'
  | 'indentLeft'
  | 'indentRight'
  | 'indentFirstLine'
  | 'spaceBefore'
  | 'spaceAfter'
  | 'lineRule'
  | 'lineRawTwips'
  | 'lineSpacing'
  | 'borders'
  | 'borderLines'
  | 'borderPad'
  | 'borderReset'
> {
  const sp = effectiveSpacingTwips(eff, compat)
  const borders = eff.borders ? paraBordersOf(eff.borders) : undefined
  const line = effectiveLineSource(eff, opts.lineSources) ? eff.spacing.line : undefined
  return {
    ...(eff.align ? { align: eff.align } : {}),
    ...(eff.bidi ? { bidi: true } : {}),
    ...(eff.indent.left ? { indentLeft: eff.indent.left } : {}),
    ...(eff.indent.right ? { indentRight: eff.indent.right } : {}),
    ...(eff.indent.firstLine ? { indentFirstLine: eff.indent.firstLine } : {}),
    ...(sp.before ? { spaceBefore: sp.before } : {}),
    ...(sp.after ? { spaceAfter: sp.after } : {}),
    ...effectiveLineFormat(line),
    ...(borders?.borders ? { borders: borders.borders } : {}),
    ...(borders?.borderLines ? { borderLines: borders.borderLines } : {}),
    ...(borders?.borderPad ? { borderPad: borders.borderPad } : {}),
    ...(borders?.borderReset ? { borderReset: borders.borderReset } : {}),
  }
}
