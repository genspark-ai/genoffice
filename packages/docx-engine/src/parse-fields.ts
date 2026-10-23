// Field-code display (PAGE, TOC, REF, ...) and TOC entry numbering.
import { inlineEqFieldResults } from './eq-field'
import { computeListMarkers, type ListItemRef } from './list-markers'
import { decodeEntities, plainText } from './parse-xml-text'
import type { Block, FieldDisplay, NumberingDef, StyleInfo, TabStop } from './types'
import {
  resolveCtxOf,
  resolveParagraphPropsXml,
  resolveRunProps,
  type EffectivePPr,
  type EffectiveRPr,
  type ResolveCtx,
} from './resolve-props'
import { nameOf, xmlParser, type XNode } from './xml-utils'

/** resolver over nothing: callers without a document (tests, tools) */
let bareCtx: ResolveCtx | undefined
function resolverOrBare(resolver: ResolveCtx | undefined): ResolveCtx {
  return resolver ?? (bareCtx ??= resolveCtxOf({ styles: new Map(), numbering: new Map() }))
}

/** parse a <w:rPr> XML slice to a node (undefined when absent or malformed) */
export function rPrNodeOf(xml: string): XNode | undefined {
  const m = /<w:rPr(?:\s[^>]*)?>[\s\S]*?<\/w:rPr>/.exec(xml)
  if (!m) return undefined
  try {
    return (xmlParser.parse(m[0]) as XNode[]).find((n) => nameOf(n) === 'w:rPr')
  } catch {
    return undefined
  }
}

/**
 * Display-only rendering hint for protected field paragraphs. The visible
 * field *result* (w:t runs; instruction text lives in w:instrText and is
 * excluded) is shown instead of a generic chip. Original XML still saves
 * byte-identical.
 */
/** TOC entry level: styleId "TOC1" (Word) / "TOC 1" (Pages), or the style's
 *  name "toc 1" when the styleId is opaque (html2docx exports numeric ids) */
export function tocLevelOf(styleId: string, styles?: Map<string, StyleInfo>): number | null {
  const m =
    /^TOC ?([1-9])$/i.exec(styleId) ?? /^toc ?([1-9])$/i.exec(styles?.get(styleId)?.name ?? '')
  if (m) return parseInt(m[1], 10)
  // table-of-figures entries (TOC \c field results) are level-1 toc lines
  if (
    /^TableofFigures$/i.test(styleId) ||
    /^table of figures$/i.test(styles?.get(styleId)?.name ?? '')
  )
    return 1
  return null
}

const TAB_LEADERS = ['none', 'dot', 'hyphen', 'underscore', 'heavy', 'middleDot'] as const

/** leader of the entry's page-number tab: the last right stop of the direct
 *  w:tabs (a stop without w:leader is a bare tab), else of the style's stops */
function tocLeaderOf(pPr: string, style?: StyleInfo): TabStop['leader'] | undefined {
  const tabsXml = /<w:tabs>[\s\S]*?<\/w:tabs>/.exec(pPr)?.[0]
  if (tabsXml) {
    const rights = Array.from(tabsXml.matchAll(/<w:tab\s[^>]*\/>/g), (m) => m[0]).filter((t) =>
      /\sw:val=(?:"right"|'right')/.test(t),
    )
    const last = rights[rights.length - 1]
    if (last) {
      const v = /\sw:leader=(?:"([^"]+)"|'([^']+)')/.exec(last)?.slice(1, 3).find(Boolean) ?? 'none'
      return (TAB_LEADERS as readonly string[]).includes(v) ? (v as TabStop['leader']) : 'none'
    }
  }
  const stop = style?.display?.tabStops?.filter((t) => t.val === 'right').pop()
  return stop ? (stop.leader ?? 'none') : undefined
}

const CJK_INK_RE = /[\u2E80-\u9FFF\uAC00-\uD7AF\uF900-\uFAFF\uFF00-\uFFEF]/

/** a face the run's own w:rFonts declares for the slot */
function directFont(r: EffectiveRPr, slot: 'ascii' | 'eastAsia'): string | undefined {
  return r.source[slot] === 'direct' ? r.fonts[slot] : undefined
}

/** direct face of a run for its script: eastAsia for CJK text, else ascii (hAnsi fallback) */
function leadingRunFont(r: EffectiveRPr, text: string): string | undefined {
  if (r.source.ascii !== 'direct' && r.source.eastAsia !== 'direct') return undefined
  const ea = CJK_INK_RE.test(text) ? directFont(r, 'eastAsia') : undefined
  return ea ?? directFont(r, 'ascii')
}

/** eastAsia face of a run that draws CJK ink */
function cjkRunFont(r: EffectiveRPr, text: string): string | undefined {
  return CJK_INK_RE.test(text) ? directFont(r, 'eastAsia') : undefined
}

/** a w:del run wrapper with its content (not the self-closing paragraph-mark w:del in pPr/rPr) */
const DEL_WRAPPER_RE = /<w:del(?:\s[^>]*)?(?<!\/)>[\s\S]*?<\/w:del>/g

/**
 * Attribute value of the first w: element that declares it. XML allows either
 * quote delimiter, so a double-quote-only pattern read a producer that wrote '
 * as if the attribute were absent.
 */
export function tagAttr(xml: string, tag: string, name: string): string | undefined {
  return new RegExp(`<${tag}\\b[^>]*\\s${name}=(["'])([^"']*)\\1`).exec(xml)?.[2]
}

/** direct paragraph geometry Word applies over the style: before/after spacing and w:ind
 *  (left/right, first-line; hanging is a negative first line, as Word merges them). Only
 *  direct-sourced values travel inline: the style chain reaches the wrapper through CSS */
export function directParaGeometry(
  eff: EffectivePPr,
): Pick<
  FieldDisplay,
  | 'spaceBeforeTwips'
  | 'spaceAfterTwips'
  | 'spaceBeforeAuto'
  | 'spaceAfterAuto'
  | 'indentLeftTwips'
  | 'indentRightTwips'
  | 'indentFirstLineTwips'
> {
  const direct = (k: string) => eff.source[k] === 'direct'
  return {
    ...(direct('spaceBefore') ? { spaceBeforeTwips: eff.spacing.before } : {}),
    ...(direct('spaceAfter') ? { spaceAfterTwips: eff.spacing.after } : {}),
    ...(direct('spaceBeforeAuto') && eff.spacing.beforeAuto ? { spaceBeforeAuto: true } : {}),
    ...(direct('spaceAfterAuto') && eff.spacing.afterAuto ? { spaceAfterAuto: true } : {}),
    ...(direct('indentLeft') ? { indentLeftTwips: eff.indent.left } : {}),
    ...(direct('indentRight') ? { indentRightTwips: eff.indent.right } : {}),
    ...(direct('indentFirstLine') ? { indentFirstLineTwips: eff.indent.firstLine } : {}),
  }
}

/** direct w:spacing line of the paragraph (OOXML defaults w:lineRule to auto) */
export function directLineOf(
  eff: EffectivePPr,
): Pick<FieldDisplay, 'lineRule' | 'lineRawTwips' | 'lineSpacing'> {
  const line = eff.spacing.line
  if (eff.source.line !== 'direct' || !line || line.raw <= 0) return {}
  return {
    lineRule: line.rule,
    lineRawTwips: line.raw,
    ...(line.rule === 'auto' ? { lineSpacing: Math.round((line.raw / 240) * 100) / 100 } : {}),
  }
}

/** compat flag the renderer's spacing helper reads instead of its module global */
export function additiveSpacingOf(ctx: ResolveCtx): Pick<FieldDisplay, 'additiveSpacing'> {
  return ctx.compat.additiveParaSpacing ? { additiveSpacing: true } : {}
}

/** a TOC entry Word wrote without TOC styles: PAGEREF page number behind a right tab stop */
function unstyledTocEntry(xml: string, pPr: string): boolean {
  return (
    /<w:instrText[^>]*>\s*PAGEREF\s/.test(xml) &&
    /<w:tabs>[\s\S]*?<w:tab\s[^>]*w:val=(?:"right"|'right')/.test(pPr)
  )
}

/** a visible result run resolved through the chain; its direct rPr decides what travels inline */
interface ResultRun {
  text: string
  run: EffectiveRPr
}

function directSize(r: EffectiveRPr): number {
  return r.source.sizeHalfPoints === 'direct' ? r.sizeHalfPoints : 0
}

/** visible result runs (w:t / w:delText; field machinery excluded) with their resolved rPr */
function resultRuns(
  xml: string,
  eff: EffectivePPr,
  ctx: ResolveCtx,
  opts: { deleted: boolean },
): ResultRun[] {
  const out: ResultRun[] = []
  const textRe = opts.deleted ? /<w:(?:t|delText)(?:\s|>)/ : /<w:t(?:\s|>)/
  const runRe = /<w:r(?:\s[^>]*)?>([\s\S]*?)<\/w:r>/g
  let run: RegExpExecArray | null
  while ((run = runRe.exec(xml)) !== null) {
    if (!textRe.test(run[1]) || run[1].includes('<w:instrText')) continue
    const text = decodeEntities(
      Array.from(
        run[1].matchAll(/<w:(?:t|delText)(?:\s[^>]*)?>([\s\S]*?)<\/w:(?:t|delText)>/g),
        (m) => m[1],
      ).join(''),
    )
    out.push({ text, run: resolveRunProps(rPrNodeOf(run[1]), eff, ctx) })
  }
  return out
}

export function fieldDisplayOf(
  xml: string,
  styles?: Map<string, StyleInfo>,
  resolver?: ResolveCtx,
): FieldDisplay | undefined {
  const ctx = resolverOrBare(resolver)
  const styleId = tagAttr(xml, 'w:pStyle', 'w:val') ?? ''
  const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(xml)?.[0] ?? ''
  const eff = resolveParagraphPropsXml(xml, ctx)
  const geometry = directParaGeometry(eff)
  const line = directLineOf(eff)
  const additive = additiveSpacingOf(ctx)
  const styledLevel = tocLevelOf(styleId, styles)
  // unstyled entries indent through their own w:ind, styled ones through the level ladder
  const tocLevel = styledLevel ?? (unstyledTocEntry(xml, pPr) ? 1 : null)
  if (styledLevel !== null) {
    delete geometry.indentLeftTwips
    delete geometry.indentRightTwips
    delete geometry.indentFirstLineTwips
  }
  if (tocLevel !== null) {
    // TOC entry: title <tab with dot leader> page number. The page number
    // follows the LAST tab — entries like "1.1.<tab>Title<tab>7" put a leading
    // outline number at the first tab stop, not the page number.
    // tracked deletions leave the result: an entry whose every run is deleted
    // shows its deleted text struck through (and vanishes in balloon view)
    const live = xml.replace(DEL_WRAPPER_RE, '')
    const deleted = !/<w:t(?:\s|>)/.test(live) && /<w:delText(?:\s|>)/.test(xml)
    const segs: string[] = ['']
    // run-level tab is attribute-less CT_Empty, but LO/Google converters emit
    // it spaced (<w:tab />) or paired (<w:tab></w:tab>): match those too, while
    // still excluding tab-stop definitions (<w:tab w:val=…/> in w:tabs).
    const re =
      /<w:(?:t|delText)(?:\s[^>]*)?>([\s\S]*?)<[/]w:(?:t|delText)>|<w:tab\s*[/]>|<w:tab>\s*<[/]w:tab>/g
    let m: RegExpExecArray | null
    while ((m = re.exec(deleted ? xml : live)) !== null) {
      if (m[1] === undefined) segs.push('')
      else segs[segs.length - 1] += m[1]
    }
    const right = segs.length > 1 ? segs.pop()! : ''
    // a short space-free first segment at its own tab stop is the outline
    // number; it renders in the num cell so the title stays clean for
    // heading matching (toc-refresh keys on `left`)
    let num: string | undefined
    if (segs.length > 1) {
      const first = decodeEntities(segs[0]).trim()
      if (/^\S{1,15}$/.test(first)) {
        num = first
        segs.shift()
      }
    }
    const left = segs
      .map((s) => decodeEntities(s).trim())
      .filter(Boolean)
      .join(' ')
    const anchor = /<w:hyperlink [^>]*w:anchor="([^"]+)"/.exec(xml)?.[1]
    // direct pPr/run metrics: Word sizes TOC lines by them while the style
    // (html2docx exports) often carries nothing
    const leader = tocLeaderOf(pPr, styles?.get(styleId))
    // font size from visible result runs only: field-machinery runs
    // (fldChar/instrText) often carry the target heading's size and would
    // inflate the whole line
    const runs = resultRuns(xml, eff, ctx, { deleted: true })
    const sz = Math.max(0, ...runs.map((r) => directSize(r.run)))
    // face/weight of the leading visible run: Word draws the entry with the
    // result runs' rPr (a Times bold TOC under a Calibri body), the paragraph
    // style alone often says nothing
    const lead = runs[0]
    const font = lead ? (leadingRunFont(lead.run, lead.text) ?? '') : undefined
    const bold = lead?.run.source.bold === 'direct' && lead.run.bold
    const runStyleId = lead?.run.styleId
    let eaFont = runs.map((r) => cjkRunFont(r.run, r.text)).find((f) => f !== undefined)
    if (eaFont === font) eaFont = undefined
    return {
      kind: 'tocLine',
      left,
      right: decodeEntities(right).trim(),
      level: tocLevel,
      ...(num ? { num } : {}),
      ...(anchor ? { anchor } : {}),
      ...(deleted ? { deleted } : {}),
      ...(deleted && /<w:del\b[^>]*\/>/.test(pPr) ? { markDeleted: true } : {}),
      ...(sz > 0 ? { szHalfPoints: sz } : {}),
      ...(font ? { fontFamily: font } : {}),
      ...(font && eaFont ? { eaFontFamily: eaFont } : {}),
      ...(bold ? { bold } : {}),
      ...(runStyleId ? { runStyleId } : {}),
      ...(leader ? { leader } : {}),
      ...line,
      ...geometry,
      ...additive,
    }
  }
  xml = inlineEqFieldResults(xml)
  const visible = plainText(xml).trim()
  if (visible === '' && /<w:br\s[^>]*w:type="page"/.test(xml)) {
    return { kind: 'pageBreak' }
  }
  if (visible === '' && loneFieldEnd(xml)) {
    // the paragraph mark's size sets the empty line (the end run's rPr does not)
    const sz = directSize(eff.mark)
    return {
      kind: 'text',
      left: '',
      emptyLine: true,
      ...(sz > 0 ? { szHalfPoints: sz } : {}),
      ...line,
      ...geometry,
      ...additive,
    }
  }
  if (visible !== '') {
    // face/size of the visible result runs (same rule as tocLine): without
    // them the passthrough div inherits the document default and a SimSun
    // field paragraph mis-snaps to a double cell on a typed line grid
    // dominant size = the size covering the most text: a manual drop-cap letter
    // (one 48pt "L" before 10pt body) must not inflate the whole field's strut
    // (real_run2/47 rendered the entire paragraph at 48pt, 10 -> 17 pages)
    let font: string | undefined
    const szWeights = new Map<number, number>()
    for (const { text, run } of resultRuns(xml, eff, ctx, { deleted: false })) {
      // unsized runs vote for the inherited default (key 0): one explicit
      // drop-cap letter must not out-vote a body of default-sized text
      const key = directSize(run)
      szWeights.set(key, (szWeights.get(key) ?? 0) + Math.max(text.length, 1))
      if (!font) font = directFont(run, 'eastAsia') ?? directFont(run, 'ascii')
    }
    let sz = 0
    let szWeight = -1
    for (const [v, w] of szWeights) {
      if (w > szWeight || (w === szWeight && v > sz)) {
        sz = v
        szWeight = w
      }
    }
    // explicit paragraph alignment: the passthrough div would inherit the
    // document default (justify in CJK docs) and stretch short lines
    const direct = eff.source.align === 'direct' ? eff.align : undefined
    const align = direct === 'distribute' ? 'justify' : direct
    return {
      kind: 'text',
      left: visible,
      ...(sz > 0 ? { szHalfPoints: sz } : {}),
      ...(font ? { fontFamily: font } : {}),
      ...(align ? { align } : {}),
      ...line,
      ...geometry,
      ...additive,
    }
  }
  return undefined
}

/**
 * TOC entries carry their outline number ("1.", "1.1.") as w:numPr numbering
 * (Pages exports one numId per entry with startOverride restarts). The field
 * result is a display-only cache, so the marker is computed once at parse time
 * and stored on the tocLine FieldDisplay. Counters run document-wide in block
 * order, shared with editable list items (same abstractNum semantics).
 */
export function applyTocEntryNumbers(blocks: Block[], numbering: Map<string, NumberingDef>): void {
  if (numbering.size === 0) return
  const items: ListItemRef[] = []
  const tocAt = new Map<number, FieldDisplay>()
  for (const block of blocks) {
    const list = block.list?.numId ? block.list : block.strayList
    if (list?.numId) {
      items.push({ numId: list.numId, ilvl: list.ilvl })
      continue
    }
    const fd = block.fieldDisplay
    if (block.type !== 'passthrough' || fd?.kind !== 'tocLine' || !block.originalXml) continue
    const numPr = /<w:numPr>[\s\S]*?<\/w:numPr>/.exec(block.originalXml)?.[0]
    if (!numPr) continue
    const numId = /<w:numId w:val="([^"]+)"/.exec(numPr)?.[1]
    if (!numId) continue
    const ilvl = parseInt(/<w:ilvl w:val="(\d+)"/.exec(numPr)?.[1] ?? '0', 10)
    tocAt.set(items.length, fd)
    items.push({ numId, ilvl })
  }
  if (tocAt.size === 0) return
  const markers = computeListMarkers(items, numbering)
  for (const [i, fd] of tocAt) {
    const marker = markers[i]
    // bullets make no sense in front of a TOC entry; only ordered markers show
    if (marker && !/^[•◦▪➢❖✓]$/.test(marker)) fd.num = marker
  }
}

/** Open fields after a paragraph's fldChars; an entry turns true past its separator. */
export function fieldStackAfter(xml: string, stack: readonly boolean[]): boolean[] {
  const next = [...stack]
  const re = /<w:fldChar\b[^>]*\bw:fldCharType=(?:"(begin|separate|end)"|'(begin|separate|end)')/g
  let m: RegExpExecArray | null
  while ((m = re.exec(xml)) !== null) {
    const kind = m[1] ?? m[2]
    if (kind === 'begin') next.push(false)
    else if (kind === 'separate') {
      if (next.length > 0) next[next.length - 1] = true
    } else next.pop()
  }
  return next
}

/** Word hides a paragraph mark inside field code (begin..separate): the paragraphs join. */
export function markInsideFieldCode(stack: readonly boolean[]): boolean {
  return stack.length > 0 && !stack[stack.length - 1]
}

const FIELD_LABELS: Record<string, string> = {
  TOC: 'Auto TOC (updates when opened in Word)',
  PAGE: 'Page number field',
  NUMPAGES: 'Page count field',
  PAGEREF: 'Page reference field',
  REF: 'Cross-reference field',
  SEQ: 'Caption number field',
  HYPERLINK: 'Hyperlink field',
  DATE: 'Date field',
  TIME: 'Time field',
  INCLUDEPICTURE: 'Linked picture field',
  STYLEREF: 'Style reference field',
}

/** A paragraph whose only field machinery closes a field begun earlier (TOC end
 *  split off into its own paragraph); no run of its own, so Word lays an empty line. */
export function loneFieldEnd(xml: string): boolean {
  const ends = xml.match(/fldCharType=["']end["']/g)?.length ?? 0
  return (
    ends === 1 &&
    !/fldCharType=["'](?:begin|separate)["']/.test(xml) &&
    !xml.includes('<w:instrText') &&
    !xml.includes('<w:fldSimple') &&
    !xml.includes('<w:drawing') &&
    !xml.includes('<w:pict') &&
    !xml.includes('<w:object') &&
    !/<w:(?:tab|br|cr)\b/.test(xml.replace(/<w:tabs>[\s\S]*?<\/w:tabs>/g, ''))
  )
}

/** Human-readable label for a protected field paragraph, based on its field code. */
export function fieldLabel(xml: string): string {
  const instr =
    /<w:instrText[^>]*>([\s\S]*?)<\/w:instrText>/.exec(xml)?.[1] ??
    /<w:fldSimple[^>]*w:instr="([^"]*)"/.exec(xml)?.[1] ??
    ''
  const keyword = instr.trim().split(/\s+/)[0]?.toUpperCase() ?? ''
  if (keyword && FIELD_LABELS[keyword]) return FIELD_LABELS[keyword]
  if (keyword) return `Field (${keyword})`
  // No field code in this paragraph: it only closes a field started earlier
  // (e.g. the paragraph holding the TOC's fldChar end + page break).
  const hasEnd = xml.includes('fldCharType="end"') || xml.includes("fldCharType='end'")
  const hasBegin = xml.includes('fldCharType="begin"') || xml.includes("fldCharType='begin'")
  if (hasEnd && !hasBegin) {
    return xml.includes('w:type="page"') ? 'Field end marker + page break' : 'Field end marker'
  }
  return 'Field (TOC/page number/etc.)'
}
