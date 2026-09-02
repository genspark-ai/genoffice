// Field-code display (PAGE, TOC, REF, ...) and TOC entry numbering.
import { computeListMarkers, type ListItemRef } from './list-markers'
import { decodeEntities, lineTwipsOf, plainText } from './parse-xml-text'
import type { Block, FieldDisplay, NumberingDef, StyleInfo } from './types'

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

export function fieldDisplayOf(
  xml: string,
  styles?: Map<string, StyleInfo>,
): FieldDisplay | undefined {
  const styleId = /<w:pStyle w:val="([^"]+)"/.exec(xml)?.[1] ?? ''
  const tocLevel = tocLevelOf(styleId, styles)
  if (tocLevel !== null) {
    // TOC entry: title <tab with dot leader> page number. The page number
    // follows the LAST tab — entries like "1.1.<tab>Title<tab>7" put a leading
    // outline number at the first tab stop, not the page number.
    const segs: string[] = ['']
    const re = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:tab\/>/g
    let m: RegExpExecArray | null
    while ((m = re.exec(xml)) !== null) {
      if (m[0] === '<w:tab/>') segs.push('')
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
    const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(xml)?.[0] ?? ''
    const spacingAttrs = /<w:spacing ([^/>]*)\/>/.exec(pPr)?.[1] ?? ''
    const line = lineTwipsOf(/w:line="([^"]+)"/.exec(spacingAttrs)?.[1])
    // OOXML defaults w:lineRule to auto when omitted
    const lineRule = (/w:lineRule="(auto|atLeast|exact)"/.exec(spacingAttrs)?.[1] ?? 'auto') as
      'auto' | 'atLeast' | 'exact'
    // font size from visible result runs only: field-machinery runs
    // (fldChar/instrText) often carry the target heading's size and would
    // inflate the whole line
    let sz = 0
    const runRe = /<w:r(?:\s[^>]*)?>([\s\S]*?)<\/w:r>/g
    let run: RegExpExecArray | null
    while ((run = runRe.exec(xml)) !== null) {
      if (!/<w:t(?:\s|>)/.test(run[1]) || run[1].includes('<w:instrText')) continue
      const v = parseInt(/<w:sz w:val="(\d+)"/.exec(run[1])?.[1] ?? '', 10)
      if (v > sz) sz = v
    }
    return {
      kind: 'tocLine',
      left,
      right: decodeEntities(right).trim(),
      level: tocLevel,
      ...(num ? { num } : {}),
      ...(anchor ? { anchor } : {}),
      ...(sz > 0 ? { szHalfPoints: sz } : {}),
      ...(line > 0 && lineRule
        ? {
            lineRule,
            lineRawTwips: line,
            ...(lineRule === 'auto' ? { lineSpacing: Math.round((line / 240) * 100) / 100 } : {}),
          }
        : {}),
    }
  }
  const visible = plainText(xml).trim()
  if (visible === '' && /<w:br\s[^>]*w:type="page"/.test(xml)) {
    return { kind: 'pageBreak' }
  }
  if (visible !== '') {
    // face/size of the visible result runs (same rule as tocLine): without
    // them the passthrough div inherits the document default and a SimSun
    // field paragraph mis-snaps to a double cell on a typed line grid
    // dominant size = the size covering the most text: a manual drop-cap letter
    // (one 48pt "L" before 10pt body) must not inflate the whole field's strut
    // (real_run2/47 rendered the entire paragraph at 48pt, 10 -> 17 pages)
    let font: string | undefined
    const sizedRuns: Array<{ text: string; szHalfPoints?: number }> = []
    const szWeights = new Map<number, number>()
    const runRe = /<w:r(?:\s[^>]*)?>([\s\S]*?)<\/w:r>/g
    let run: RegExpExecArray | null
    while ((run = runRe.exec(xml)) !== null) {
      if (!/<w:t(?:\s|>)/.test(run[1]) || run[1].includes('<w:instrText')) continue
      const text = decodeEntities(
        Array.from(run[1].matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g), (m) => m[1]).join(''),
      )
      const v = parseInt(/<w:sz w:val="(\d+)"/.exec(run[1])?.[1] ?? '', 10)
      sizedRuns.push({ text, ...(v > 0 ? { szHalfPoints: v } : {}) })
      // unsized runs vote for the inherited default (key 0): one explicit
      // drop-cap letter must not out-vote a body of default-sized text
      const key = v > 0 ? v : 0
      szWeights.set(key, (szWeights.get(key) ?? 0) + Math.max(text.length, 1))
      if (!font) {
        const fonts = /<w:rFonts [^/>]*/.exec(run[1])?.[0] ?? ''
        font = /w:eastAsia="([^"]+)"/.exec(fonts)?.[1] ?? /w:ascii="([^"]+)"/.exec(fonts)?.[1]
      }
    }
    let sz = 0
    let szWeight = -1
    for (const [v, w] of szWeights) {
      if (w > szWeight || (w === szWeight && v > sz)) {
        sz = v
        szWeight = w
      }
    }
    // sized spans only for mixed sizes, and only when the runs reproduce the
    // visible text (tabs/breaks fall back to the plain string)
    const norm = (t: string) => t.replace(/\s+/g, ' ').trim()
    const mixedRuns =
      szWeights.size > 1 && norm(sizedRuns.map((r) => r.text).join('')) === norm(visible)
        ? sizedRuns.filter((r) => r.text !== '')
        : undefined
    // explicit paragraph alignment: the passthrough div would inherit the
    // document default (justify in CJK docs) and stretch short lines
    const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(xml)?.[0] ?? ''
    const jc = /<w:jc w:val="([^"]+)"/.exec(pPr)?.[1]
    const align =
      jc === 'left' || jc === 'start'
        ? 'left'
        : jc === 'right' || jc === 'end'
          ? 'right'
          : jc === 'center'
            ? 'center'
            : undefined
    // explicit line spacing (same extraction as tocLine): the renderer must
    // not collapse a 1.5x field paragraph to single-spacing
    const spacingAttrs = /<w:spacing ([^/>]*)\/>/.exec(pPr)?.[1] ?? ''
    const line = lineTwipsOf(/w:line="([^"]+)"/.exec(spacingAttrs)?.[1])
    const lineRule = (/w:lineRule="(auto|atLeast|exact)"/.exec(spacingAttrs)?.[1] ?? 'auto') as
      'auto' | 'atLeast' | 'exact'
    return {
      kind: 'text',
      left: visible,
      ...(sz > 0 ? { szHalfPoints: sz } : {}),
      ...(mixedRuns ? { runs: mixedRuns } : {}),
      ...(font ? { fontFamily: font } : {}),
      ...(align ? { align } : {}),
      ...(line > 0 && lineRule
        ? {
            lineRule,
            lineRawTwips: line,
            ...(lineRule === 'auto' ? { lineSpacing: Math.round((line / 240) * 100) / 100 } : {}),
          }
        : {}),
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
    if (block.list?.numId) {
      items.push({ numId: block.list.numId, ilvl: block.list.ilvl })
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
  if (xml.includes('fldCharType="end"') && !xml.includes('fldCharType="begin"')) {
    return xml.includes('w:type="page"') ? 'Field end marker + page break' : 'Field end marker'
  }
  return 'Field (TOC/page number/etc.)'
}
