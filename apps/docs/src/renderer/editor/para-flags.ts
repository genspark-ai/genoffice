import type { DocDefaults, ParaFormat, StyleDisplay } from '@genoffice/docx-engine'
import type { BlockMeta } from '../pagination-types'

export interface DirectParaFlags {
  keepNext?: boolean
  keepLines?: boolean
  widowControl?: boolean
  suppressLineNumbers?: boolean
}

const KEYS = ['keepNext', 'keepLines', 'widowControl', 'suppressLineNumbers'] as const

/** Pagination flags set directly on the paragraph (blockAttrs' data-para payload);
 *  undefined = not set on the pPr, the style's value applies. */
export function directParaFlags(el: Element | undefined | null): DirectParaFlags {
  const raw = el?.getAttribute('data-para')
  if (!raw) return {}
  let parsed: unknown
  try {
    parsed = JSON.parse(raw)
  } catch {
    return {}
  }
  if (typeof parsed !== 'object' || parsed === null) return {}
  const out: DirectParaFlags = {}
  for (const key of KEYS) {
    const v = (parsed as Record<string, unknown>)[key]
    if (typeof v === 'boolean') out[key] = v
  }
  return out
}

/** Pagination constraints of a paragraph the parser emitted without a format model
 *  (picture-only paragraphs become image blocks), read from its own pPr so a keepNext
 *  picture still chains to its heading and the next picture (Word). The block is no
 *  paragraph node, nothing edits its flags: they also fill the style slot the element
 *  path reads. */
export function imageParaPaginationMeta(
  xml: string | null | undefined,
  styleOf: (styleId: string | undefined) => StyleDisplay | undefined,
  docDefaults?: Pick<DocDefaults, 'widowControl'>,
): BlockMeta | undefined {
  const own = paraFlagsFromXml(xml)
  const meta = paraPaginationMeta(own.format, styleOf(own.styleId), docDefaults)
  if (!meta || !own.format) return meta
  const paraStyle = { ...meta.paraStyle }
  for (const key of KEYS) if (typeof own.format[key] === 'boolean') paraStyle[key] = own.format[key]
  return { ...meta, paraStyle }
}

/** Direct pagination flags and style id of a paragraph's own pPr (parse-layer XML). */
export function paraFlagsFromXml(xml: string | null | undefined): {
  format?: ParaFormat
  styleId?: string
} {
  const pPr = xml
    ? /^<w:p\b[^>]*>\s*<w:pPr>([\s\S]*?)<\/w:pPr>/.exec(
        xml.replace(/<w:pPrChange\b[\s\S]*?<\/w:pPrChange>/g, ''),
      )?.[1]
    : undefined
  if (!pPr) return {}
  const onOff = (name: string): boolean | undefined => {
    const m = new RegExp(`<w:${name}\\b([^>]*?)/?>`).exec(pPr)
    return m ? !/w:val="(?:0|false|off)"/.test(m[1]) : undefined
  }
  const format: ParaFormat = {}
  const keepNext = onOff('keepNext')
  const keepLines = onOff('keepLines')
  const pageBreakBefore = onOff('pageBreakBefore')
  const widowControl = onOff('widowControl')
  if (keepNext !== undefined) format.keepNext = keepNext
  if (keepLines !== undefined) format.keepLines = keepLines
  if (pageBreakBefore !== undefined) format.pageBreakBefore = pageBreakBefore
  if (widowControl !== undefined) format.widowControl = widowControl
  const styleId = /<w:pStyle\b[^>]*w:val="([^"]+)"/.exec(pPr)?.[1]
  return { format, ...(styleId ? { styleId } : {}) }
}

/** Parse-layer pagination constraints of a paragraph block (direct pPr, else its style).
 *  A direct page break is left out on purpose: the element's page-break-before class
 *  already carries it, and an unchecked one must not be restored from the parsed block. */
export function paraPaginationMeta(
  format: ParaFormat | undefined,
  style: StyleDisplay | undefined,
  docDefaults?: Pick<DocDefaults, 'widowControl'>,
): BlockMeta | undefined {
  const keepNext = format?.keepNext ?? style?.keepNext
  const keepLines = format?.keepLines ?? style?.keepLines
  const breakBefore = format?.pageBreakBefore === undefined && style?.pageBreakBefore
  const styleWidow = style?.widowControl ?? docDefaults?.widowControl
  const widowOff = (format?.widowControl ?? styleWidow) === false
  const noLineNo = format?.suppressLineNumbers ?? style?.suppressLineNumbers
  const paraStyle: DirectParaFlags = {}
  for (const key of KEYS) if (typeof style?.[key] === 'boolean') paraStyle[key] = style[key]
  if (paraStyle.widowControl === undefined && styleWidow === false) paraStyle.widowControl = false
  if (
    !keepNext &&
    !keepLines &&
    !breakBefore &&
    !widowOff &&
    !noLineNo &&
    Object.keys(paraStyle).length === 0
  )
    return undefined
  return {
    ...(keepNext ? { keepNext: true } : {}),
    ...(keepLines ? { keepLines: true } : {}),
    ...(noLineNo ? { suppressLineNumbers: true } : {}),
    ...(breakBefore ? { breakBefore: true } : {}),
    ...(widowOff ? { widowControl: false as const } : {}),
    paraStyle,
  }
}

/** A table's cells follow a document-wide widowControl off (Word 2013+ layout)
 *  unless a cell paragraph turns it back on, directly or through its pStyle. */
export function tableWidowOff(
  xml: string,
  docWidowControl: boolean | undefined,
  styleWidowControl: (styleId: string) => boolean | undefined,
): boolean {
  if (docWidowControl !== false) return false
  if (/<w:widowControl(\/>|\s+w:val="(1|true|on)")/.test(xml)) return false
  for (const m of xml.matchAll(/<w:pStyle w:val="([^"]+)"/g))
    if (styleWidowControl(m[1]) === true) return false
  return true
}

/** Effective pagination flags of a measured block: the element's direct flags, then
 *  the style chain; a block without a paragraph element falls back to the merged meta.
 *  Reading the direct part off the element (not the parsed block) is what makes a
 *  dialog edit or Ctrl+Q re-paginate before the document is saved. */
export function effectiveParaFlags(
  el: Element | undefined | null,
  meta: BlockMeta | undefined,
): DirectParaFlags {
  const direct = directParaFlags(el)
  const inherited = el && meta?.paraStyle ? meta.paraStyle : meta
  const out: DirectParaFlags = {}
  for (const key of KEYS) {
    const v = direct[key] ?? inherited?.[key]
    if (typeof v === 'boolean') out[key] = v
  }
  return out
}
