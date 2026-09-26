import type { ParaFormat, StyleDisplay } from '@genoffice/docx-engine'
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

/** Parse-layer pagination constraints of a paragraph block (direct pPr, else its style).
 *  A direct page break is left out on purpose: the element's page-break-before class
 *  already carries it, and an unchecked one must not be restored from the parsed block. */
export function paraPaginationMeta(
  format: ParaFormat | undefined,
  style: StyleDisplay | undefined,
): BlockMeta | undefined {
  const keepNext = format?.keepNext ?? style?.keepNext
  const keepLines = format?.keepLines ?? style?.keepLines
  const breakBefore = format?.pageBreakBefore === undefined && style?.pageBreakBefore
  const widowOff = (format?.widowControl ?? style?.widowControl) === false
  const noLineNo = format?.suppressLineNumbers ?? style?.suppressLineNumbers
  const paraStyle: DirectParaFlags = {}
  for (const key of KEYS) if (typeof style?.[key] === 'boolean') paraStyle[key] = style[key]
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
