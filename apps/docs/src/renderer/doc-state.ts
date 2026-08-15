/**
 * Shared document-state types and header/footer helpers used by App.tsx and
 * the extracted action modules (file-actions, review-actions, …).
 */
import type { HeaderFooter, HfPartInfo, ParsedDocFull } from '@genoffice/docx-engine'

/** first-page / even-page header & footer variants */
export type HfVariantKey = 'headerFirst' | 'footerFirst' | 'headerEven' | 'footerEven'
export type HfVariantsState = Record<HfVariantKey, HeaderFooter | null>
export type HfView = 'default' | 'first' | 'even'

export const EMPTY_HF_VARIANTS: HfVariantsState = {
  headerFirst: null,
  footerFirst: null,
  headerEven: null,
  footerEven: null,
}

export interface DocState {
  parsed: ParsedDocFull
  /** null until a new document is saved for the first time */
  filePath: string | null
  fileName: string
  hash: string
  /** created from the built-in blank template (its numbering ids are known) */
  isBlank?: boolean
}

/** Pending numbering definitions to append (saved via SaveOptions.numbering) */
export interface PendingNumbering {
  newDefs: Array<{
    numId: string
    kind: 'bullet' | 'ordered'
    levels?: import('@genoffice/docx-engine').CustomNumberingLevel[]
  }>
  restartNums: Array<{
    numId: string
    abstractNumId: string
    startOverrides: Record<number, number>
  }>
}

export function hfFromPart(part: HfPartInfo | null | undefined): HeaderFooter | null {
  if (!part || (!part.text && !part.hasPageNumber && part.paras.length === 0)) return null
  return {
    text: part.text,
    pageNumber: part.hasPageNumber,
    paras: part.paras.length > 0 ? part.paras : undefined,
  }
}

/**
 * Variant a resting canvas header/footer area shows (no variant chip picked):
 * the header area sits on page 1 (titlePg -> first-page variant, blank when
 * that part is absent — Word semantics), the footer area on the last page.
 */
export function restingHfAreaVariant(
  kind: 'header' | 'footer',
  opts: { titlePg: boolean; evenOddHf: boolean; pageCount: number; lastPageNo?: number },
): HfView {
  if (kind === 'header') return opts.titlePg ? 'first' : 'default'
  if (opts.titlePg && opts.pageCount <= 1) return 'first'
  if (opts.evenOddHf && (opts.lastPageNo ?? opts.pageCount) % 2 === 0) return 'even'
  return 'default'
}

export function hfVariantsFromParsed(parsed: ParsedDocFull): HfVariantsState {
  return {
    headerFirst: hfFromPart(parsed.headerFirst),
    footerFirst: hfFromPart(parsed.footerFirst),
    headerEven: hfFromPart(parsed.headerEven),
    footerEven: hfFromPart(parsed.footerEven),
  }
}
