// Header/footer variant resolution and page-number sequences / formats.
import type { HeaderFooter, HfImage, HfPartInfo, SectionInfo } from '@genoffice/docx-engine'

import type { PageSlice } from './pagination-types'

export type HfVariant = 'default' | 'first' | 'even'

/** the parsed strips of a one-section document, one slot per variant */
export interface HfSet {
  header: HeaderFooter | null
  footer: HeaderFooter | null
  headerFirst: HeaderFooter | null
  footerFirst: HeaderFooter | null
  headerEven: HeaderFooter | null
  footerEven: HeaderFooter | null
  titlePg: boolean
  evenOddHf: boolean
  /** images in each variant part (logos etc., display-only) */
  images?: Partial<
    Record<
      'header' | 'footer' | 'headerFirst' | 'footerFirst' | 'headerEven' | 'footerEven',
      HfImage[]
    >
  >
}

export interface PageHf {
  header: HeaderFooter | null
  footer: HeaderFooter | null
  headerImages?: HfImage[]
  footerImages?: HfImage[]
}

export type HfResolve = (kind: 'header' | 'footer') => {
  value: HeaderFooter | null
  images?: HfImage[]
}

/**
 * The strips a page shows. The resolver carries the pending state (strip edits,
 * Link to Previous) and wins whenever the host supplies one; the parsed set is
 * the fallback for a preview mounted without it.
 */
export function pageHfStrips(
  variant: HfVariant,
  resolve: HfResolve | undefined,
  set: HfSet,
): PageHf {
  if (resolve) {
    const h = resolve('header')
    const f = resolve('footer')
    return { header: h.value, footer: f.value, headerImages: h.images, footerImages: f.images }
  }
  const slot = variant === 'default' ? '' : variant === 'first' ? 'First' : 'Even'
  const headerSlot = `header${slot}` as const
  const footerSlot = `footer${slot}` as const
  return {
    header: set[headerSlot],
    footer: set[footerSlot],
    headerImages: set.images?.[headerSlot],
    footerImages: set.images?.[footerSlot],
  }
}

export interface SectionHfRefs {
  header: Partial<Record<HfVariant, string>>
  footer: Partial<Record<HfVariant, string>>
}

/** Effective header/footer refs per section: undefined variants inherit from earlier sections */
export function effectiveHfRefs(sections: SectionInfo[]): SectionHfRefs[] {
  const out: SectionHfRefs[] = []
  let prev: SectionHfRefs = { header: {}, footer: {} }
  for (const s of sections) {
    const cur: SectionHfRefs = {
      header: { ...prev.header, ...s.headerRefs },
      footer: { ...prev.footer, ...s.footerRefs },
    }
    out.push(cur)
    prev = cur
  }
  return out
}

function hfHasContent(hf: HeaderFooter | HfPartInfo | null | undefined): boolean {
  if (!hf) return false
  if ((hf as HeaderFooter).pageNumber || (hf as HfPartInfo).hasPageNumber) return true
  if (hf.text.trim()) return true
  if ((hf as HfPartInfo).images?.length) return true
  return (hf.paras ?? []).some((p) => p.runs.some((r) => r.text.trim()))
}

/**
 * Direct (no-preview) PDF export prints the edit canvas, where the header/footer exists
 * once per document instead of once per page — so any printable header/footer must force
 * the preview-merge export path. Empty parts don't count.
 */
export function hasPrintableHeaderFooter(input: {
  /** local edit state: global header/footer, active variants, per-section edits */
  edited: Array<HeaderFooter | null | undefined>
  sections: SectionInfo[]
  hfParts?: Record<string, HfPartInfo>
  evenOddHf?: boolean
}): boolean {
  if (input.edited.some(hfHasContent)) return true
  const refs = effectiveHfRefs(input.sections)
  return refs.some((ref, i) => {
    const variants: HfVariant[] = ['default']
    if (input.sections[i]?.titlePg) variants.push('first')
    if (input.evenOddHf) variants.push('even')
    return variants.some((v) => {
      const h = ref.header[v]
      const f = ref.footer[v]
      return (
        hfHasContent(h ? input.hfParts?.[h] : null) || hfHasContent(f ? input.hfParts?.[f] : null)
      )
    })
  })
}

/** Header/footer variant a page renders: the titlePg first page of its section
 *  takes 'first', an even page under evenAndOddHeaders takes 'even' */
export function hfVariantOf(
  titlePg: boolean | undefined,
  firstOfSection: boolean,
  evenOddHf: boolean,
  pageNo: number,
): 'first' | 'even' | 'default' {
  if (titlePg && firstOfSection) return 'first'
  if (evenOddHf && pageNo % 2 === 0) return 'even'
  return 'default'
}

/** Displayed page number per page: restart at the section's pgNumType w:start, otherwise continue;
 *  evenPage/oddPage section breaks skip a number to fix parity (when not restarting) */
export function pageNumbers(slices: PageSlice[], sections: SectionInfo[]): number[] {
  const nums: number[] = []
  let n = 0
  let prevSection = -1
  for (const slice of slices) {
    if (slice.section !== prevSection) {
      const sec = sections[slice.section]
      const start = sec?.pageNumberStart
      n = start ?? n + 1
      if (start === undefined && prevSection !== -1) {
        if (sec?.startType === 'evenPage' && n % 2 === 1) n += 1
        if (sec?.startType === 'oddPage' && n % 2 === 0) n += 1
      }
      prevSection = slice.section
    } else {
      n += 1
    }
    nums.push(n)
  }
  return nums
}

const ROMAN: Array<[number, string]> = [
  [1000, 'M'],
  [900, 'CM'],
  [500, 'D'],
  [400, 'CD'],
  [100, 'C'],
  [90, 'XC'],
  [50, 'L'],
  [40, 'XL'],
  [10, 'X'],
  [9, 'IX'],
  [5, 'V'],
  [4, 'IV'],
  [1, 'I'],
]

/** Page numbers arrive from the file (w:pgNumType w:start) and AI setups: a
 *  non-finite or huge value must not hang toRoman/toLetters or OOM toGreek's
 *  repeat(). Bound to a million pages (far beyond real documents). */
const MAX_PAGE_NUMBER = 1_000_000

function boundedPageNumber(n: number): number {
  if (!Number.isFinite(n)) return 1
  const floored = Math.floor(n)
  if (floored < 1) return 1
  return Math.min(floored, MAX_PAGE_NUMBER)
}

function toRoman(n: number): string {
  let out = ''
  let rest = boundedPageNumber(n)
  for (const [v, s] of ROMAN) {
    while (rest >= v) {
      out += s
      rest -= v
    }
  }
  return out
}

/** 1→A ... 26→Z, 27→AA (Word letter numbering) */
function toLetters(n: number): string {
  let out = ''
  let rest = boundedPageNumber(n)
  while (rest > 0) {
    rest -= 1
    out = String.fromCharCode(65 + (rest % 26)) + out
    rest = Math.floor(rest / 26)
  }
  return out
}

function toGreek(n: number, base: number): string {
  const bounded = boundedPageNumber(n)
  if (bounded < 1) return String(bounded)
  // 24-letter alphabet (no final sigma); 25 -> αα, skip the ς slot from the 18th letter on
  const idx = ((bounded - 1) % 24) + 1
  const repeat = Math.floor((bounded - 1) / 24) + 1
  return String.fromCharCode(base + idx - 1 + (idx >= 18 ? 1 : 0)).repeat(repeat)
}

const CN_DIGITS = '〇一二三四五六七八九'

function toChinese(n: number): string {
  if (n < 10) return CN_DIGITS[n]
  if (n < 20) return `十${n % 10 ? CN_DIGITS[n % 10] : ''}`
  if (n < 100) return `${CN_DIGITS[Math.floor(n / 10)]}十${n % 10 ? CN_DIGITS[n % 10] : ''}`
  return String(n)
    .split('')
    .map((d) => CN_DIGITS[Number(d)])
    .join('')
}

/** Display the page number in the section's number format (w:pgNumType w:fmt) */
export function formatPageNumber(n: number, fmt?: string): string {
  switch (fmt) {
    case 'numberInDash':
      return `- ${n} -`
    case 'lowerLetter':
      return toLetters(n).toLowerCase()
    case 'upperLetter':
      return toLetters(n)
    case 'lowerRoman':
      return toRoman(n).toLowerCase()
    case 'upperRoman':
      return toRoman(n)
    case 'lowerGreek':
      return toGreek(n, 0x3b1)
    case 'upperGreek':
      return toGreek(n, 0x391)
    case 'chineseCounting':
    case 'chineseCountingThousand':
      return toChinese(n)
    default:
      return String(n)
  }
}
