/// Page bookkeeping for Excel's differentFirst / differentOddEven header
/// and footer variants. Chromium's printToPDF takes ONE header/footer
/// template pair per call, so the export prints the whole sheet with the
/// odd-page templates and then, when a variant is active, re-prints page 1
/// (first-page templates) and/or the even pages (even templates) with
/// `pageRanges`; the final PDF takes each page from the pass that owns it.
/// Chromium numbers the pages of a ranged print by their document index, so
/// `&P`/`&N` stay correct in every pass. Pure — no Electron here.

export type PageVariant = 'odd' | 'even' | 'first'

export interface VariantFlags {
  readonly hasFirst: boolean
  readonly hasEven: boolean
}

/// Which pass prints the given 1-based page.
export function variantForPage(pageNumber: number, flags: VariantFlags): PageVariant {
  if (pageNumber === 1 && flags.hasFirst) return 'first'
  if (pageNumber % 2 === 0 && flags.hasEven) return 'even'
  return 'odd'
}

/// `pageRanges` for the even pass of a `total`-page document ('' when none).
export function evenPageRanges(total: number): string {
  const pages: string[] = []
  for (let page = 2; page <= total; page += 2) pages.push(String(page))
  return pages.join(',')
}

export interface StitchStep {
  /// 1-based document page.
  readonly page: number
  readonly source: PageVariant
  /// Page index within the source pass's PDF.
  readonly index: number
}

/// Page-by-page assembly plan: the odd pass holds every page at its
/// document index; the first pass is page 1 alone; the even pass holds
/// pages 2, 4, 6, … in order.
export function stitchPlan(total: number, flags: VariantFlags): StitchStep[] {
  const steps: StitchStep[] = []
  for (let page = 1; page <= total; page += 1) {
    const source = variantForPage(page, flags)
    steps.push({
      page,
      source,
      index: source === 'first' ? 0 : source === 'even' ? page / 2 - 1 : page - 1,
    })
  }
  return steps
}
