/**
 * Word opens a page's header when you double-click anywhere in its top margin,
 * and its footer anywhere in its bottom margin, whether or not the part has
 * content yet. The strips alone are no target for that: an empty header has
 * no strip on the first page, and a short one covers a sliver of the margin.
 * These helpers place each page's margin bands on the canvas so the
 * double-click can be resolved from its Y.
 */
import { pagePaperEdges, type PageCut } from '../pagination-cuts'

/** one page's top and bottom margin, in flow px from the editor box top (zoom 1) */
export interface PageMarginBand {
  page: number
  headerTop: number
  headerBottom: number
  footerTop: number
  footerBottom: number
}

/**
 * Margin bands of every page. The top margin is the one the gap widget opens
 * the page with (after the header push), the first page's the editor's own
 * top padding. The bottom margin comes from `bottomMarginOf`: the gap's
 * bottom part also holds the page's unused body height, which belongs to the
 * body (Word puts the caret at the end of the last line there).
 */
export function pageMarginBands(
  cuts: readonly PageCut[],
  firstMarginTop: number,
  flowHeight: number,
  bottomMarginOf: (page: number) => number,
): PageMarginBand[] {
  if (cuts.length === 0) return []
  const tops: number[] = [0]
  const topMargins: number[] = [firstMarginTop]
  const bottoms: number[] = []
  for (const edge of pagePaperEdges(cuts, firstMarginTop)) {
    // a page without a gap (parity twin, dropped boundary) shares its neighbour's paper
    while (tops.length < edge.page) {
      tops.push(tops[tops.length - 1]!)
      topMargins.push(topMargins[topMargins.length - 1]!)
      bottoms.push(edge.prevBottom)
    }
    bottoms.push(edge.prevBottom)
    tops.push(edge.top)
    topMargins.push(edge.gap.marginTopPx)
  }
  while (tops.length < cuts.length) {
    tops.push(tops[tops.length - 1]!)
    topMargins.push(topMargins[topMargins.length - 1]!)
  }
  while (bottoms.length < cuts.length) bottoms.push(flowHeight)
  return cuts.map((_, page) => ({
    page,
    headerTop: tops[page]!,
    headerBottom: tops[page]! + topMargins[page]!,
    footerTop: bottoms[page]! - bottomMarginOf(page),
    footerBottom: bottoms[page]!,
  }))
}

/** the header or footer whose margin band holds flow Y `y`, if any */
export function hfBandAt(
  bands: readonly PageMarginBand[],
  y: number,
): { kind: 'header' | 'footer'; page: number } | null {
  for (const b of bands) {
    if (y >= b.headerTop && y < b.headerBottom) return { kind: 'header', page: b.page }
    if (y >= b.footerTop && y < b.footerBottom) return { kind: 'footer', page: b.page }
  }
  return null
}

/** targets a margin double-click leaves alone: a picture, a control or a text box keeps its own double-click */
export const HF_MARGIN_SKIP =
  'img, svg, canvas, a[href], button, input, textarea, select, .doc-textbox'
