/// Excel's fit-to-page scale: the printout shrinks (never grows) until it
/// fits in `fitToWidth` pages across and `fitToHeight` pages tall; 0 leaves
/// that axis unconstrained, and both 0 means 100%. Rows never split across
/// pages and the repeated header (heading strip + title rows) reappears on
/// every page, so the height axis is solved by simulating the pagination
/// rather than by plain division. A multi-area print area prints each area
/// on its own pages (Excel too), so the page budget applies per area and the
/// sheet-wide scale is the smallest one any area needs.

/// Print-space heights of one print area (each area starts a new page).
export interface PrintAreaHeights {
  /// Height repeated at the top of every page of this area.
  readonly repeatedHeightPt: number
  /// Body row heights in print order.
  readonly rowHeightsPt: readonly number[]
}

export interface FitToPageInput {
  /// Paper minus the side margins / minus the top and bottom margins.
  readonly printableWidthPt: number
  readonly printableHeightPt: number
  readonly fitToWidth: number
  readonly fitToHeight: number
  /// The widest area (heading strip included).
  readonly contentWidthPt: number
  readonly areas: readonly PrintAreaHeights[]
}

export const MIN_PRINT_SCALE = 0.1
export const MAX_PRINT_SCALE = 2

/// Step the height search takes; Excel's own fit results are whole percents.
const SCALE_STEP = 0.005

export function fitToPageScale(input: FitToPageInput): number {
  let scale = 1
  if (input.fitToWidth > 0 && input.contentWidthPt > 0 && input.printableWidthPt > 0) {
    scale = Math.min(scale, (input.printableWidthPt * input.fitToWidth) / input.contentWidthPt)
  }
  if (input.fitToHeight > 0 && input.printableHeightPt > 0) {
    const areas = input.areas.filter(
      (area) => area.repeatedHeightPt + sumHeights(area.rowHeightsPt) > 0,
    )
    // Plain division is an upper bound: pagination waste (unsplittable
    // rows, repeated headers) can only add pages, so step down from there
    // until every area's simulated page count fits its budget.
    for (const area of areas) {
      const total = area.repeatedHeightPt + sumHeights(area.rowHeightsPt)
      scale = Math.min(scale, (input.printableHeightPt * input.fitToHeight) / total)
    }
    scale = Math.max(scale, MIN_PRINT_SCALE)
    const overflows = (capacityPt: number): boolean =>
      areas.some((area) => countPages([area], capacityPt) > input.fitToHeight)
    while (scale > MIN_PRINT_SCALE && overflows(input.printableHeightPt / scale)) {
      scale = Math.max(MIN_PRINT_SCALE, scale - SCALE_STEP)
    }
  }
  return Math.min(1, Math.max(MIN_PRINT_SCALE, scale))
}

/// Pages a set of areas occupies when a page holds `capacityPt` of content:
/// each area starts a new page, rows never split (a row taller than the
/// remaining space moves to the next page), and the repeated header takes
/// its share of every page.
export function countPages(areas: readonly PrintAreaHeights[], capacityPt: number): number {
  let pages = 0
  for (const area of areas) {
    pages += 1
    let used = area.repeatedHeightPt
    for (const row of area.rowHeightsPt) {
      if (used + row > capacityPt && used > area.repeatedHeightPt) {
        pages += 1
        used = area.repeatedHeightPt
      }
      used += row
    }
  }
  return pages
}

function sumHeights(heights: readonly number[]): number {
  return heights.reduce((total, height) => total + height, 0)
}
