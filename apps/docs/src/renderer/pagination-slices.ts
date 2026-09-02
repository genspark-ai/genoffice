// The slicing engine: greedy page breaking of measured blocks into page slices
// (F2 model with table-row and line-level placement), plus page lookups.
import { FOOTNOTE_SEPARATOR_H } from './line-metrics'
import type {
  BlockBox,
  PageSlice,
  SectionGeom,
  SliceOutputs,
  TableRowBox,
} from './pagination-types'

/**
 * Page owning a page-pinned float: the page its anchor lands on (anchorTop and
 * slices share gapless virtual coordinates); out-of-range anchors clamp.
 */
export function pinnedFloatPage(slices: PageSlice[], anchorTop: number): number {
  const idx = slices.findIndex((s) => anchorTop >= s.start && anchorTop < s.end)
  if (idx >= 0) return idx
  return anchorTop < (slices[0]?.start ?? 0) ? 0 : Math.max(0, slices.length - 1)
}

/** minimum room a bandKeep block needs on the page (its anchor line, px):
 *  with less, the anchor itself moves to the next page like any line */
const BAND_KEEP_MIN_H = 14

export function computePageSlices(
  blocks: BlockBox[],
  contentHeight: number,
  totalHeight: number,
): PageSlice[] {
  return computeSectionedSlices(blocks, [{ contentHeight, forceBreak: false }], totalHeight)
}

/**
 * Line-level cut point for a page-crossing block: the last line boundary before the
 * page limit that satisfies widow/orphan constraints. Constraints: when the block
 * starts on this page, keep ≥ splitMinLines lines at the head; keep ≥ splitMinLines
 * lines in the tail after the cut. Returns null when there are no line boundaries or
 * the constraints fail (caller pushes the whole block / falls back to pixel cut).
 */
function lineCut(block: BlockBox, pageStart: number, limit: number): number | null {
  const offs = block.lineOffsets
  if (!offs || offs.length === 0) return null
  const minLines = block.splitMinLines ?? 1
  const headMinIdx = block.top >= pageStart ? minLines - 1 : 0
  const tailMaxIdx = offs.length - minLines
  let cut: number | null = null
  for (let k = headMinIdx; k <= tailMaxIdx; k++) {
    const y = block.top + offs[k]
    if (y > limit) break
    if (y > pageStart) cut = y
  }
  return cut
}

export function computeSectionedSlices(
  blocks: BlockBox[],
  geoms: SectionGeom[],
  totalHeight: number,
): PageSlice[] {
  const total = Math.max(totalHeight, 0)
  // see computeSectionedSlicesF2: block-less sections (lone sectPr chips)
  // still claim their own pages ahead of true page starts
  const sectionHasBlocks = new Set<number>()
  for (const b of blocks) sectionHasBlocks.add(b.section ?? 0)
  const firstSection = blocks[0]?.section ?? 0
  if (geoms.length === 0 || geoms.every((g) => g.contentHeight <= 0)) {
    return [{ start: 0, end: total, section: firstSection }]
  }
  const geomOf = (s: number) => geoms[Math.max(0, Math.min(s, geoms.length - 1))]
  const emptySectionClaimsPage = (s: number) => {
    // only true page starts keep a blank page for an empty preceding section;
    // a promoted nextColumn (single-column, n#750255) or a continuous size
    // change absorbs it (Word skips the blank first page there)
    const st = geomOf(s).startType
    return (
      (st === undefined || st === 'nextPage' || st === 'evenPage' || st === 'oddPage') &&
      !sectionHasBlocks.has(s - 1)
    )
  }
  const initSection = firstSection > 0 && emptySectionClaimsPage(1) ? 0 : firstSection

  const starts: Array<{ y: number; section: number }> = [{ y: 0, section: initSection }]
  let pageStart = 0
  let curSection = initSection
  let contentH = Math.max(geomOf(curSection).contentHeight, 1)
  let pendingBreak = false
  const newPage = (y: number, section: number) => {
    pageStart = y
    starts.push({ y, section })
  }
  for (const block of blocks) {
    const bSection = block.section ?? curSection
    if (bSection !== curSection) {
      for (let s = curSection + 1; s <= bSection; s++) {
        const gs = geomOf(s)
        if (gs.forceBreak && (block.top > pageStart || emptySectionClaimsPage(s))) {
          newPage(block.top, s)
        }
      }
      curSection = bSection
      contentH = Math.max(geomOf(curSection).contentHeight, 1)
    }
    if ((pendingBreak || block.breakBefore) && block.top > pageStart) {
      newPage(block.top, curSection)
    }
    pendingBreak = false
    const bottom = block.top + block.height
    // page-crossing block: with line boundaries, cut in place (with widow/orphan
    // constraints); if not cuttable, push the whole block (or the cuttable block's
    // start) to the next page; blocks with no line boundaries taller than a page fall back to hard pixel cuts
    while (bottom > pageStart + contentH) {
      const cut = lineCut(block, pageStart, pageStart + contentH)
      if (cut !== null) {
        newPage(cut, curSection)
      } else if (block.top > pageStart && (block.height <= contentH || block.lineOffsets?.length)) {
        newPage(block.top, curSection)
      } else {
        newPage(pageStart + contentH, curSection)
      }
    }
    if (block.breakAfter) pendingBreak = true
  }
  if (pendingBreak) newPage(Math.max(total, pageStart), curSection)

  const end = Math.max(total, pageStart)
  return starts.map((s, i) => ({
    start: s.y,
    end: i + 1 < starts.length ? starts[i + 1].y : end,
    section: s.section,
  }))
}

/**
 * F2: line-level page splitting + Word pagination constraint solving (incl. column flow).
 *
 * Coordinates:
 *   - block.top: absolute Y in the content flow (px)
 *   - pageStart: starting Y of the current page in the content flow
 *   - usedInCol: height already placed in the current column (single-column doc = height used on the page)
 *   - fits(h): usedInCol + h <= colH + 0.01
 *
 * Columns (SectionGeom.cols>1): three levels, page → region → column. Each column
 * is a "mini page" (column height = content height − region top); overflow moves to
 * the next column, the last column turns the page; forced page breaks turn the page directly.
 * A continuous section changing column count opens a new region on the same page
 * (section capacity = columns × remaining height).
 *
 * Constraint priority: pageBreakBefore > keepNext chain > keepLines > widowControl
 */
export function computeSectionedSlicesF2(
  blocks: BlockBox[],
  geoms: SectionGeom[],
  totalHeight: number,
  out?: SliceOutputs,
): PageSlice[] {
  if (out?.rowFills) out.rowFills.length = 0
  if (out?.floatVShifts) out.floatVShifts.length = 0
  if (out?.oversizeClips) out.oversizeClips.length = 0
  const total = Math.max(totalHeight, 0)
  // sections whose every block collapsed to a zero-height chip (lone sectPr
  // paragraphs) never appear in the measured blocks; they still claim their
  // own (blank) pages ahead of true page starts
  const sectionHasBlocks = new Set<number>()
  for (const b of blocks) sectionHasBlocks.add(b.section ?? 0)
  const firstSection = blocks[0]?.section ?? 0
  if (geoms.length === 0 || geoms.every((g) => g.contentHeight <= 0)) {
    return [{ start: 0, end: total, section: firstSection }]
  }
  const geomOf = (s: number) => geoms[Math.max(0, Math.min(s, geoms.length - 1))]
  const colsOf = (s: number) => Math.max(1, geomOf(s).cols ?? 1)
  const emptySectionClaimsPage = (s: number) => {
    // only true page starts keep a blank page for an empty preceding section;
    // a promoted nextColumn (single-column, n#750255) or a continuous size
    // change absorbs it (Word skips the blank first page there)
    const st = geomOf(s).startType
    return (
      (st === undefined || st === 'nextPage' || st === 'evenPage' || st === 'oddPage') &&
      !sectionHasBlocks.has(s - 1)
    )
  }
  const initSection = firstSection > 0 && emptySectionClaimsPage(1) ? 0 : firstSection

  type ColEntry = { y: number; repeatHeader?: { top: number; height: number } }
  type Region = { top: number; height: number; section: number; cols: number; entries: ColEntry[] }
  const pages: Array<{ section: number; regions: Region[] }> = []

  let pageStart = 0 // starting Y of the current page (absolute)
  let curSection = initSection
  let contentH = Math.max(geomOf(curSection).contentHeight, 1)
  let regionTop = 0 // current region top (relative to page content-area top)
  let colCount = 1 // column count of the current region
  let colH = contentH // column height of the current region
  let colIdx = 0 // current column index
  let usedInCol = 0 // height used in the current column
  let pendingBreak = false
  let pendingForce = false
  let pendingColBreak = false
  // Word draws one footnote separator per page, not per referencing paragraph:
  // the first footnote-bearing block on a page shrinks the page's usable height
  // by the separator; carried to the next page when that block turns the page
  let pageNoteSepPx = 0
  // flow-coord bottom of the current page's floated blocks: floats consume no
  // column height, but a section/page break right after one must not cut into
  // its band — the closing page keeps the float visible and the next page
  // starts below it (a landscape form built as one page-filling positioned
  // table otherwise collapses to a sliver, prod100r2/109)
  let pageFloatBottom = 0
  // clamp to pageStart too: startPage resets the float bottom, so a second
  // forced start in the same iteration (crossed empty sections, double breaks)
  // must not reopen above the already-clamped previous start (inverted slice)
  const breakY = (y: number) => Math.max(y, pageFloatBottom, pageStart)
  // bandKeep overflow bottom, consumed by the next startPage (see there)
  let pendingBandBottom = 0
  let curBlockNotes = false

  // Safety net: a document legitimately needs at most a few column turns per
  // block (forced breaks, line/row splits) — far beyond that means a placement
  // loop stopped converging. Degrade by treating everything as fitting (the
  // rest piles onto the current page, with a warning) instead of looping forever.
  const maxColumnTurns = Math.max(65536, blocks.length * 8)
  let columnTurns = 0
  let runaway = false

  // first block index of the current region (balancing walks this range for line boundaries)
  let regionStartBi = 0
  // an explicit column break inside the region disqualifies it from balancing
  // ("any effective page breaks stop the balancing act"); natural overflow turns
  // are re-distributed by the balance pass
  let regionBroke = false
  // main-loop block index, visible to the region closures for regionStartBi bookkeeping
  let curBi = 0

  const pushColumn = (y: number, headerH = 0, headerTop = 0) => {
    if (++columnTurns > maxColumnTurns && !runaway) {
      runaway = true
      console.warn(
        `[pagination] column-turn limit ${maxColumnTurns} exceeded at y=${y}; placing remaining content without page breaks`,
      )
    }
    const page = pages[pages.length - 1]
    page.regions[page.regions.length - 1].entries.push({
      y,
      ...(headerH > 0 ? { repeatHeader: { top: headerTop, height: headerH } } : {}),
    })
    usedInCol = headerH
  }
  // open a new region at the current page's regionTop (column count/height per section)
  const openRegion = (y: number, section: number, headerH = 0, headerTop = 0) => {
    colCount = colsOf(section)
    colH = Math.max(contentH - regionTop, 1)
    colIdx = 0
    regionStartBi = curBi
    regionBroke = false
    pages[pages.length - 1].regions.push({
      top: regionTop,
      height: colH,
      section,
      cols: colCount,
      entries: [],
    })
    pushColumn(y, headerH, headerTop)
  }
  const startPage = (y: number, section: number, headerH = 0, headerTop = 0) => {
    // a bandKeep block overflowed the closing page: the next page starts below
    // the true box bottom, or the preview window would cut the box and re-pin it
    if (pendingBandBottom > 0) {
      y = Math.max(y, pendingBandBottom)
      pendingBandBottom = 0
    }
    pageStart = y
    regionTop = 0
    // the section's first page renders the titlePg header/footer variant, so it
    // gets its own capacity (same "first page" rule as sectionFirstPages)
    const firstOfSection = pages.length === 0 || pages[pages.length - 1].section !== section
    const g = geomOf(section)
    contentH = Math.max((firstOfSection ? g.firstContentHeight : undefined) ?? g.contentHeight, 1)
    pages.push({ section, regions: [] })
    openRegion(y, section, headerH, headerTop)
    pageNoteSepPx = curBlockNotes ? FOOTNOTE_SEPARATOR_H : 0
    pageFloatBottom = 0
  }
  // advance on overflow: change column if not the last, turn the page on the last (headerH/headerTop: table header repeated at column top after a table break)
  const newColumn = (y: number, section: number, headerH = 0, headerTop = 0) => {
    if (colIdx + 1 < colCount) {
      colIdx += 1
      pushColumn(y, headerH, headerTop)
    } else {
      startPage(y, section, headerH, headerTop)
    }
  }

  /**
   * Word column balancing: a multi-column region closed mid-page by a continuous
   * column-count change redistributes its single-column content across the columns.
   * Word fills line quotas left to right — target = ceil(visible lines / cols) per
   * column, empty paragraph marks flow but don't count, and widow/orphan atomicity
   * keeps short paragraphs whole (trailing columns may stay empty). Explicit column
   * breaks disable it ("any effective page breaks stop the balancing act"), as do
   * tables (row structure, v1). Mutates the current region's entries/height;
   * returns the balanced region height, or null when not applicable.
   */
  const tryBalanceRegion = (endBi: number, endY: number): number | null => {
    const page = pages[pages.length - 1]
    const region = page.regions[page.regions.length - 1]
    if (region.cols <= 1 || regionBroke || runaway) return null
    const startY = region.entries[0]?.y
    if (startY === undefined || region.entries[0].repeatHeader) return null
    // boundary units: block tops always cuttable; in-block line starts cuttable
    // per widow/orphan atomicity. Each unit carries the counted content height of
    // its piece (0 for empty paragraph marks — they flow but don't add quota)
    type Unit = { y: number; h: number; cut: boolean }
    const units: Unit[] = []
    // trailing whitespace before the closing block (space-after / inter-section
    // spacing) must not inflate the last column's extent
    let contentEnd = startY
    for (let i = regionStartBi; i < endBi; i++) {
      const b = blocks[i]
      if (b.tableRows) return null
      if (b.floated) continue
      if (b.top + b.height < startY + 0.01 || b.top > endY - 0.01) continue
      const bottom = Math.min(b.top + b.height, endY)
      // trailing empty paragraph marks are absorbed (they don't extend the region)
      if (!b.emptyPara) contentEnd = Math.max(contentEnd, bottom)
      const minKeep = b.widowControl !== false ? 2 : 1
      const lbs = b.lineBoxes && b.lineBoxes.length > 0 ? b.lineBoxes : null
      const n = lbs ? lbs.length : 1
      for (let k = 0; k < n; k++) {
        // a block split across the page boundary contributes only its lines within
        // the region (clip to startY); its clipped head is not a cut point
        const rawY = lbs ? b.top + lbs[k].offsetInBlock : b.top
        if (rawY > endY - 0.01) break
        const next = lbs && k + 1 < n ? Math.min(b.top + lbs[k + 1].offsetInBlock, bottom) : bottom
        if (next < startY + 0.01) continue
        const y = Math.max(rawY, startY)
        units.push({
          y,
          h: b.emptyPara ? 0 : Math.max(next - y, 0),
          cut: rawY >= startY - 0.01 && (k === 0 || (k >= minKeep && n - k >= minKeep)),
        })
      }
    }
    const countedH = units.reduce((s, u) => s + u.h, 0)
    if (countedH <= 0.01 || contentEnd - startY <= 0.01) return null
    const target = countedH / region.cols
    const cuts: number[] = []
    let acc = 0
    let prevCut = startY
    for (const u of units) {
      if (cuts.length >= region.cols - 1) break
      if (acc >= target - 0.01 && u.cut && u.y > prevCut + 0.01) {
        cuts.push(u.y)
        prevCut = u.y
        acc = 0
      }
      acc += u.h
    }
    while (cuts.length < region.cols - 1) cuts.push(endY) // trailing empty columns
    region.entries = [region.entries[0], ...cuts.map((y) => ({ y }))]
    // column extents measured to the content end (trailing whitespace excluded)
    const cutEdges = [startY, ...cuts, endY].map((y) => Math.min(y, contentEnd))
    let maxExtent = 0
    for (let k = 1; k < cutEdges.length; k++)
      maxExtent = Math.max(maxExtent, cutEdges[k] - cutEdges[k - 1])
    region.height = maxExtent
    return maxExtent
  }

  // usable column capacity: the page's footnote separator strip is not placeable
  const capColH = () => colH - pageNoteSepPx
  // capacity of a fresh page for the CURRENT block: the separator only follows
  // blocks that carry footnotes (empty-page / one-page-height checks)
  const freshColH = () => colH - (curBlockNotes ? FOOTNOTE_SEPARATOR_H : 0)
  // whether height h fits in the current column (runaway degrade: everything fits)
  const fits = (h: number): boolean => runaway || usedInCol + h <= capColH() + 0.01
  // whether the current column is empty (just changed columns or at column top)
  const colEmpty = () => usedInCol <= 0.01
  // whether the current page is entirely blank (guards forced breaks against empty pages)
  const pageBlank = () => colIdx === 0 && regionTop <= 0.01 && usedInCol <= 0.01
  // place height h (unconditional accumulation)
  let anyContent = false
  const place = (h: number) => {
    usedInCol += h
    anyContent = true
  }

  startPage(0, initSection)

  // precompute keepNext chains (runs of consecutive keepNext blocks; the last block closes the chain)
  // chainStart[i] = chain start index (-1 when not in a chain)
  const chainStart = new Int32Array(blocks.length).fill(-1)
  for (let i = 0; i < blocks.length; i++) {
    if (blocks[i].keepNext) {
      let j = i
      while (j < blocks.length - 1 && blocks[j].keepNext) j++
      for (let k = i; k <= j; k++) chainStart[k] = i
      i = j
    }
  }

  // ── Main loop ───────────────────────────────────────────────────────────────
  for (let bi = 0; bi < blocks.length; bi++) {
    const block = blocks[bi]
    curBi = bi
    curBlockNotes = (block.footnoteExtraPx ?? 0) > 0
    if (curBlockNotes) pageNoteSepPx = FOOTNOTE_SEPARATOR_H

    // section change
    const bSection = block.section ?? curSection
    if (bSection !== curSection) {
      // sections crossed without any measured block (a lone sectPr paragraph
      // renders as a zero-height chip) still claim their own page in Word:
      // break once per crossed next-page boundary
      for (let s = curSection + 1; s < bSection; s++) {
        const gs = geomOf(s)
        if (gs.forceBreak && (block.top > pageStart || emptySectionClaimsPage(s))) {
          contentH = Math.max(gs.contentHeight, 1)
          startPage(breakY(block.top), s)
        }
      }
      const g = geomOf(bSection)
      const newCols = colsOf(bSection)
      // capacity of the page hosting the section change: a region opened on it
      // stays bounded by the host page's own footer (a titlePg first page can
      // be shorter than the new section's default capacity, prod-sas 043)
      const hostContentH = contentH
      curSection = bSection
      contentH = Math.max(g.contentHeight, 1)
      if (g.forceBreak && (block.top > pageStart || emptySectionClaimsPage(bSection))) {
        startPage(breakY(block.top), bSection)
      } else if (newCols !== colCount) {
        // continuous section changing column count: balance the closed multi-column
        // region (Word), then open a new region in the remaining page height; if the
        // page is used up, turn the page. An unbalanced region ends at its
        // tallest column's content, NOT the full column height: a short
        // letterhead row split by an explicit column break must not eat the
        // page (prod100r3/45 lost its whole first page to a 2-line region) —
        // natural overflow still yields a full first-column segment.
        const balancedH = tryBalanceRegion(bi, block.top)
        const entries = pages[pages.length - 1].regions.at(-1)?.entries ?? []
        const segMax = entries.length
          ? Math.max(
              ...entries.map(
                (e, i) => (i + 1 < entries.length ? entries[i + 1].y : block.top) - e.y,
              ),
            )
          : usedInCol
        const regionBottom = regionTop + (balancedH ?? Math.min(Math.max(segMax, 0), colH))
        if (regionBottom >= Math.min(hostContentH, contentH) - 1) {
          startPage(block.top, bSection)
        } else {
          // the new region's columns end at the host page's footer, not the new
          // section's default capacity (later pages take the full capacity via startPage)
          contentH = Math.min(hostContentH, contentH)
          regionTop = regionBottom
          openRegion(block.top, bSection)
        }
      } else {
        // column count unchanged (continuous flow continues on the same page): update column height per the new section's content height
        colH = Math.max(contentH - regionTop, 1)
        const page = pages[pages.length - 1]
        page.regions[page.regions.length - 1].height = colH
        // nextColumn into a same-count section: advance one column at the boundary (no-op at a column top)
        if (g.colBreakStart && !colEmpty()) {
          regionBroke = true
          newColumn(block.top, bSection)
        }
      }
    }

    // pageBreakBefore (highest priority: force a new page before this block; mid-column breaks also turn the page directly)
    // a pending w:br plus this block's own leading w:br are two distinct break
    // characters: both turn the page, leaving a deliberate blank sheet between
    const doubleBreak = pendingBreak && block.breakBeforeBr && !pageBlank()
    if (doubleBreak) startPage(breakY(block.top), curSection)
    // a leading w:br on the document's first content still breaks (Word keeps the
    // blank first page); breaks landing on a later blank page stay suppressed
    if (
      (pendingBreak || block.breakBefore) &&
      (doubleBreak ||
        pendingForce ||
        !pageBlank() ||
        (block.breakBeforeBr && !anyContent && pages.length === 1))
    ) {
      startPage(breakY(block.top), curSection)
    }
    pendingBreak = false
    pendingForce = false
    // column break: change column (turn the page on the last column); no-op at column top
    if (pendingColBreak && !colEmpty()) {
      regionBroke = true
      newColumn(block.top, curSection)
    }
    pendingColBreak = false

    // overflow advancement for this block: a non-reflowable block (table/textbox)
    // never advances into a column narrower than itself — Word turns the page
    // instead (explicit column breaks above are honored regardless)
    const advance =
      block.fixedWidthPx === undefined
        ? newColumn
        : (y: number, section: number, headerH = 0, headerTop = 0) => {
            const ws = geomOf(section).colWidths
            if (
              colIdx + 1 < colCount &&
              ws &&
              (ws[colIdx + 1] ?? Infinity) < block.fixedWidthPx! - 0.5
            ) {
              startPage(y, section, headerH, headerTop)
            } else {
              newColumn(y, section, headerH, headerTop)
            }
          }

    // CSS-floated block (wrapped image / w:tblpPr table): following block boxes
    // stack ignoring it (only their line boxes shorten), so it consumes no column
    // height — counting it would double-book the overlap and break pages early
    if (block.floated) {
      if (!fits(block.height) && !colEmpty()) advance(block.top, curSection)
      // page/margin-anchored w:tblpY: shift the float down to its target Y on
      // the page it lands on (never up — flow position is the floor, like the
      // X clamp keeping floats on the page)
      let floatDy = 0
      if (block.pageRelVyPx !== undefined) {
        const target =
          block.pageRelVAnchor === 'page'
            ? block.pageRelVyPx - (geomOf(curSection).topPx ?? 0)
            : block.pageRelVyPx
        floatDy = Math.max(0, target - (regionTop + usedInCol))
        out?.floatVShifts?.push({ blockTop: block.top, dyPx: floatDy })
      }
      pageFloatBottom = Math.max(pageFloatBottom, block.top + floatDy + block.height)
      place(0)
      if (block.breakAfter) {
        pendingBreak = true
        if (block.breakForce) pendingForce = true
      }
      if (block.colBreakAfter) pendingColBreak = true
      continue
    }

    // column-spanning wrap band (anchored box): Word keeps the box on its
    // anchor's page and lets it overflow the bottom margin. Fill the rest of
    // the column instead of pushing the whole band — unless not even the
    // anchor line fits, which moves the anchor (and its box) like Word.
    if (block.bandKeep && !fits(block.height)) {
      if (!fits(BAND_KEEP_MIN_H) && !colEmpty()) advance(block.top, curSection)
      const bandBottom = Math.max(block.top + block.height, block.floatBottom ?? 0)
      pageFloatBottom = Math.max(pageFloatBottom, bandBottom)
      if (!fits(block.height)) pendingBandBottom = Math.max(pendingBandBottom, bandBottom)
      place(Math.min(block.height, Math.max(capColH() - usedInCol, 0)))
      if (block.breakAfter) {
        pendingBreak = true
        if (block.breakForce) pendingForce = true
      }
      if (block.colBreakAfter) pendingColBreak = true
      continue
    }

    // break-only paragraph: Word places its break line like any other line — it
    // must fully fit below the preceding content (including that content's
    // space-after; nothing is handed back and nothing overflows the bottom
    // margin), else the whole line moves down and opens a deliberate blank page
    // (Word probe 20260901: an exact-12pt break line fits at exactly 12pt
    // remaining and blanks at 11pt; 13pt remaining minus an 8pt trailing
    // space-after still blanks). breakOnlyLineH already excludes the auto
    // line-spacing multiplier, the part Word never charges at the page bottom.
    if (block.breakOnlyLineH !== undefined) {
      if (!fits(block.breakOnlyLineH) && !pageBlank()) startPage(block.top, curSection)
      place(block.height)
      if (block.breakAfter) {
        pendingBreak = true
        if (block.breakForce) pendingForce = true
      }
      if (block.colBreakAfter) pendingColBreak = true
      continue
    }

    // ── Tables: row-level page breaking ────────────────────────────────────
    if (block.tableRows && block.tableRows.length > 0) {
      _placeTable(
        block,
        block.tableRows,
        freshColH(),
        fits,
        place,
        () => Math.max(capColH() - usedInCol, 0),
        colEmpty,
        advance,
        curSection,
        out?.rowFills
          ? (row, targetPx) => out.rowFills!.push({ blockTop: block.top, row, targetPx })
          : undefined,
      )
      if (block.spaceAfterPx) place(block.spaceAfterPx) // space after the table (may overflow into the bottom margin)
      // in-table footnote refs: the reservation (in the block height) still consumes capacity
      if (block.footnoteExtraPx) place(block.footnoteExtraPx)
      if (block.breakAfter) {
        pendingBreak = true
        if (block.breakForce) pendingForce = true
      }
      if (block.colBreakAfter) pendingColBreak = true
      continue
    }

    // single line taller than its column (oversized inline picture): Word
    // starts it at a fresh column top and overflow-clips it at the column
    // bottom instead of flowing on; the renderer bakes the clip from
    // out.oversizeClips, so the measured height converges to the capacity
    if (block.oversizeLineH !== undefined) {
      // a keepNext chain placed just above must keep its anchor: stay and clip
      // into the remainder (the chain head already pushed the pair together)
      const keptWithPrev = bi > 0 && blocks[bi - 1].keepNext === true
      if (!keptWithPrev && !colEmpty() && !fits(block.height - (block.spaceAfterPx ?? 0))) {
        advance(block.top, curSection)
      }
      out?.oversizeClips?.push({
        blockTop: block.top,
        clipPx: Math.max(capColH() - usedInCol, 1),
      })
      place(block.height)
      if (block.breakAfter) {
        pendingBreak = true
        if (block.breakForce) pendingForce = true
      }
      if (block.colBreakAfter) pendingColBreak = true
      continue
    }

    // ── Paragraph line-level placement ────────────────────────────────────
    const lineBoxes = block.lineBoxes
    const hasLines = lineBoxes && lineBoxes.length > 0
    const widowOn = block.widowControl !== false
    const spaceBeforePx = block.spaceBeforePx ?? 0
    const spaceAfterPx = block.spaceAfterPx ?? 0

    // keepNext chain (a keepNext on the document's last block has no anchor — plain placement)
    // checked before keepLines: Word heading styles carry both, and the chain decides the page push
    if (block.keepNext && chainStart[bi] === bi && bi < blocks.length - 1) {
      // chain tail: the last keepNext=true block (excluding the anchor block)
      const chainEnd = (() => {
        let j = bi
        while (j < blocks.length - 1 && blocks[j].keepNext) j++
        // j is now the first non-keepNext block (the anchor)
        // the chain tail is j-1 (the last keepNext block), while j is the anchor (next paragraph)
        // note: the while loop stops at j < length-1, so if the chain tail is at document end, j = length-1
        return j
      })()
      // chainEnd now points at the first non-keepNext block (the anchor), e.g. block[56]
      // the actual keepNext chain is bi..chainEnd-1; the anchor is chainEnd
      const lastKeepNextIdx = chainEnd - 1 // last keepNext block
      const anchorBlock = blocks[chainEnd] // anchor block (first non-keepNext)

      // a pageBreakBefore inside the chain truncates it (highest priority)
      let effectiveChainEnd = lastKeepNextIdx
      for (let k = bi + 1; k <= lastKeepNextIdx; k++) {
        if (blocks[k].breakBefore) {
          effectiveChainEnd = k - 1
          break
        }
      }
      // check whether the anchor has breakBefore (if so, the anchor is handled independently)
      const anchorHasBreak = anchorBlock?.breakBefore ?? false

      // compute the chain height (keepNext blocks) + the anchor's demand
      let chainH = 0
      for (let k = bi; k <= effectiveChainEnd; k++) chainH += blocks[k].height

      // anchor demand: the chain keeps only with the anchor's first line (first 2
      // with widow control on — Word's orphan minimum); the anchor is not part of
      // the atomic unit and flows normally after the chain. Exceptions: keepLines
      // anchors follow whole; table anchors count their first row (Word keeps the
      // heading with the table head, not the whole table)
      let anchorNeedH = 0
      if (!anchorHasBreak && anchorBlock) {
        const aLines = anchorBlock.lineBoxes
        if (anchorBlock.keepLines) anchorNeedH = anchorBlock.height
        else if (anchorBlock.tableRows?.length) anchorNeedH = anchorBlock.tableRows[0].height
        else if (anchorBlock.oversizeLineH !== undefined) {
          // an oversized anchor clips to whatever remains below the chain, so it
          // demands the rest of a fresh column: the chain pushes and the pair
          // shares the page (Word keeps the heading above the clipped picture).
          // The 0.5 keeps chainH + need inside the raw <= freshColH guard below.
          anchorNeedH = Math.max(freshColH() - chainH - 0.5, 0)
        } else if (aLines?.length) {
          const need = anchorBlock.widowControl !== false ? Math.min(2, aLines.length) : 1
          anchorNeedH = anchorBlock.spaceBeforePx ?? 0
          for (let li = 0; li < need; li++) anchorNeedH += aLines[li].height
        } else {
          anchorNeedH = anchorBlock.height // no line data yet (first pass): conservative
        }
      }
      const chainPlusAnchorH = chainH + anchorNeedH

      if (chainH <= freshColH()) {
        // whole chain (keepNext blocks) fits on a page: the chain + anchor demand
        // must share a page (keepNext semantics); if it doesn't fit, push the whole chain
        // to the next page (Word behavior; corpus 04 evidence: section 3.2 chain pushed).
        // Only abandon the constraint when chain + anchor demand can't fit even an
        // empty page (no solution; avoids infinite loops).
        if (!fits(chainPlusAnchorH) && !colEmpty() && chainPlusAnchorH <= freshColH()) {
          advance(block.top, curSection)
        }
        // place chain head through chain tail (the keepNext blocks)
        for (let k = bi; k <= effectiveChainEnd; k++) place(blocks[k].height)
        bi = effectiveChainEnd
        if (blocks[effectiveChainEnd].breakAfter) {
          pendingBreak = true
          if (blocks[effectiveChainEnd].breakForce) pendingForce = true
        }
        if (blocks[effectiveChainEnd].colBreakAfter) pendingColBreak = true
        continue
      }

      // chain exceeds one page: only guarantee the chain head + anchor demand share a page (minimum guarantee)
      const headH = block.height + anchorNeedH
      if (!fits(headH) && !colEmpty()) {
        advance(block.top, curSection)
      }
      if (!block.keepLines) {
        // place the chain head block
        _placeParaBlock(
          block,
          hasLines ? lineBoxes! : null,
          widowOn,
          spaceBeforePx,
          spaceAfterPx,
          freshColH(),
          fits,
          place,
          colEmpty,
          advance,
          curSection,
          colCount > 1,
        )
        if (block.breakAfter) {
          pendingBreak = true
          if (block.breakForce) pendingForce = true
        }
        if (block.colBreakAfter) pendingColBreak = true
        continue
      }
      // keepLines head falls through to the keepLines branch (must not split)
    }

    // keepLines: the whole paragraph must stay on one page (one column in multi-column layout)
    if (block.keepLines) {
      if (!fits(block.height) && block.height <= freshColH() && !colEmpty()) {
        advance(block.top, curSection)
      }
      if (!fits(block.height)) {
        // paragraph exceeds one page: hard line-level cut (best effort)
        if (hasLines) {
          _hardCutLines(
            block,
            lineBoxes!,
            spaceBeforePx,
            spaceAfterPx,
            fits,
            place,
            colEmpty,
            advance,
            curSection,
          )
        } else {
          // no line data: hard pixel cuts consuming the remainder column by
          // column. (Retesting the full block height after each turn never
          // fits a block taller than one column and used to loop forever.)
          let offset = 0
          while (block.height - offset > capColH() - usedInCol + 0.01 && !runaway) {
            offset += Math.max(capColH() - usedInCol, 1)
            advance(block.top + Math.min(offset, block.height), curSection)
          }
          place(block.height - offset)
        }
      } else {
        place(block.height)
      }
      if (block.breakAfter) {
        pendingBreak = true
        if (block.breakForce) pendingForce = true
      }
      if (block.colBreakAfter) pendingColBreak = true
      continue
    }

    // ordinary block (incl. mid/tail keepNext chain blocks; chain constraints were handled by the chain head)
    _placeParaBlock(
      block,
      hasLines ? lineBoxes! : null,
      widowOn,
      spaceBeforePx,
      spaceAfterPx,
      freshColH(),
      fits,
      place,
      colEmpty,
      advance,
      curSection,
      colCount > 1,
    )
    if (block.breakAfter) {
      pendingBreak = true
      if (block.breakForce) pendingForce = true
    }
    if (block.colBreakAfter) pendingColBreak = true
  }

  // a trailing page break keeps its deliberate blank last page (Word: the final
  // paragraph mark lands after the break; LO dropping it is tdf#99090); a trailing
  // column break advances the same way (new page when the last column is used)
  if (pendingBreak) startPage(Math.max(total, pageStart), curSection)
  else if (pendingColBreak && !colEmpty()) newColumn(Math.max(total, pageStart), curSection)

  // ── Output: flatten column starts into ranges, aggregate by page (pages with cols>1 regions get regions attached) ──
  const flat: ColEntry[] = []
  for (const p of pages) for (const r of p.regions) for (const e of r.entries) flat.push(e)
  const flowEnd = Math.max(total, pageStart)
  const endOf = new Map<ColEntry, number>()
  flat.forEach((e, i) => endOf.set(e, i + 1 < flat.length ? flat[i + 1].y : flowEnd))

  return pages.map((p) => {
    const entries = p.regions.flatMap((r) => r.entries)
    const first = entries[0]
    const multiCol = p.regions.length > 1 || p.regions.some((r) => r.cols > 1)
    const last = p.regions[p.regions.length - 1]
    const lastExtent = Math.min(
      last.height,
      Math.max(...last.entries.map((e) => endOf.get(e)! - e.y + (e.repeatHeader?.height ?? 0)), 0),
    )
    return {
      start: first.y,
      end: endOf.get(entries[entries.length - 1])!,
      section: p.section,
      ...(first.repeatHeader ? { repeatHeader: first.repeatHeader } : {}),
      ...(multiCol
        ? {
            regions: p.regions.map((r) => ({
              top: r.top,
              height: r.height,
              section: r.section,
              columns: r.entries.map((e) => ({
                start: e.y,
                end: endOf.get(e)!,
                ...(e.repeatHeader ? { repeatHeader: e.repeatHeader } : {}),
              })),
            })),
            physHeight: last.top + lastExtent,
          }
        : {}),
    }
  })
}

/**
 * Place a table (row-level page breaking).
 */
function _placeTable(
  block: BlockBox,
  rows: TableRowBox[],
  contentH: number,
  fits: (h: number) => boolean,
  place: (h: number) => void,
  remain: () => number,
  pageEmpty: () => boolean,
  newPage: (y: number, section: number, headerH?: number, headerTop?: number) => void,
  curSection: number,
  onRowFill?: (row: number, targetPx: number) => void,
) {
  // find header rows (the first N consecutive isHeader rows); a header block
  // taller than a full page doesn't repeat (Word probe 2026-08-16: a block at
  // 94% of the page still repeats on every page, so the gate is the full
  // content height, not half of it)
  let headerHeight = 0
  let leadHeaderRows = 0
  for (const r of rows) {
    if (!r.isHeader) break
    headerHeight += r.height
    leadHeaderRows++
  }
  const headerBlockH = headerHeight
  // headerRows drives per-page repetition only; push-whole protection keeps
  // using leadHeaderRows — Word never splits a tblHeader row even when the
  // header block is too tall to repeat
  let headerRows = leadHeaderRows
  if (headerHeight > contentH) {
    headerHeight = 0
    headerRows = 0
  }

  // Word 2013+ layout (compatibilityMode >= 15): a multirow tblHeader block
  // that doesn't fit the remaining space starts the table on a fresh page —
  // even when the block exceeds a full page (probe 2026-08-16, tdf88496 F30/F31;
  // legacy mode instead splits the header block in place)
  if (block.modernTableHeaders && leadHeaderRows > 0 && !pageEmpty() && headerBlockH > remain()) {
    newPage(block.top, curSection)
  }

  let rowCursor = block.top
  let placedHeader = false

  for (let ri = 0; ri < rows.length; ri++) {
    const row = rows[ri]

    if (row.vMergeContinue) {
      rowCursor += row.height
      continue
    }

    // table broken onto a new page: repeat headers only if they already appeared on a prior page (reserve header space at page top)
    const repeatH = placedHeader && ri >= headerRows ? headerHeight : 0

    if (!fits(row.height)) {
      const contentEnd =
        row.contentBottom !== undefined ? Math.min(row.contentBottom, row.height) : row.height
      // in-row page break (Word default): without cantSplit and with safe cut points,
      // place segment by segment at the cut points. If the first segment doesn't fit,
      // turn the page first (equivalent to pushing the whole row)
      // Word probe (2026-08-13): a plain first row has no special rule — it splits
      // like any row. Only tblHeader/cantSplit rows push whole. Over-page header/
      // cantSplit rows Word overflow-clips instead; we split them (DOM clipping is
      // costly) — a deliberate deviation.
      // Declared atLeast height is reserved space, not breakable content: when
      // the declared minimum overflows the page remainder, Word starts the row
      // on a fresh page instead of splitting into the remainder (near-empty
      // TOC pages, prod100r3/50; Word probe 2026-08-23). Rows taller than a
      // full page still split afterwards — from the fresh page.
      let turnedForMinH = false
      if (row.minHPx !== undefined && row.minHPx > remain() + 0.01 && !pageEmpty()) {
        newPage(rowCursor, curSection, repeatH, block.top)
        turnedForMinH = true
      }
      const keepWhole = row.isHeader && ri < leadHeaderRows && row.height <= contentH + 0.01
      let cuts = !row.cantSplit && !keepWhole && row.cutYs ? [...row.cutYs] : []
      let segTurns = 0
      let lastTurnPrev = 0
      const placeSegments = (bounds: number[]) => {
        let prev = 0
        for (const cut of bounds) {
          const seg = cut - prev
          if (seg <= 0.5) continue
          if (!fits(seg) && !pageEmpty()) {
            newPage(rowCursor + prev, curSection, repeatH, block.top)
            segTurns++
            lastTurnPrev = prev
          }
          place(seg)
          prev = cut
        }
        return prev
      }
      // Word probe 2026-08-27 (kr_fill_repro): when a declared-height (atLeast
      // trHeight) row splits across pages, the continuation fragment honors the
      // full declared height again as its own minimum — Word lays the row out
      // afresh on the new page with the same trPr. Stretch the bookkeeping and
      // report the row's target total so the DOM tr can be patched to match.
      // The target derives from the content bands (decoration-independent) and
      // is re-emitted on every measure: a previously patched row keeps its
      // decoration instead of oscillating with an edge-triggered report.
      const continuationFill = () => {
        if (segTurns === 0 || row.minHPx === undefined || !onRowFill) return
        const want = Math.min(row.minHPx, contentH)
        const contentB = Math.min(row.contentBottom ?? row.height, row.height)
        const target = lastTurnPrev + Math.max(contentB - lastTurnPrev, want)
        const extra = target - row.height
        if (extra > 0.5) place(Math.min(extra, remain()))
        // report the content-derived target as-is: it never undercuts the row's
        // own declared-height style (target >= the declared minimum by
        // construction), and a stale taller patch shrinks back after edits
        onRowFill(ri, target)
      }
      // Only rows taller than a page get their declared-height fill clipped to the
      // page remainder (Word truncates over-tall rows within the page). Page-sized
      // rows keep the fill glued to the last segment: a clipped fill would desync
      // bookkeeping from DOM height, leaking shading/empty rows past the page edge.
      if (row.height > contentH + 0.01) {
        if (!row.cantSplit && contentEnd > contentH) {
          // A fixed-height row can be taller than a page while containing only one
          // text band, so DOM line sampling may provide too few natural cuts. Keep
          // every segment page-sized; natural inter-band cuts remain preferred and
          // a hard content-band cut is only inserted where no legal cut advances.
          const bounded: number[] = []
          let previous = 0
          for (const candidate of [...cuts, contentEnd]) {
            while (candidate - previous > contentH + 0.01) {
              previous += contentH
              bounded.push(previous)
            }
            if (candidate < row.height - 0.5 && candidate > previous + 0.5) {
              bounded.push(candidate)
              previous = candidate
            }
          }
          cuts = bounded
        }
        cuts = cuts.filter((c) => c < contentEnd - 0.01)
        if (cuts.length > 0 || contentEnd < row.height - 0.5) {
          const prev = placeSegments([...cuts, contentEnd])
          const fill = row.height - prev
          if (fill > 0.5) place(Math.min(fill, remain()))
          continuationFill()
          rowCursor += row.height
          if (row.isHeader && ri < headerRows) placedHeader = true
          continue
        }
      } else if (cuts.length > 0) {
        placeSegments([...cuts, row.height])
        continuationFill()
        rowCursor += row.height
        if (row.isHeader && ri < headerRows) placedHeader = true
        continue
      }
      // cantSplit / empty / no cut points: the row is atomic; turn the page
      // first if it doesn't fit. A repeated header keeps the fresh page
      // non-empty, so the minH turn above must not double up here.
      if (!pageEmpty() && !turnedForMinH) newPage(rowCursor, curSection, repeatH, block.top)
    }
    place(row.height)
    rowCursor += row.height
    if (row.isHeader && ri < headerRows) placedHeader = true
  }
}

/**
 * Per-line footnote reservations: bandH[li] = note heights charged with line li
 * (the page holding a reference line hosts its note, like Word). Bands whose
 * markers were not resolved ride the last line. Null when the block has none.
 */
function noteBandHeights(
  block: BlockBox,
  lineBoxes: Array<{ offsetInBlock: number; height: number }>,
): Float64Array | null {
  const fnExtra = block.footnoteExtraPx ?? 0
  if (fnExtra <= 0) return null
  const bands =
    block.noteBands && block.noteBands.length > 0
      ? block.noteBands
      : [{ offset: Infinity, height: fnExtra }]
  const bandH = new Float64Array(lineBoxes.length)
  for (const band of bands) {
    let li = lineBoxes.length - 1
    for (let k = 1; k < lineBoxes.length; k++) {
      if (band.offset < lineBoxes[k].offsetInBlock) {
        li = k - 1
        break
      }
    }
    bandH[li] += band.height
  }
  return bandH
}

/**
 * Place a paragraph block (with widowControl).
 * With lineBoxes = null, degrades to F1 block-level placement.
 */
function _placeParaBlock(
  block: BlockBox,
  lineBoxes: Array<{ offsetInBlock: number; height: number }> | null,
  widowOn: boolean,
  spaceBeforePx: number,
  spaceAfterPx: number,
  contentH: number,
  fits: (h: number) => boolean,
  place: (h: number) => void,
  pageEmpty: () => boolean,
  newPage: (y: number, section: number) => void,
  curSection: number,
  multiCol = false,
) {
  const totalH = block.height

  // whole paragraph (text + note reservation) fits: place directly. Trailing
  // space doesn't consume capacity (Word breaks by text only; it may overflow
  // into the bottom margin) — the note reservation is in the height, not here
  if (fits(totalH - spaceAfterPx)) {
    place(totalH)
    return
  }

  // whole paragraph doesn't fit
  if (!lineBoxes || lineBoxes.length === 0) {
    // F1 block-level placement: push to the next page if it doesn't fit (when <= one page), otherwise F1-style hard cut
    if (totalH <= contentH) {
      if (!pageEmpty()) newPage(block.top, curSection)
    } else if (multiCol && !pageEmpty()) {
      // over-column block in a column region: single-column pages place it
      // directly and the NEXT block turns the page, but short columns can see
      // a run of such blocks — placed directly they'd all stack in one column
      // (never advancing) and overprint the footer. Start each from a fresh
      // column so the pile-up is bounded to one block per column.
      newPage(block.top, curSection)
    }
    // over one page with no line data: place directly and let the next block
    // trigger the page turn (F1 handles a big block the same way)
    place(totalH)
    return
  }

  // line-level placement
  const nLines = lineBoxes.length
  const bandH = noteBandHeights(block, lineBoxes)

  if (totalH > contentH) {
    // paragraph exceeds one page: hard line-level cut
    _hardCutLines(
      block,
      lineBoxes,
      spaceBeforePx,
      spaceAfterPx,
      fits,
      place,
      pageEmpty,
      newPage,
      curSection,
    )
    return
  }

  // paragraph <= one page but doesn't fit on the current page: widow/orphan decision
  // count how many lines fit on the current page (a line with a footnote ref
  // demands its note's page-bottom reservation along with the line itself)
  let splitLine = -1 // line break point (-1 = push the whole paragraph)

  if (widowOn && nLines >= 2) {
    let sumH = spaceBeforePx
    for (let li = 0; li < nLines; li++) {
      sumH += lineBoxes[li].height + (bandH?.[li] ?? 0)
      if (!fits(sumH)) {
        splitLine = li // line li doesn't fit
        break
      }
    }
    if (splitLine === -1) {
      // only the trailing space overflows: the paragraph stays on this page and
      // the space overflows into the bottom margin (Word paginates by text only;
      // corpus 14 PDF measurement: end-of-page text stops at 758.9pt < bottom
      // bound 769.9, and the overflowing space-after doesn't push the paragraph)
      place(totalH)
      return
    }

    // widow/orphan adjustment: at least 2 lines at page bottom, at least 2 at page top
    // tailLines = lines on the current page, headLines = lines on the next page
    const tailLines = splitLine // splitLine lines stay on the current page (0..splitLine-1)
    // headLines = nLines - splitLine

    if (tailLines === 0) {
      // not even one line fits: push the whole paragraph
      splitLine = -1
    } else if (tailLines === 1) {
      // orphan at page bottom: drop one line (push line0 to the next page too)
      if (splitLine - 1 <= 0) {
        // nothing left after dropping: push the whole paragraph
        splitLine = -1
      } else {
        splitLine -= 1 // try tailLines = splitLine - 1
      }
    }

    if (splitLine > 0) {
      const newHead = nLines - splitLine
      if (newHead === 1) {
        // widow at page top: give up one line here so the next page gets 2 lines (Word)
        splitLine -= 1
        // fewer than 2 lines left at page bottom would be an orphan: push the whole paragraph
        if (splitLine < 2) splitLine = -1
      }
    }
  } else if (!widowOn) {
    // widowControl off: find the first line that doesn't fit
    let sumH = spaceBeforePx
    for (let li = 0; li < nLines; li++) {
      sumH += lineBoxes[li].height + (bandH?.[li] ?? 0)
      if (!fits(sumH)) {
        splitLine = li
        break
      }
    }
    if (splitLine === -1) {
      // only the trailing space overflows (see the widow-control twin above)
      place(totalH)
      return
    }
  } else {
    // only 1 line: push the whole paragraph
    splitLine = -1
  }

  if (splitLine <= 0) {
    // push the whole paragraph to the next page
    if (!pageEmpty()) newPage(block.top, curSection)
    place(totalH)
  } else {
    // break the page before line splitLine (each line carries its notes' reservation)
    if (spaceBeforePx > 0) place(spaceBeforePx)
    for (let li = 0; li < splitLine; li++) place(lineBoxes[li].height + (bandH?.[li] ?? 0))
    // page break (line offsets are element-relative: they start after the space-before)
    newPage(block.top + spaceBeforePx + lineBoxes[splitLine].offsetInBlock, curSection)
    // place remaining lines on the new page
    for (let li = splitLine; li < nLines; li++) place(lineBoxes[li].height + (bandH?.[li] ?? 0))
    if (spaceAfterPx > 0) place(spaceAfterPx)
  }
}

/**
 * Hard-cut lines (best effort when the paragraph exceeds one page).
 */
function _hardCutLines(
  block: BlockBox,
  lineBoxes: Array<{ offsetInBlock: number; height: number }>,
  spaceBeforePx: number,
  spaceAfterPx: number,
  fits: (h: number) => boolean,
  place: (h: number) => void,
  pageEmpty: () => boolean,
  newPage: (y: number, section: number) => void,
  curSection: number,
) {
  const bandH = noteBandHeights(block, lineBoxes)
  if (spaceBeforePx > 0) {
    if (!fits(spaceBeforePx) && !pageEmpty()) {
      newPage(block.top, curSection)
    }
    place(spaceBeforePx)
  }
  for (let li = 0; li < lineBoxes.length; li++) {
    const lb = lineBoxes[li]
    const need = lb.height + (bandH?.[li] ?? 0)
    if (!fits(need) && !pageEmpty()) {
      newPage(block.top + spaceBeforePx + lb.offsetInBlock, curSection)
    }
    place(need)
  }
  if (spaceAfterPx > 0) {
    if (!fits(spaceAfterPx) && !pageEmpty()) {
      // geometric end of the block's text (height also carries the footnote reservation)
      newPage(block.top + block.height - spaceAfterPx - (block.footnoteExtraPx ?? 0), curSection)
    }
    place(spaceAfterPx)
  }
}

/**
 * evenPage/oddPage sections: insert a zero-height blank page slice when the section's
 * first page has the wrong physical parity. Parity is approximated by physical
 * page order (1-based) — exact when page numbers run from 1.
 */
export function insertParityBlanks(slices: PageSlice[], geoms: SectionGeom[]): PageSlice[] {
  if (!geoms.some((g) => g.startType === 'evenPage' || g.startType === 'oddPage')) return slices
  const out: PageSlice[] = []
  for (const s of slices) {
    const prev = out[out.length - 1]
    if (prev && s.section !== prev.section) {
      const st = geoms[Math.max(0, Math.min(s.section, geoms.length - 1))]?.startType
      const ordinal = out.length + 1
      if ((st === 'evenPage' && ordinal % 2 === 1) || (st === 'oddPage' && ordinal % 2 === 0)) {
        out.push({ start: s.start, end: s.start, section: prev.section })
      }
    }
    out.push(s)
  }
  return out
}

/** Whether each page is the first page of its section (for section-level titlePg) */
export function sectionFirstPages(slices: PageSlice[]): boolean[] {
  let prev = -1
  return slices.map((s) => {
    const first = s.section !== prev
    prev = s.section
    return first
  })
}

/** Page containing content-area Y (1-based) */
export function pageAt(slices: PageSlice[], y: number): number {
  let page = 1
  for (let i = 1; i < slices.length; i++) {
    if (y >= slices[i].start) page = i + 1
  }
  return page
}

/**
 * Pages the user can see: an even/odd-section parity blank shares its start with the
 * neighbouring slice and draws no page, so NUMPAGES, the status bar, and the gap
 * header/footer widgets all count only distinct slice starts (up to `upTo` slices).
 */
export function visiblePageCount(slices: PageSlice[], upTo = slices.length): number {
  let n = 0
  for (let i = 0; i < Math.min(upTo, slices.length); i++) {
    // a zero-height predecessor is a deliberate blank page (leading/double w:br,
    // even/odd parity): the page after it is still its own visible page
    if (
      i === 0 ||
      slices[i].start !== slices[i - 1].start ||
      slices[i - 1].end === slices[i - 1].start
    )
      n++
  }
  return n
}

/**
 * Index of each non-first-page page-leading block (a page gap should be inserted before it).
 * Hard pixel-cut boundaries (inside over-page big blocks, with no matching block) are skipped.
 */
export function pageStartBlocks(blocks: BlockBox[], slices: PageSlice[]): number[] {
  const starts: number[] = []
  for (const slice of slices.slice(1)) {
    const i = blocks.findIndex((b) => Math.abs(b.top - slice.start) < 0.5)
    if (i >= 0) starts.push(i)
  }
  return starts
}
