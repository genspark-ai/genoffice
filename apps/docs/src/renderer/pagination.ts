/**
 * Pagination slicing: greedy page breaking over the continuous-flow render result, by top-level block.
 * Pure functions; all coordinates are content-area Y at 100% zoom (px, 0 = top of page 1 content).
 */
import type { HeaderFooter, HfPartInfo, SectionInfo, SectionSettings } from '@genoffice/docx-engine'

export interface BlockBox {
  top: number
  height: number
  /** paragraph pageBreakBefore: force a page break before the block */
  breakBefore?: boolean
  /** block contains a page-break field (w:br type=page): force a page break after it */
  breakAfter?: boolean
  /** block contains a column break (w:br type=column): force a column change after it (new page on last column) */
  colBreakAfter?: boolean
  /** source DOM block (filled during canvas measurement, used to position page-gap decorations) */
  el?: HTMLElement
  /** the block's docxIndex (DOM data-idx; new unsaved blocks lack one) */
  docxIndex?: number
  /** owning section index (filled by assignSections) */
  section?: number
  /** in-block line boundaries (relative to block top, ascending, each = a line's starting Y): used to split page-crossing blocks by line */
  lineOffsets?: number[]
  /** min lines kept on each side of a split (widow/orphan control): paragraphs 2, table rows 1 (default) */
  splitMinLines?: number

  // ── F2 line-level page-split extensions ─────────────────────────────────
  /**
   * Paragraph line-box list (from computeLineMetrics, for line-level page splitting).
   * When absent, degrades to F1 block-level greedy placement.
   */
  lineBoxes?: Array<{ offsetInBlock: number; height: number }>
  /** space before (px), from line-metrics output */
  spaceBeforePx?: number
  /** space after (px), from line-metrics output */
  spaceAfterPx?: number

  // ── F2 pagination constraints ───────────────────────────────────────────
  /** keepLines: all lines of the paragraph must be on the same page */
  keepLines?: boolean
  /** keepNext: the paragraph and the next paragraph's first line must be on the same page */
  keepNext?: boolean
  /** widowControl: false = widow/orphan protection off (Word default on) */
  widowControl?: boolean

  // ── F2 table row-level page-split extensions ─────────────────────────────
  /**
   * Table row data (from parseDocx).
   * When present, table rows become the page-split unit (instead of hard pixel cuts).
   */
  tableRows?: TableRowBox[]

  /** virtual endnotes-area block (appended by appendEndnotesBlock; no DOM/docxIndex) */
  isEndnotes?: boolean
}

/**
 * Table row box (for F2 table row-level page splitting).
 */
export interface TableRowBox {
  /** row height (px) */
  height: number
  /** cantSplit: the row cannot be broken internally (whole row stays on one page) */
  cantSplit?: boolean
  /** tblHeader: the row is a header row, repeated at the top of the next page after a break */
  isHeader?: boolean
  /** vertical merge (vMerge continue): the row continues a merged row; its height is not counted independently */
  vMergeContinue?: boolean
  /** in-row safe cut points (relative to row top, px, ascending): spanning all cells without splitting any text line/image.
   *  Word allows in-row page breaks by default; without cut points or with cantSplit the row is atomic */
  cutYs?: number[]
  /** bottom of the lowest content band (text/image rects) relative to row top (px, 0 = empty row):
   *  lets pagination clip declared-height fill below the content instead of pushing pages */
  contentBottom?: number
}

/** One column of a multi-column page: a content range in the continuous flow (in-column break semantics match pages) */
export interface PageColumn {
  start: number
  end: number
  /** table continued into the column: header-row range repeated at column top (virtual coordinates) */
  repeatHeader?: { top: number; height: number }
}

/** A column-flow region within a page (a continuous column-count change can stack multiple regions vertically on one page) */
export interface PageRegion {
  /** region top relative to the page content-area top (px) */
  top: number
  /** available height per column within the region (px) */
  height: number
  /** owning section index of the region (for column count/width) */
  section: number
  /** content ranges per column (ascending; single-column regions have length 1) */
  columns: PageColumn[]
}

/** Content range [start, end) shown on one page; height ≤ the owning section's page content height */
export interface PageSlice {
  start: number
  end: number
  /** owning section index of this page */
  section: number
  /**
   * tblHeader repetition: this page starts mid-table, so the source table's header
   * rows must render first. top/height is the header rows' range in the continuous
   * flow (virtual coordinates); the preview clones and crops accordingly.
   */
  repeatHeader?: { top: number; height: number }
  /**
   * Column flow: provided when this page has cols>1 regions (omitted for single-column
   * pages; consumers use the original path). start/end is still the whole-page flow
   * range (= first column start .. last column end); the span can reach columns × column height.
   */
  regions?: PageRegion[]
}

/** Pagination geometry for one section */
export interface SectionGeom {
  contentHeight: number
  /** section start forces a page break (nextPage/evenPage/oddPage, or continuous with different page geometry) */
  forceBreak: boolean
  /** section break type: evenPage/oddPage need physical blank pages inserted to align parity */
  startType?: SectionInfo['startType']
  /** equal-width column count (w:cols w:num, default 1): page capacity = columns × column height */
  cols?: number
}

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
  const initSection = blocks[0]?.section ?? 0
  if (geoms.length === 0 || geoms.every((g) => g.contentHeight <= 0)) {
    return [{ start: 0, end: total, section: initSection }]
  }
  const geomOf = (s: number) => geoms[Math.max(0, Math.min(s, geoms.length - 1))]

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
      if (geomOf(bSection).forceBreak && block.top > pageStart) newPage(block.top, bSection)
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

  const end = Math.max(total, pageStart)
  return starts.map((s, i) => ({
    start: s.y,
    end: i + 1 < starts.length ? starts[i + 1].y : end,
    section: s.section,
  }))
}

const twipsToPx = (twips: number) => (twips / 1440) * 96

/** Physical section page box shared by editor and pagination preview. */
export function sectionPageBox(set: SectionSettings): {
  width: number
  height: number
  contentWidth: number
  headerDist: number
  footerDist: number
} {
  return {
    width: twipsToPx(set.pageWidth),
    height: twipsToPx(set.pageHeight),
    contentWidth: twipsToPx(set.pageWidth - set.marginLeft - set.marginRight),
    headerDist: twipsToPx(set.headerDist ?? 720),
    footerDist: twipsToPx(set.footerDist ?? 720),
  }
}

// ── F2 line-level page splitting + Word pagination constraints ───────────────

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
): PageSlice[] {
  const total = Math.max(totalHeight, 0)
  const initSection = blocks[0]?.section ?? 0
  if (geoms.length === 0 || geoms.every((g) => g.contentHeight <= 0)) {
    return [{ start: 0, end: total, section: initSection }]
  }
  const geomOf = (s: number) => geoms[Math.max(0, Math.min(s, geoms.length - 1))]
  const colsOf = (s: number) => Math.max(1, geomOf(s).cols ?? 1)

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
  let pendingColBreak = false

  // Safety net: a document legitimately needs at most a few column turns per
  // block (forced breaks, line/row splits) — far beyond that means a placement
  // loop stopped converging. Degrade by treating everything as fitting (the
  // rest piles onto the current page, with a warning) instead of looping forever.
  const maxColumnTurns = Math.max(65536, blocks.length * 8)
  let columnTurns = 0
  let runaway = false

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
    pageStart = y
    regionTop = 0
    pages.push({ section, regions: [] })
    openRegion(y, section, headerH, headerTop)
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

  // whether height h fits in the current column (runaway degrade: everything fits)
  const fits = (h: number): boolean => runaway || usedInCol + h <= colH + 0.01
  // whether the current column is empty (just changed columns or at column top)
  const colEmpty = () => usedInCol <= 0.01
  // whether the current page is entirely blank (guards forced breaks against empty pages)
  const pageBlank = () => colIdx === 0 && regionTop <= 0.01 && usedInCol <= 0.01
  // place height h (unconditional accumulation)
  const place = (h: number) => {
    usedInCol += h
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

    // section change
    const bSection = block.section ?? curSection
    if (bSection !== curSection) {
      const g = geomOf(bSection)
      const newCols = colsOf(bSection)
      curSection = bSection
      contentH = Math.max(g.contentHeight, 1)
      if (g.forceBreak && block.top > pageStart) {
        startPage(block.top, bSection)
      } else if (newCols !== colCount) {
        // continuous section changing column count: open a new region in the remaining page height; if the page is used up, turn the page
        const regionBottom = regionTop + (colIdx > 0 ? colH : Math.min(usedInCol, colH))
        if (regionBottom >= contentH - 1) {
          startPage(block.top, bSection)
        } else {
          regionTop = regionBottom
          openRegion(block.top, bSection)
        }
      } else {
        // column count unchanged (continuous flow continues on the same page): update column height per the new section's content height
        colH = Math.max(contentH - regionTop, 1)
        const page = pages[pages.length - 1]
        page.regions[page.regions.length - 1].height = colH
      }
    }

    // pageBreakBefore (highest priority: force a new page before this block; mid-column breaks also turn the page directly)
    if ((pendingBreak || block.breakBefore) && !pageBlank()) {
      startPage(block.top, curSection)
    }
    pendingBreak = false
    // column break: change column (turn the page on the last column); no-op at column top
    if (pendingColBreak && !colEmpty()) {
      newColumn(block.top, curSection)
    }
    pendingColBreak = false

    // ── Tables: row-level page breaking ────────────────────────────────────
    if (block.tableRows && block.tableRows.length > 0) {
      _placeTable(
        block,
        block.tableRows,
        colH,
        fits,
        place,
        () => Math.max(colH - usedInCol, 0),
        colEmpty,
        newColumn,
        curSection,
      )
      if (block.spaceAfterPx) place(block.spaceAfterPx) // space after the table (may overflow into the bottom margin)
      if (block.breakAfter) pendingBreak = true
      if (block.colBreakAfter) pendingColBreak = true
      continue
    }

    // ── Paragraph line-level placement ────────────────────────────────────
    const lineBoxes = block.lineBoxes
    const hasLines = lineBoxes && lineBoxes.length > 0
    const widowOn = block.widowControl !== false
    const spaceBeforePx = block.spaceBeforePx ?? 0
    const spaceAfterPx = block.spaceAfterPx ?? 0

    // keepLines: the whole paragraph must stay on one page (one column in multi-column layout)
    if (block.keepLines) {
      if (!fits(block.height) && block.height <= colH && !colEmpty()) {
        newColumn(block.top, curSection)
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
            newColumn,
            curSection,
          )
        } else {
          // no line data: hard pixel cuts consuming the remainder column by
          // column. (Retesting the full block height after each turn never
          // fits a block taller than one column and used to loop forever.)
          let offset = 0
          while (block.height - offset > colH - usedInCol + 0.01 && !runaway) {
            offset += Math.max(colH - usedInCol, 1)
            newColumn(block.top + Math.min(offset, block.height), curSection)
          }
          place(block.height - offset)
        }
      } else {
        place(block.height)
      }
      if (block.breakAfter) pendingBreak = true
      if (block.colBreakAfter) pendingColBreak = true
      continue
    }

    // keepNext chain
    if (block.keepNext && chainStart[bi] === bi) {
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

      // compute the chain height (keepNext blocks) + the anchor's first-line height
      let chainH = 0
      for (let k = bi; k <= effectiveChainEnd; k++) chainH += blocks[k].height

      // anchor first-line height (only relevant when the anchor lacks breakBefore);
      // anchored tables count their first row only (Word keeps the heading with the table head, not the whole table)
      const anchorFirstLineH =
        !anchorHasBreak && anchorBlock
          ? (anchorBlock.lineBoxes?.[0]?.height ??
            anchorBlock.tableRows?.[0]?.height ??
            anchorBlock.height)
          : 0
      const chainPlusAnchorH = chainH + anchorFirstLineH

      if (chainH <= colH) {
        // whole chain (keepNext blocks) fits on a page: the chain + anchor's first line
        // must share a page (keepNext semantics); if it doesn't fit, push the whole chain
        // to the next page (Word behavior; corpus 04 evidence: section 3.2 chain pushed).
        // Only abandon the constraint when chain + anchor first line can't fit even an
        // empty page (no solution; avoids infinite loops).
        if (!fits(chainPlusAnchorH) && !colEmpty() && chainPlusAnchorH <= colH) {
          newColumn(block.top, curSection)
        }
        // place chain head through chain tail (the keepNext blocks)
        for (let k = bi; k <= effectiveChainEnd; k++) place(blocks[k].height)
        bi = effectiveChainEnd
        if (blocks[effectiveChainEnd].breakAfter) pendingBreak = true
        if (blocks[effectiveChainEnd].colBreakAfter) pendingColBreak = true
        continue
      }

      // chain exceeds one page: only guarantee the chain head + anchor first line share a page (minimum guarantee)
      const headH = block.height + anchorFirstLineH
      if (!fits(headH) && !colEmpty()) {
        newColumn(block.top, curSection)
      }
      // place the chain head block
      _placeParaBlock(
        block,
        hasLines ? lineBoxes! : null,
        widowOn,
        spaceBeforePx,
        spaceAfterPx,
        colH,
        fits,
        place,
        colEmpty,
        newColumn,
        curSection,
      )
      if (block.breakAfter) pendingBreak = true
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
      colH,
      fits,
      place,
      colEmpty,
      newColumn,
      curSection,
    )
    if (block.breakAfter) pendingBreak = true
    if (block.colBreakAfter) pendingColBreak = true
  }

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
          }
        : {}),
    }
  })
}

// ── Internal helpers ────────────────────────────────────────────────────────

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
) {
  // find header rows (the first N consecutive isHeader rows); headers filling more than half a page don't repeat (Word behavior)
  let headerHeight = 0
  let headerRows = 0
  for (const r of rows) {
    if (!r.isHeader) break
    headerHeight += r.height
    headerRows++
  }
  if (headerHeight > contentH / 2) {
    headerHeight = 0
    headerRows = 0
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
      // Word never splits the first/header row: push whole (over-page rows still split)
      const keepWhole =
        (ri === 0 || (row.isHeader && ri < headerRows)) && row.height <= contentH + 0.01
      let cuts = !row.cantSplit && !keepWhole && row.cutYs ? [...row.cutYs] : []
      const placeSegments = (bounds: number[]) => {
        let prev = 0
        for (const cut of bounds) {
          const seg = cut - prev
          if (seg <= 0.5) continue
          if (!fits(seg) && !pageEmpty()) newPage(rowCursor + prev, curSection, repeatH, block.top)
          place(seg)
          prev = cut
        }
        return prev
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
          rowCursor += row.height
          if (row.isHeader && ri < headerRows) placedHeader = true
          continue
        }
      } else if (cuts.length > 0) {
        placeSegments([...cuts, row.height])
        rowCursor += row.height
        if (row.isHeader && ri < headerRows) placedHeader = true
        continue
      }
      // cantSplit / empty / no cut points: the row is atomic; turn the page first if it doesn't fit
      if (!pageEmpty()) newPage(rowCursor, curSection, repeatH, block.top)
    }
    place(row.height)
    rowCursor += row.height
    if (row.isHeader && ri < headerRows) placedHeader = true
  }
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
) {
  const totalH = block.height

  // whole paragraph fits: place directly. Trailing space doesn't consume capacity (Word breaks by text only; it may overflow into the bottom margin)
  if (fits(totalH - spaceAfterPx)) {
    place(totalH)
    return
  }

  // whole paragraph doesn't fit
  if (!lineBoxes || lineBoxes.length === 0) {
    // F1 block-level placement: push to the next page if it doesn't fit (when <= one page), otherwise F1-style hard cut
    if (totalH <= contentH) {
      if (!pageEmpty()) newPage(block.top, curSection)
    } else {
      // big block over one page (no line data): F1 style — each time reset usedOnPage to "block extends past boundary"
      // equivalent to simulating a page break every contentH
      // usedOnPage is currently u, block height H > contentH
      // needs ceil((u + H) / contentH) - 1 page turns
      // but we have no pageStartY, so we can only simulate
      // simple approach: set usedOnPage to 0 (like a big block), then keep placing
      // in practice we only need to avoid infinite loops: when totalH > contentH, place directly and let the next block trigger the page turn
      // note: the F1 algorithm handles this the same way (falls through after determining it's a big block)
    }
    place(totalH)
    return
  }

  // line-level placement
  const nLines = lineBoxes.length

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

  // text part (spaceBefore + all lines) fits, only spaceAfter overflows: Word behavior —
  // the paragraph stays on this page and the trailing space overflows into the bottom
  // margin (Word paginates by text only; corpus 14 PDF measurement: end-of-page text
  // stops at 758.9pt < bottom bound 769.9, and the overflowing space-after doesn't push the paragraph)
  let textH = spaceBeforePx
  for (const lb of lineBoxes) textH += lb.height
  if (fits(textH)) {
    place(totalH)
    return
  }

  // paragraph <= one page but doesn't fit on the current page: widow/orphan decision
  // count how many lines fit on the current page
  let splitLine = -1 // line break point (-1 = push the whole paragraph)

  if (widowOn && nLines >= 2) {
    let sumH = spaceBeforePx
    for (let li = 0; li < nLines; li++) {
      sumH += lineBoxes[li].height
      if (!fits(sumH)) {
        splitLine = li // line li doesn't fit
        break
      }
    }
    if (splitLine === -1) {
      // theoretically unreachable (the textH check covers this); conservative fallback
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
      sumH += lineBoxes[li].height
      if (!fits(sumH)) {
        splitLine = li
        break
      }
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
    // break the page before line splitLine
    if (spaceBeforePx > 0) place(spaceBeforePx)
    for (let li = 0; li < splitLine; li++) place(lineBoxes[li].height)
    // page break
    newPage(block.top + lineBoxes[splitLine].offsetInBlock, curSection)
    // place remaining lines on the new page
    for (let li = splitLine; li < nLines; li++) place(lineBoxes[li].height)
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
  if (spaceBeforePx > 0) {
    if (!fits(spaceBeforePx) && !pageEmpty()) {
      newPage(block.top, curSection)
    }
    place(spaceBeforePx)
  }
  for (const lb of lineBoxes) {
    if (!fits(lb.height) && !pageEmpty()) {
      newPage(block.top + lb.offsetInBlock, curSection)
    }
    place(lb.height)
  }
  if (spaceAfterPx > 0) {
    if (!fits(spaceAfterPx) && !pageEmpty()) {
      newPage(block.top + block.height - spaceAfterPx, curSection)
    }
    place(spaceAfterPx)
  }
}

/** Per-section header/footer content heights (px); sectionGeoms uses these to compute body push-down */
export interface SectionHfHeights {
  headerPx: number
  footerPx: number
}

/** Body top = max(marginTop, headerDist + header height) */
export function effectiveTopPx(set: SectionSettings, headerPx: number): number {
  const dist = twipsToPx(set.headerDist ?? 720)
  return Math.max(twipsToPx(set.marginTop), headerPx > 0 ? dist + headerPx : 0)
}

/** Body bottom margin = max(marginBottom, footerDist + footer height) */
export function effectiveBottomPx(set: SectionSettings, footerPx: number): number {
  const dist = twipsToPx(set.footerDist ?? 720)
  return Math.max(twipsToPx(set.marginBottom), footerPx > 0 ? dist + footerPx : 0)
}

/**
 * Uniform typed line grid (w:docGrid type lines/linesAndChars): the pitch in
 * points when EVERY section declares the same typed pitch, else null. Mixed or
 * untyped docs don't snap (matches LO only for the uniform case; per-section
 * pitches would need per-block plumbing, and mixed-grid docs are rare).
 * The value feeds .doc-page { --doc-grid-pitch } which line-height round(up)
 * expressions consume.
 */
export function docGridPitchPt(sections: SectionInfo[]): number | null {
  if (sections.length === 0) return null
  let pitch: number | null = null
  for (const s of sections) {
    const g = s.settings.docGrid
    if (!g || (g.type !== 'lines' && g.type !== 'linesAndChars') || !g.linePitch) return null
    if (pitch === null) pitch = g.linePitch
    else if (pitch !== g.linePitch) return null
  }
  return pitch === null ? null : pitch / 20
}

/** Equal-width column count of a section (w:cols w:num).
 *  equalWidth="0" (local layout columns in PDF-converted docs, varying widths) is not modeled; treated as 1 column */
export function sectionColumns(s: SectionInfo): number {
  if (/<w:cols[^>]*w:equalWidth="0"/.test(s.sectPrXml ?? '')) return 1
  return Math.max(1, s.settings.columns ?? 1)
}

/** Section column width/gap (px): equal-width columns divide evenly per w:cols w:space (default 720 twips) */
export function sectionColGeom(s: SectionInfo): {
  cols: number
  colWidthPx: number
  gapPx: number
} {
  const set = s.settings
  const contentW = twipsToPx(set.pageWidth - set.marginLeft - set.marginRight)
  const cols = sectionColumns(s)
  const gapPx = twipsToPx(set.colSpace ?? 720)
  return {
    cols,
    gapPx,
    colWidthPx: cols > 1 ? (contentW - gapPx * (cols - 1)) / cols : contentW,
  }
}

/** SectionInfo[] → pagination geometry
 *  - continuous with unchanged page geometry: no forced break (content flows on the same page)
 *  - continuous with changed page geometry (width/height change, e.g. landscape → portrait): forced break
 *  - nextPage/evenPage/oddPage: forced break
 *  - with hfHeights, oversized headers/footers squeeze body capacity
 */
export function sectionGeoms(
  sections: SectionInfo[],
  hfHeights?: SectionHfHeights[],
): SectionGeom[] {
  return sections.map((s, i) => {
    let forceBreak = false
    if (i > 0) {
      if (s.startType !== 'continuous') {
        forceBreak = true
      } else {
        // continuous section: force a page break if the page size differs from the previous section (e.g. landscape → portrait)
        const prev = sections[i - 1].settings
        const cur = s.settings
        if (prev.pageWidth !== cur.pageWidth || prev.pageHeight !== cur.pageHeight) {
          forceBreak = true
        }
      }
    }
    const set = s.settings
    const hf = hfHeights?.[i]
    const cols = sectionColumns(s)
    return {
      contentHeight:
        twipsToPx(set.pageHeight) -
        effectiveTopPx(set, hf?.headerPx ?? 0) -
        effectiveBottomPx(set, hf?.footerPx ?? 0),
      forceBreak,
      startType: s.startType,
      ...(cols > 1 ? { cols } : {}),
    }
  })
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

export type HfVariant = 'default' | 'first' | 'even'
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

function toRoman(n: number): string {
  let out = ''
  let rest = Math.max(1, Math.floor(n))
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
  let rest = Math.max(1, Math.floor(n))
  while (rest > 0) {
    rest -= 1
    out = String.fromCharCode(65 + (rest % 26)) + out
    rest = Math.floor(rest / 26)
  }
  return out
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
    case 'chineseCounting':
    case 'chineseCountingThousand':
      return toChinese(n)
    default:
      return String(n)
  }
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

/**
 * Live section list: when a non-final section's break paragraph (the block at
 * lastBlockIndex) has been deleted from the canvas, that section merges into the
 * next (content before a section break takes the following section's
 * page setup). This is derived and doesn't mutate the authoritative sections state,
 * so undoing the deletion restores naturally; readSections rebuilds after saving.
 */
export function liveSections(sections: SectionInfo[], blocks: BlockBox[]): SectionInfo[] {
  if (sections.length <= 1) return sections
  const present = new Set<number>()
  for (const b of blocks) if (b.docxIndex !== undefined) present.add(b.docxIndex)
  const out: SectionInfo[] = []
  let carryFirst: number | null = null
  let changed = false
  sections.forEach((s, i) => {
    const first = carryFirst ?? s.firstBlockIndex
    carryFirst = null
    if (i < sections.length - 1 && !present.has(s.lastBlockIndex)) {
      changed = true
      carryFirst = first
      return
    }
    out.push(first === s.firstBlockIndex ? s : { ...s, firstBlockIndex: first })
  })
  return changed ? out : sections
}

/** Tag each block's owning section by the sections' block ranges (lastBlockIndex); new blocks without docxIndex inherit from the previous block */
export function assignSections(blocks: BlockBox[], sections: SectionInfo[]): void {
  const ends = sections.map((s) => s.lastBlockIndex)
  let prev = 0
  for (const block of blocks) {
    let s = prev
    if (block.docxIndex !== undefined) {
      const i = ends.findIndex((end) => block.docxIndex! <= end)
      s = i >= 0 ? i : ends.length - 1
    }
    block.section = s
    prev = s
  }
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
    if (i === 0 || slices[i].start !== slices[i - 1].start) n++
  }
  return n
}

export interface MeasuredContent {
  blocks: BlockBox[]
  totalHeight: number
}

/** Page-bottom footnote entry (number/text/estimated height): shared by canvas page gaps and the pagination preview */
export interface PageNoteItem {
  no: number
  id: string
  text: string
  height: number
  /** rich display runs (one group per paragraph); omitted for unformatted footnotes, rendering falls back to plain text */
  richParas?: Array<
    Array<{
      text: string
      bold?: boolean
      italic?: boolean
      underline?: boolean
      strike?: boolean
      color?: string
      sizeHalfPoints?: number
    }>
  >
}

/**
 * Collect the editor's top-level block boxes (relative to the content-area top,
 * converted back to 100% zoom). origin is the content-area top's screen Y (page
 * rect.top + top margin × zoom). Page-gap decorations (.page-gap) are not content:
 * they are skipped and subtracted from subsequent block coordinates, yielding
 * "gapless continuous flow" virtual coordinates so slicing is independent of the gaps.
 */
export function measureBlocks(
  pm: HTMLElement,
  origin: number,
  zoomFactor: number,
): MeasuredContent {
  const blocks: BlockBox[] = []
  let totalHeight = 0
  let gapAccum = 0
  for (const el of Array.from(pm.children) as HTMLElement[]) {
    const rect = el.getBoundingClientRect()
    if (el.classList.contains('page-gap')) {
      gapAccum += rect.height
      continue
    }
    if (rect.height <= 0) continue
    // in-block gaps from mid-paragraph page breaks: subtract from block height and add to the gap accumulator for later blocks
    const innerGap = innerGapHeight(el)
    const top = (rect.top - origin - gapAccum) / zoomFactor
    const height = (rect.height - innerGap) / zoomFactor
    const idxAttr = el.getAttribute('data-idx')
    const hasBreak = !!el.querySelector('.doc-field-pagebreak, .doc-page-br')
    // break-only paragraph: count the whole block (br line + ProseMirror trailing-break
    // phantom line) as trailing space so it never pushes to — and blanks — a page. Word's
    // real rule is that the br/paragraph-mark line must fit or the paragraph pushes into a
    // deliberate blank page, but our page fill drifts a few px per page in both directions,
    // and a fit-check flips borderline pages into spurious mid-document blanks (worse than
    // occasionally dropping an intentional one)
    const breakOnly = hasBreak && !(el.textContent ?? '').trim() && !el.querySelector('img')
    blocks.push({
      top,
      height,
      breakBefore: el.classList.contains('page-break-before') || undefined,
      breakAfter: hasBreak || undefined,
      el,
      ...(breakOnly ? { spaceAfterPx: height } : {}),
      ...(idxAttr ? { docxIndex: parseInt(idxAttr, 10) } : {}),
    })
    gapAccum += innerGap
    totalHeight = Math.max(totalHeight, top + height)
  }
  // inter-block CSS margin (space after): rect height excludes it, but it occupies
  // vertical layout space. Attribute it to the previous block's spaceAfterPx and add
  // it to the height, so the engine's capacity bookkeeping matches Y coordinates and
  // the "trailing space doesn't consume page capacity" rule (Word breaks by text only) applies.
  for (let i = 0; i + 1 < blocks.length; i++) {
    const gap = blocks[i + 1].top - (blocks[i].top + blocks[i].height)
    if (gap > 0.5) {
      blocks[i].spaceAfterPx = (blocks[i].spaceAfterPx ?? 0) + gap
      blocks[i].height += gap
    }
  }
  return { blocks, totalHeight }
}

/**
 * Endnote layout: endnotes gather at the end of the document
 * (or section) right after the body, flowing to later pages when they don't fit.
 * Before slicing, the endnotes area is appended as a virtual block at flow end: one
 * line box per endnote (separator height merged into the first), widowControl off →
 * page breaks are allowed between any entries. Returns the endnotes area's top Y.
 */
export function appendEndnotesBlock(
  blocks: BlockBox[],
  totalHeight: number,
  items: PageNoteItem[],
  separatorH: number,
): { totalHeight: number; top: number } | null {
  if (items.length === 0) return null
  const top = totalHeight
  const lineBoxes: Array<{ offsetInBlock: number; height: number }> = []
  let off = 0
  for (let i = 0; i < items.length; i++) {
    const h = (i === 0 ? separatorH : 0) + items[i].height
    lineBoxes.push({ offsetInBlock: off, height: h })
    off += h
  }
  blocks.push({
    top,
    height: off,
    lineBoxes,
    widowControl: false,
    isEndnotes: true,
    ...(blocks.length > 0 && blocks[blocks.length - 1].section !== undefined
      ? { section: blocks[blocks.length - 1].section }
      : {}),
  })
  return { totalHeight: top + off, top }
}

/**
 * Two-pass slicing: slice by block first, then collect DOM line-box boundaries for
 * blocks crossing page bounds and re-slice. Only blocks that cross a page or exceed
 * one page get line collection (at most one per page, negligible cost).
 * metaOf: docxIndex → parse-layer pagination constraints (keepNext/widow/table row flags).
 */
export function sliceWithLineSplit(
  blocks: BlockBox[],
  geoms: SectionGeom[],
  totalHeight: number,
  zoomFactor: number,
  metaOf?: BlockMetaOf,
): PageSlice[] {
  if (metaOf) applyBlockMeta(blocks, metaOf)
  const slices = computeSectionedSlicesF2(blocks, geoms, totalHeight)
  const changed = fillLineBoxes(blocks, geoms, zoomFactor, slices, metaOf)
  return insertParityBlanks(
    changed ? computeSectionedSlicesF2(blocks, geoms, totalHeight) : slices,
    geoms,
  )
}

/**
 * Collect DOM line-box data for pagination candidate blocks (F2 model): blocks that
 * cross a page bound, exceed one page, or were pushed wholesale to a page top (the
 * second pass may pull lines back to the previous page). Other blocks are skipped, so cost is negligible.
 * Table blocks → tableRows (tr boundaries; never cuts into text lines inside cells); text blocks → lineBoxes.
 * Returns whether any block was filled (true means the caller must re-slice).
 */
/**
 * DOM line/row sampling is the hot path of repeated repagination: the set of
 * page-crossing blocks is stable across edits, so raw samples are cached by
 * element identity plus a cheap content/geometry signature. Entries drop with
 * their element (WeakMap) or when the signature stops matching.
 */
const lineSampleCache = new WeakMap<
  HTMLElement,
  { sig: string; boundaries?: number[]; rows?: TableRowBox[] }
>()

function lineSampleSig(el: HTMLElement, textH: number): string {
  // djb2 over the text: equal-length edits must still invalidate
  const text = el.textContent ?? ''
  let h = 5381
  for (let i = 0; i < text.length; i++) h = ((h << 5) + h + text.charCodeAt(i)) | 0
  // width guards width-only reflows; descendant count guards nested (e.g. table
  // cell) structure changes that keep the direct-child count. A stale miss only
  // costs one re-sample, so quantization errs toward invalidating.
  const w = el.getBoundingClientRect().width
  const nodes = el.getElementsByTagName('*').length
  return `${Math.round(textH * 4)}:${Math.round(w * 4)}:${nodes}:${h}`
}

export function fillLineBoxes(
  blocks: BlockBox[],
  geoms: SectionGeom[],
  zoomFactor: number,
  slices?: PageSlice[],
  metaOf?: BlockMetaOf,
): boolean {
  const geomOf = (s: number) => geoms[Math.max(0, Math.min(s, geoms.length - 1))]
  // cut bounds = page bounds + column bounds of multi-column pages (blocks crossing within a column also need line-level splits)
  const breaks: number[] = []
  ;(slices ?? []).forEach((s, i) => {
    if (i > 0) breaks.push(s.start)
    for (const r of s.regions ?? []) {
      for (const c of r.columns) {
        if (c.start > 0.5 && !breaks.includes(c.start)) breaks.push(c.start)
      }
    }
  })
  let changed = false
  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i]
    if (!block.el || block.lineBoxes || block.tableRows) continue
    const contentH = geomOf(block.section ?? 0)?.contentHeight ?? 0
    if (contentH <= 0) continue
    const bottom = block.top + block.height
    const crossing = breaks.some((y) => block.top < y && y < bottom)
    const atPageTop = breaks.some((y) => Math.abs(block.top - y) < 0.5)
    if (
      block.height <= contentH &&
      !crossing &&
      !atPageTop &&
      // tables anchoring a keepNext chain always need row data: the chain head only
      // has to share a page with the first row, so the whole-table height is misleading
      !(i > 0 && blocks[i - 1].keepNext && !block.keepNext && block.el.querySelector('tr'))
    )
      continue

    // line boxes tile only the text area (block height includes the merged-in space-after, which lines must not cover)
    const textH = block.height - (block.spaceAfterPx ?? 0)
    const sig = lineSampleSig(block.el, textH)
    const cached = lineSampleCache.get(block.el)
    const hit = cached?.sig === sig ? cached : null

    if (block.el.querySelector('tr')) {
      // flags mutate the rows, so cached rows are cloned per use
      const rows = hit?.rows
        ? hit.rows.map((r) => ({ ...r }))
        : domTableRows(block.el, textH, zoomFactor)
      if (!hit?.rows) lineSampleCache.set(block.el, { sig, rows: rows.map((r) => ({ ...r })) })
      if (rows.length > 0) {
        const flags =
          block.docxIndex !== undefined ? metaOf?.(block.docxIndex)?.tableRowFlags : undefined
        if (flags)
          rows.forEach((r, i) => {
            if (flags[i]?.isHeader) r.isHeader = true
            if (flags[i]?.cantSplit) r.cantSplit = true
          })
        block.tableRows = rows
        changed = true
      }
      continue
    }
    // synthesized over-page cuts below mutate the list, so cached entries are copied out
    const boundaries = hit?.boundaries
      ? [...hit.boundaries]
      : domLineBoundaries(block.el, zoomFactor)
    if (!hit?.boundaries) lineSampleCache.set(block.el, { sig, boundaries: [...boundaries] })
    if (boundaries.length === 0 && block.height > contentH) {
      // over-page block with no text lines (e.g. a large image): synthesize cut points at page height, equivalent to hard pixel cuts
      for (let y = contentH; y < block.height; y += contentH) boundaries.push(y)
    }
    if (boundaries.length > 0) {
      block.lineBoxes = tileBoxes(boundaries, textH)
      changed = true
    }
  }
  return changed
}

/** Boundary list (excluding 0) → line boxes tiling the block height (heights are adjacent-boundary diffs; the first box starts at 0) */
function tileBoxes(
  boundaries: number[],
  blockHeight: number,
): Array<{ offsetInBlock: number; height: number }> {
  const tops = [0, ...boundaries.filter((b) => b > 0.5 && b < blockHeight)]
  return tops.map((top, i) => ({
    offsetInBlock: top,
    height: (i + 1 < tops.length ? tops[i + 1] : blockHeight) - top,
  }))
}

/** Extract each tr's tblHeader/cantSplit flags from table XML (header repetition across breaks / unsplittable rows) */
export function tableRowFlags(tableXml: string): Array<{ isHeader: boolean; cantSplit: boolean }> {
  const flags: Array<{ isHeader: boolean; cantSplit: boolean }> = []
  for (const m of tableXml.matchAll(/<w:tr[\s>][\s\S]*?(?=<w:tr[\s>]|<\/w:tbl>)/g)) {
    const trPr = m[0].match(/<w:trPr>[\s\S]*?<\/w:trPr>/)?.[0] ?? ''
    flags.push({
      isHeader: /<w:tblHeader(?!\s+w:val="(?:0|false)")/.test(trPr),
      cantSplit: /<w:cantSplit(?!\s+w:val="(?:0|false)")/.test(trPr),
    })
  }
  return flags
}

/** Extract each tr's tblHeader flag from table XML (header repeated at page top after a table break) */
export function tableHeaderFlags(tableXml: string): boolean[] {
  return tableRowFlags(tableXml).map((f) => f.isHeader)
}

/**
 * Block-level pagination constraints (injection channel from parse-layer results into DOM-measured blocks).
 * The canvas/preview measureBlocks only has geometry; semantics like keepNext are attached by docxIndex.
 */
export interface BlockMeta {
  keepNext?: boolean
  keepLines?: boolean
  /** false only when explicitly disabled (Word default on) */
  widowControl?: false
  /** table blocks: per-tr header/unsplittable flags (applied by fillLineBoxes when collecting rows) */
  tableRowFlags?: Array<{ isHeader: boolean; cantSplit: boolean }>
  /** page-bottom height reserved for footnote refs inside the block (px): merged into block height and space-after (doesn't consume text capacity) */
  footnoteExtraPx?: number
}

export type BlockMetaOf = (docxIndex: number) => BlockMeta | undefined

/** Inject parse-layer constraints into measured blocks (call before slicing; table row flags are applied by fillLineBoxes) */
export function applyBlockMeta(blocks: BlockBox[], metaOf: BlockMetaOf): void {
  for (const b of blocks) {
    if (b.docxIndex === undefined) continue
    const meta = metaOf(b.docxIndex)
    if (!meta) continue
    if (meta.keepNext) b.keepNext = true
    if (meta.keepLines) b.keepLines = true
    if (meta.widowControl === false) b.widowControl = false
    if (meta.footnoteExtraPx) {
      // matches the parity model: footnote height enters the block-height bookkeeping to consume page capacity, and also the space-after
      // (fits' "text fits" check subtracts space-after, so the reservation doesn't affect the paragraph's own line breaking)
      b.height += meta.footnoteExtraPx
      b.spaceAfterPx = (b.spaceAfterPx ?? 0) + meta.footnoteExtraPx
    }
  }
}

/** Table block: one line box per tr, heights tiling the block height (borders folded into first/last rows).
 *  In-table page gaps (table-break decoration rows) don't count as rows; their height is subtracted from the offsets of rows below */
function domTableRows(el: HTMLElement, blockHeight: number, zoomFactor: number): TableRowBox[] {
  const gaps = Array.from(el.querySelectorAll('.page-gap-inline')).map((g) =>
    g.getBoundingClientRect(),
  )
  const gapAbove = (top: number) => gaps.reduce((s, g) => (g.top <= top ? s + g.height : s), 0)
  const elTop = el.getBoundingClientRect().top
  // take only the outer table's real rows: trs of nested tables inside cells
  // (.doc-nested-table) are in-row content, and decoration rows (page gaps /
  // repeated tblHeader clones) are not page-split units — counting them would
  // add phantom boundaries and shift the tableRowFlags index alignment
  const trs = Array.from(el.querySelectorAll('tr')).filter(
    (tr) =>
      !tr.closest('.doc-nested-table') &&
      !tr.classList.contains('page-gap') &&
      !tr.classList.contains('page-repeat-header'),
  )
  const tops: number[] = []
  for (const tr of trs) {
    const trTop = tr.getBoundingClientRect().top
    const off = (trTop - elTop - gapAbove(trTop)) / zoomFactor
    if (off > 0.5) tops.push(off)
  }
  return tileBoxes(tops, blockHeight).map((b, i) => {
    if (!trs[i]) return { height: b.height }
    const { cuts, contentBottom } = rowCutYs(
      trs[i],
      b.offsetInBlock,
      b.height,
      elTop,
      gapAbove,
      zoomFactor,
    )
    return { height: b.height, contentBottom, ...(cuts.length > 0 ? { cutYs: cuts } : {}) }
  })
}

/** In-row safe cut points (relative to row top, px, ascending): line-level candidates
 *  per cell (Word breaks between any two lines), rejecting cuts that would cross a
 *  line box in another cell. Also reports the lowest content-band bottom. */
function rowCutYs(
  tr: Element,
  rowTop: number,
  rowHeight: number,
  elTop: number,
  gapAbove: (top: number) => number,
  zoomFactor: number,
): { cuts: number[]; contentBottom: number } {
  const cells = Array.from(tr.children).filter((c) => c.tagName === 'TD' || c.tagName === 'TH')
  const range = document.createRange()
  const cellBands: Array<Array<[number, number]>> = []
  for (const cell of cells) {
    const bands: Array<[number, number]> = []
    const add = (r: DOMRect) => {
      if (r.height <= 0 || r.width <= 0) return
      bands.push([
        (r.top - elTop - gapAbove(r.top)) / zoomFactor - rowTop,
        (r.bottom - elTop - gapAbove(r.bottom)) / zoomFactor - rowTop,
      ])
    }
    const walker = document.createTreeWalker(cell, NodeFilter.SHOW_TEXT)
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      if (n.parentElement?.closest('.page-gap')) continue
      range.selectNodeContents(n)
      for (const r of range.getClientRects()) add(r)
    }
    for (const obj of cell.querySelectorAll('img, svg, canvas')) {
      // in-cell gap decorations may carry header/footer images: not row content
      if (obj.closest('.page-gap')) continue
      add(obj.getBoundingClientRect())
    }
    if (bands.length > 0) cellBands.push(bands)
  }
  const contentBottom = cellBands.reduce(
    (max, bands) => bands.reduce((m, [, b]) => Math.max(m, b), max),
    0,
  )
  return { cuts: cellCutYs(cellBands, rowHeight), contentBottom }
}

/** Rects sharing vertical overlap collapse into one line interval (same criterion as domLineRects) */
function clusterLineBands(bands: Array<[number, number]>): Array<[number, number]> {
  const sorted = [...bands].sort((a, b) => a[0] - b[0])
  const lines: Array<[number, number]> = []
  for (const [top, bottom] of sorted) {
    const last = lines[lines.length - 1]
    if (last && top < last[1] - 1) last[1] = Math.max(last[1], bottom)
    else lines.push([top, bottom])
  }
  return lines
}

/** Pure core of rowCutYs (testable without DOM): per-cell rect bands → safe cut ys.
 *  Candidates are midpoints between a cell's consecutive lines (a zero gap still
 *  counts); a candidate falling inside any cell's line box is unsafe (0.5px
 *  tolerance for sub-pixel jitter). */
export function cellCutYs(cellBands: Array<Array<[number, number]>>, rowHeight: number): number[] {
  const cellLines = cellBands.map(clusterLineBands)
  const candidates: number[] = []
  for (const lines of cellLines) {
    for (let i = 0; i + 1 < lines.length; i++) {
      candidates.push((lines[i][1] + lines[i + 1][0]) / 2)
    }
  }
  candidates.sort((a, b) => a - b)
  const cuts: number[] = []
  for (const y of candidates) {
    if (y <= 2 || y >= rowHeight - 2) continue
    if (cellLines.some((lines) => lines.some(([t, b]) => y > t + 0.5 && y < b - 0.5))) continue
    if (cuts.length > 0 && y - cuts[cuts.length - 1] < 1) continue
    cuts.push(y)
  }
  return cuts
}

/** Total height of in-block inline gaps (mid-paragraph page-break decorations) (screen px) */
function innerGapHeight(el: HTMLElement): number {
  let sum = 0
  for (const g of el.querySelectorAll('.page-gap-inline')) sum += g.getBoundingClientRect().height
  return sum
}

/**
 * In-block text lines (first rect of each line): offset is the virtual in-block Y
 * after subtracting inline gaps; left/top are screen coordinates; node is the text
 * node owning the line's first rect (DOM anchor for viewport-independent positioning).
 * Text inside gaps (e.g. footnotes) doesn't count as lines.
 */
function domLineRects(
  el: HTMLElement,
  zoomFactor: number,
): Array<{ offset: number; left: number; top: number; node: Text }> {
  const gaps = Array.from(el.querySelectorAll('.page-gap-inline')).map((g) =>
    g.getBoundingClientRect(),
  )
  const gapAbove = (top: number) => gaps.reduce((s, g) => (g.top <= top ? s + g.height : s), 0)
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT)
  const range = document.createRange()
  const rects: Array<{ r: DOMRect; node: Text }> = []
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if (n.parentElement?.closest('.page-gap')) continue
    range.selectNodeContents(n)
    for (const r of range.getClientRects()) {
      if (r.height > 0 && r.width > 0) rects.push({ r, node: n as Text })
    }
  }
  rects.sort((a, b) => a.r.top - b.r.top)
  const elTop = el.getBoundingClientRect().top
  const lines: Array<{ offset: number; left: number; top: number; node: Text }> = []
  let lineBottom = -Infinity
  for (const { r, node } of rects) {
    if (r.top >= lineBottom - 1) {
      lines.push({
        offset: (r.top - elTop - gapAbove(r.top)) / zoomFactor,
        left: r.left,
        top: r.top,
        node,
      })
      lineBottom = r.bottom
    } else {
      lineBottom = Math.max(lineBottom, r.bottom)
      const last = lines[lines.length - 1]
      if (last && r.left < last.left) last.left = r.left
    }
  }
  return lines
}

function domLineBoundaries(el: HTMLElement, zoomFactor: number): number[] {
  return lineBreakBoundaries(domLineRects(el, zoomFactor).map((ln) => ln.offset))
}

/**
 * Convert DOM text-rect tops into safe line-break boundaries.
 *
 * The first rect is the glyph box inside the first line box, so its top can be
 * a few pixels below the block top. Treating it as a boundary creates a phantom
 * first line and lets pagination clip through glyphs. Only subsequent line
 * starts are valid page-break positions.
 */
export function lineBreakBoundaries(lineOffsets: number[]): number[] {
  return lineOffsets.slice(1).filter((off) => off > 0.5)
}

/** DOM anchor of a line start: the line's first text node + character offset within it (feed to view.posAtDOM) */
export interface LineAnchor {
  node: Text
  charOffset: number
}

/**
 * Character offset within a text node where the line whose top is lineTop begins.
 * Uses per-character Range rects (layout data, not viewport hit-testing), so it works
 * for lines scrolled outside the viewport — posAtCoords/caretRangeFromPoint do not:
 * off-screen coordinates resolve to degenerate document positions, which used to drop
 * in-table cut markers before the table's first row where they inflate the canvas
 * table by an anonymous-row line-height and skew all pagination measurement below.
 */
function lineStartCharOffset(node: Text, lineTop: number): number {
  const len = node.length
  if (len === 0) return 0
  const range = document.createRange()
  const topAt = (i: number): number => {
    range.setStart(node, i)
    range.setEnd(node, i + 1)
    for (const r of range.getClientRects()) if (r.height > 0) return r.top
    // collapsed characters (e.g. wrap-point whitespace) have no rect: treat as belonging to an earlier line
    return -Infinity
  }
  // first character at/below the line top (character tops are non-decreasing in flowing text)
  let lo = 0
  let hi = len - 1
  let ans = 0
  while (lo <= hi) {
    const mid = (lo + hi) >> 1
    if (topAt(mid) >= lineTop - 1) {
      ans = mid
      hi = mid - 1
    } else {
      lo = mid + 1
    }
  }
  return ans
}

const toAnchor = (ln: { node: Text; top: number }): LineAnchor => ({
  node: ln.node,
  charOffset: lineStartCharOffset(ln.node, ln.top),
})

/**
 * DOM anchor of the line start matching an in-block virtual Y (offsetInBlock)
 * (used to position mid-paragraph page-break decorations).
 * Returns null when no matching line is found (non-text block / hard pixel cut point).
 */
export function lineStartAnchor(
  el: HTMLElement,
  offsetInBlock: number,
  zoomFactor: number,
): LineAnchor | null {
  for (const ln of domLineRects(el, zoomFactor)) {
    if (Math.abs(ln.offset - offsetInBlock) < 1.5) return toAnchor(ln)
  }
  return null
}

/** DOM anchor of the first line at or after (≥) a given in-block Y: used by in-row cut points (cuts in inter-line gaps) to locate the next page's first line */
export function nextLineAnchor(
  el: HTMLElement,
  offsetInBlock: number,
  zoomFactor: number,
): LineAnchor | null {
  for (const ln of domLineRects(el, zoomFactor)) {
    if (ln.offset >= offsetInBlock - 1.5) return toAnchor(ln)
  }
  return null
}

/** In-row cut decoration policy: a single-cell row can host a real inline gap band
 *  (one cell spans the whole cut) — returns that cell; multi-cell rows keep the
 *  zero-height cut marker (same-y bands across cells are not modeled) — returns null. */
export function singleCutCell(row: Element | null): Element | null {
  const cells = row
    ? Array.from(row.children).filter((c) => c.tagName === 'TD' || c.tagName === 'TH')
    : []
  return cells.length === 1 ? cells[0] : null
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
