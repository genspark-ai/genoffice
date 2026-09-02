// DOM line-box sampling for page-crossing blocks (cached per element), the
// line-split re-slice, table row cut positions and line anchors.
import { applyBlockMeta } from './pagination-measure'
import {
  computeSectionedSlicesF2,
  insertParityBlanks,
  sectionFirstPages,
} from './pagination-slices'
import type {
  BlockBox,
  BlockMetaOf,
  PageSlice,
  SectionGeom,
  SliceOutputs,
  TableRowBox,
} from './pagination-types'

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
  out?: SliceOutputs,
): PageSlice[] {
  if (metaOf) applyBlockMeta(blocks, metaOf, zoomFactor)
  let slices = computeSectionedSlicesF2(blocks, geoms, totalHeight, out)
  // re-slicing can surface new candidate blocks (a block pushed to a page top only
  // after an earlier block gained line data) — iterate to a fixed point, bounded.
  // The cascade can run one block per page boundary (a dense two-column grid doc
  // packs tighter on every pass: SAS prod_043 left whole paragraphs unsplit at
  // column bottoms with a bound of 3), so the bound follows the page count.
  const maxPasses = Math.max(3, Math.min(24, slices.length))
  for (let i = 0; i < maxPasses; i++) {
    const changed = fillLineBoxes(blocks, geoms, zoomFactor, slices, metaOf)
    if (!changed) break
    slices = computeSectionedSlicesF2(blocks, geoms, totalHeight, out)
  }
  return insertParityBlanks(slices, geoms)
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
  { sig: string; boundaries?: number[]; rows?: TableRowBox[]; soleLineBottom?: number }
>()

// webfont loads shift line boxes without changing block height (explicit line
// heights), so the geometry/content signature alone cannot see them
let lineSampleFontEpoch = 0

export function bumpLineSampleFontEpoch(): void {
  lineSampleFontEpoch++
}

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
  return `${lineSampleFontEpoch}:${Math.round(textH * 4)}:${Math.round(w * 4)}:${nodes}:${h}`
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
  const firsts = sectionFirstPages(slices ?? [])
  let changed = false
  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i]
    if (!block.el || block.lineBoxes || block.tableRows) continue
    // Anchored textbox/shape blocks are atomic like Word shapes: their inner
    // text lines (usually inside fixed, clipped boxes) are not page-break
    // points. Left line-less, an over-page block places whole and overlaps the
    // bottom margin; the next block turns the page (Word-like shape overflow).
    if (block.el.classList.contains('doc-protected-textboxes')) continue
    const contentH = geomOf(block.section ?? 0)?.contentHeight ?? 0
    if (contentH <= 0) continue
    const bottom = block.top + block.height
    const crossing = breaks.some((y) => block.top < y && y < bottom)
    const atPageTop = breaks.some((y) => Math.abs(block.top - y) < 0.5)
    // region-aware capacity: a block in a mixed-column page's later region has
    // only the region's height, not the full page — a table there must get row
    // data even when its height fits a page (otherwise the first pass places
    // it whole as an "over-page" block and the collapse becomes a fixed point:
    // the ANSI table after a 3-col region, real_run2/61)
    let capH = contentH
    for (const s of slices ?? []) {
      for (const r of s.regions ?? []) {
        if (r.columns.some((c) => c.start - 0.5 <= block.top && block.top < c.end - 0.5)) {
          capH = Math.min(capH, r.height)
        }
      }
    }
    // a block on a section's FIRST page has only the titlePg capacity: it must
    // get line/row data even when it fits the default page height, or the first
    // pass places it whole and the collapse becomes a fixed point
    slices?.forEach((s, si) => {
      if (!firsts[si] || block.top < s.start - 0.5 || block.top >= s.end - 0.5) return
      const fc = geomOf(s.section)?.firstContentHeight
      if (fc !== undefined) capH = Math.min(capH, fc)
    })
    // a block already carrying the renderer-baked oversize clip measures at the
    // clipped (fitting) height, so it must bypass the fit gates to re-qualify —
    // otherwise the flag drops, the clip clears, and the layout oscillates
    const clipMarked = block.el.dataset.oversizeClip !== undefined
    if (
      !clipMarked &&
      block.height <= capH &&
      !crossing &&
      !atPageTop &&
      // chain anchors always need line/row data: the chain only keeps with the
      // anchor's first line(s)/row, so the whole-block height is misleading
      !(i > 0 && blocks[i - 1].keepNext && !block.keepNext)
    )
      continue
    // a protected image block is atomic like a shape: the tiny lead line its
    // anchor marker / leading spaces form is not a page-break point (cutting
    // there left a sliver strip of a page-crossing photo behind). Over-page
    // images still fall through for the oversize-clip flag.
    if (!clipMarked && block.el.classList.contains('doc-protected-image') && block.height <= capH)
      continue

    // line boxes tile only the text area (block height includes the merged-in
    // space-after, any folded-in leading space-before, and the footnote
    // reservation, none of which lines may cover)
    const textH =
      block.height -
      (block.spaceAfterPx ?? 0) -
      (block.spaceBeforePx ?? 0) -
      (block.footnoteExtraPx ?? 0)
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
            // Editable native tables publish an explicit live value on each tr;
            // it must beat the source XML so turning repetition off takes effect
            // before the document is saved and reopened.
            if (r.isHeader === undefined && flags[i]?.isHeader) r.isHeader = true
            if (flags[i]?.cantSplit) r.cantSplit = true
            if (flags[i]?.minHPx) r.minHPx = flags[i].minHPx
          })
        block.tableRows = rows
        changed = true
      }
      continue
    }
    // synthesized over-page cuts below mutate the list, so cached entries are copied out
    let boundaries: number[]
    let soleLineBottom: number | undefined
    if (hit?.boundaries) {
      boundaries = [...hit.boundaries]
      soleLineBottom = hit.soleLineBottom
    } else {
      const lines = domLineRects(block.el, zoomFactor)
      boundaries = lineBreakBoundaries(lines)
      soleLineBottom = lines.length === 1 ? lines[0].bottom : undefined
      lineSampleCache.set(block.el, {
        sig,
        boundaries: [...boundaries],
        ...(soleLineBottom !== undefined ? { soleLineBottom } : {}),
      })
    }
    // sole line taller than the column (oversized inline picture): no break
    // points exist, so flag the block for the atomic page-bottom clip instead
    // of cutting (Word overflow-clips the line; see the placement branch)
    if (boundaries.length === 0 && soleLineBottom !== undefined && soleLineBottom > capH + 0.5) {
      if (block.oversizeLineH !== soleLineBottom) {
        block.oversizeLineH = soleLineBottom
        changed = true
      }
      continue
    }
    if (boundaries.length === 0 && block.height > contentH) {
      // over-page block with no in-flow lines at all (floated/absolute content):
      // synthesize cut points at page height, equivalent to hard pixel cuts
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
  // skip trs[0]: the first row starts at box 0 by definition — its measured
  // offset is just the collapsed-border half-width (1px at w:sz=12), and
  // letting it through creates a phantom row that shifts the trs[i] pairing
  for (const tr of trs.slice(1)) {
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
    return {
      height: b.height,
      contentBottom,
      ...(trs[i].hasAttribute('data-repeat-header')
        ? { isHeader: trs[i].getAttribute('data-repeat-header') === '1' }
        : {}),
      ...(cuts.length > 0 ? { cutYs: cuts } : {}),
    }
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
  const paraBands: Array<Array<[number, number]>> = []
  for (const cell of cells) {
    const bands: Array<[number, number]> = []
    const byPara = new Map<Element, Array<[number, number]>>()
    const toBand = (r: DOMRect): [number, number] => [
      (r.top - elTop - gapAbove(r.top)) / zoomFactor - rowTop,
      (r.bottom - elTop - gapAbove(r.bottom)) / zoomFactor - rowTop,
    ]
    const walker = document.createTreeWalker(cell, NodeFilter.SHOW_TEXT)
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      const parent = n.parentElement
      if (parent?.closest('.page-gap, .page-float-host')) continue
      const para = parent?.closest('p, h1, h2, h3, h4, h5, h6, .doc-li')
      range.selectNodeContents(n)
      for (const r of range.getClientRects()) {
        if (r.height <= 0 || r.width <= 0) continue
        const band = toBand(r)
        bands.push(band)
        if (para) {
          const list = byPara.get(para) ?? []
          list.push(band)
          byPara.set(para, list)
        }
      }
    }
    for (const obj of cell.querySelectorAll('img, svg, canvas')) {
      // in-cell gap decorations may carry header/footer images: not row content
      if (obj.closest('.page-gap, .page-float-host')) continue
      const r = obj.getBoundingClientRect()
      if (r.height > 0 && r.width > 0) bands.push(toBand(r))
    }
    if (bands.length > 0) cellBands.push(bands)
    for (const list of byPara.values()) paraBands.push(list)
  }
  const contentBottom = cellBands.reduce(
    (max, bands) => bands.reduce((m, [, b]) => Math.max(m, b), max),
    0,
  )
  return { cuts: cellCutYs(cellBands, rowHeight, paraBands), contentBottom }
}

/** Rects sharing vertical overlap collapse into one line interval. Same-line rects
 *  overlap near-fully; adjacent tight table rows merely graze (line boxes 1-2px
 *  taller than the row pitch), and chain-merging them would swallow a whole
 *  nested table into one cut-less band (fdo48718), so a merge needs substantial
 *  overlap relative to the smaller band. */
function clusterLineBands(bands: Array<[number, number]>): Array<[number, number]> {
  const sorted = [...bands].sort((a, b) => a[0] - b[0])
  const lines: Array<[number, number]> = []
  for (const [top, bottom] of sorted) {
    const last = lines[lines.length - 1]
    const overlap = last ? last[1] - top : 0
    const minH = last ? Math.min(last[1] - last[0], bottom - top) : 0
    if (last && overlap > 1 && overlap > 0.4 * minH) last[1] = Math.max(last[1], bottom)
    else lines.push([top, bottom])
  }
  return lines
}

/** Pure core of rowCutYs (testable without DOM): per-cell rect bands → safe cut ys.
 *  Candidates are midpoints between a cell's consecutive lines (a zero gap still
 *  counts); a candidate falling inside any cell's line box is unsafe (0.5px
 *  tolerance for sub-pixel jitter). paraBands (per-paragraph line bands across
 *  all cells) add Word's widow/orphan rule to in-row breaks: a cut through a
 *  paragraph must leave at least two of its lines on each side, else the whole
 *  row pushes (Word pushes a row whose 2-line cell would split, prod100r4/32+47). */
export function cellCutYs(
  cellBands: Array<Array<[number, number]>>,
  rowHeight: number,
  paraBands?: Array<Array<[number, number]>>,
): number[] {
  const cellLines = cellBands.map(clusterLineBands)
  const paraLines = (paraBands ?? []).map(clusterLineBands)
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
    // 2px tolerance: grazing line boxes of tight table rows overlap their row
    // boundary by ~1px, and the between-rows midpoint must stay a legal cut
    if (cellLines.some((lines) => lines.some(([t, b]) => y > t + 2 && y < b - 2))) continue
    if (
      paraLines.some((lines) => {
        const above = lines.filter(([, b]) => b <= y + 2).length
        const below = lines.filter(([t]) => t >= y - 2).length
        // the cut is inside this paragraph only when lines exist on both sides
        return above > 0 && below > 0 && (above < 2 || below < 2)
      })
    )
      continue
    if (cuts.length > 0 && y - cuts[cuts.length - 1] < 1) continue
    cuts.push(y)
  }
  return cuts
}

/**
 * In-block text lines (first rect of each line): offset is the virtual in-block Y
 * after subtracting inline gaps; left/top are screen coordinates; node is the text
 * node owning the line's first rect (DOM anchor for viewport-independent positioning),
 * or an in-flow inline image element when the line holds no text (picture-only lines).
 * Text inside gaps (e.g. footnotes) doesn't count as lines.
 */
type DomLineRect = {
  offset: number
  /** virtual in-block Y of the line's lowest ink (same space as offset) */
  bottom: number
  left: number
  top: number
  node: Text | Element
}

export type DomLineRectsFn = (el: HTMLElement, zoomFactor: number) => DomLineRect[]

/**
 * Per-pass memo for domLineRects: one remeasure can query hundreds of cut anchors
 * against the same block, and each query re-walks every text rect (forced layout
 * reads). Scope the cache to a single pass — DOM/scroll are stable within it.
 */
export function createLineRectsCache(): DomLineRectsFn {
  const memo = new Map<HTMLElement, DomLineRect[]>()
  return (el, zoomFactor) => {
    let lines = memo.get(el)
    if (!lines) {
      lines = domLineRects(el, zoomFactor)
      memo.set(el, lines)
    }
    return lines
  }
}

/** In normal flow inside el (no floated/absolutely-positioned ancestor): only such
 *  content forms text lines — overlays and wrap-floats don't consume flow height. */
function inFlowWithin(node: Element, el: HTMLElement): boolean {
  for (let e: Element | null = node; e && e !== el; e = e.parentElement) {
    const cs = getComputedStyle(e)
    // jsdom leaves unset properties '' — treat as in flow
    if ((cs.float && cs.float !== 'none') || cs.position === 'absolute' || cs.position === 'fixed')
      return false
  }
  return true
}

function domLineRects(el: HTMLElement, zoomFactor: number): DomLineRect[] {
  const gaps = Array.from(el.querySelectorAll('.page-gap-inline')).map((g) =>
    g.getBoundingClientRect(),
  )
  const gapAbove = (top: number) => gaps.reduce((s, g) => (g.top <= top ? s + g.height : s), 0)
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT)
  const range = document.createRange()
  const rects: Array<{ r: DOMRect; node: Text | Element }> = []
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if (n.parentElement?.closest('.page-gap, .page-float-host')) continue
    range.selectNodeContents(n)
    for (const r of range.getClientRects()) {
      if (r.height > 0 && r.width > 0) rects.push({ r, node: n as Text })
    }
  }
  // in-flow inline pictures form lines too (a picture-only paragraph has no text
  // rect at all, which used to leave over-page image stacks without break points)
  for (const im of Array.from(el.querySelectorAll('img'))) {
    if (im.closest('.page-gap, .page-float-host')) continue
    if (!inFlowWithin(im, el)) continue
    const r = im.getBoundingClientRect()
    if (r.height > 0 && r.width > 0) rects.push({ r, node: im })
  }
  rects.sort((a, b) => a.r.top - b.r.top)
  const elTop = el.getBoundingClientRect().top
  const lines: DomLineRect[] = []
  let lineBottom = -Infinity
  for (const { r, node } of rects) {
    if (r.top >= lineBottom - 1) {
      lines.push({
        offset: (r.top - elTop - gapAbove(r.top)) / zoomFactor,
        bottom: (r.bottom - elTop - gapAbove(r.top)) / zoomFactor,
        left: r.left,
        top: r.top,
        node,
      })
      lineBottom = r.bottom
    } else {
      lineBottom = Math.max(lineBottom, r.bottom)
      const last = lines[lines.length - 1]
      if (last) {
        last.bottom = Math.max(last.bottom, (r.bottom - elTop - gapAbove(r.top)) / zoomFactor)
        if (r.left < last.left) last.left = r.left
        // text anchors the line whenever it has any (image rects only stand in
        // on text-less lines; char anchors keep RTL/offset resolution exact)
        if (last.node instanceof Element && !(node instanceof Element)) {
          last.node = node
          last.top = r.top
        }
      }
    }
  }
  return lines
}

/**
 * Convert DOM text-rect lines into safe line-break boundaries.
 *
 * The first rect is the glyph box inside the first line box, so its top can be
 * a few pixels below the block top. Treating it as a boundary creates a phantom
 * first line and lets pagination clip through glyphs. Only subsequent line
 * starts are valid page-break positions.
 *
 * A boundary sits midway between the previous line's ink bottom and the next
 * line's ink top (like cellCutYs), not at the ink top itself: a cut at the ink
 * top leaves zero clearance, and print rasterization (PDF y quantizes to
 * device pixels) slices the next line's glyph tops onto the page bottom
 * (prod100r4 088/099/049). Overlapping ink keeps the old ink-top cut.
 */
export function lineBreakBoundaries(lines: Array<{ offset: number; bottom: number }>): number[] {
  return lines
    .slice(1)
    .map((ln, i) => Math.min(ln.offset, (lines[i].bottom + ln.offset) / 2))
    .filter((off) => off > 0.5)
}

/** DOM anchor of a line start: the line's first text node + character offset within it
 *  (feed to view.posAtDOM), or the line's inline image element on text-less lines. */
export interface LineAnchor {
  node: Text | Element
  charOffset: number
}

/** Element an anchor hangs off (the element itself, or the text node's parent). */
export function anchorElement(a: LineAnchor): Element | null {
  return a.node instanceof Element ? a.node : a.node.parentElement
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

const toAnchor = (ln: { node: Text | Element; top: number }): LineAnchor =>
  ln.node instanceof Element
    ? { node: ln.node, charOffset: 0 }
    : { node: ln.node, charOffset: lineStartCharOffset(ln.node, ln.top) }

/**
 * DOM anchor of the line start matching an in-block virtual Y (offsetInBlock)
 * (used to position mid-paragraph page-break decorations).
 * Returns null when no matching line is found (non-text block / hard pixel cut point).
 */
export function lineStartAnchor(
  el: HTMLElement,
  offsetInBlock: number,
  zoomFactor: number,
  rectsOf: DomLineRectsFn = domLineRects,
): LineAnchor | null {
  for (const ln of rectsOf(el, zoomFactor)) {
    if (Math.abs(ln.offset - offsetInBlock) < 1.5) return toAnchor(ln)
  }
  return null
}

/** DOM anchor of the first line at or after (≥) a given in-block Y: used by in-row cut points (cuts in inter-line gaps) to locate the next page's first line */
export function nextLineAnchor(
  el: HTMLElement,
  offsetInBlock: number,
  zoomFactor: number,
  rectsOf: DomLineRectsFn = domLineRects,
): LineAnchor | null {
  for (const ln of rectsOf(el, zoomFactor)) {
    if (ln.offset >= offsetInBlock - 1.5) return toAnchor(ln)
  }
  return null
}

/** Screen top of a cut anchor's line (the char rect at the anchor, else its parent box) */
function anchorLineTop(a: LineAnchor): number | null {
  if (a.node instanceof Element) return a.node.getBoundingClientRect().top
  if (a.node.length > 0) {
    const range = document.createRange()
    range.setStart(a.node, Math.min(a.charOffset, a.node.length - 1))
    range.setEnd(a.node, Math.min(a.charOffset + 1, a.node.length))
    // jsdom has no Range.getClientRects: fall through to the parent box
    for (const r of range.getClientRects?.() ?? []) if (r.height > 0) return r.top
  }
  return a.node.parentElement?.getBoundingClientRect().top ?? null
}

/** Bottom (screen px) of a cell's content boxes (direct block children, pagination
 *  widgets excluded); -Infinity when the cell has no measurable content.
 *  Text-less blocks holding only spacer-gif struts (≤2px in one dimension, the
 *  HTML-era invisible layout filler) don't count as content. */
function cellContentBottom(cell: Element): number {
  let bottom = -Infinity
  for (const child of Array.from(cell.children)) {
    if (
      child.classList.contains('page-gap') ||
      child.classList.contains('page-gap-cut') ||
      child.classList.contains('page-float-host')
    )
      continue
    if ((child.textContent ?? '').trim() === '' && !child.querySelector('table, svg')) {
      const imgs = Array.from(child.querySelectorAll('img'))
      const isStrut = (im: Element) => {
        const r = im.getBoundingClientRect()
        return r.width <= 2.5 || r.height <= 2.5
      }
      // img-less empty blocks are NOT skipped: they still take real height
      // (an empty every() would be vacuously true) — measure them below
      if (imgs.length > 0 && imgs.every(isStrut)) continue
    }
    const r = child.getBoundingClientRect()
    if (r.height > 0 || r.width > 0) bottom = Math.max(bottom, r.bottom)
  }
  return bottom
}

/** In-row cut decoration policy: a single-cell row can host a real inline gap band
 *  (one cell spans the whole cut) — returns that cell. A multi-cell row can too when
 *  every other cell's content ends above the cut (nothing at the band's y to
 *  misalign — HTML→docx layout tables put whole articles in one cell next to
 *  spacer-gif sliver cells); otherwise the zero-height cut marker stays (same-y
 *  bands across cells are not modeled) — returns null. */
export function singleCutCell(row: Element | null, anchor?: LineAnchor | null): Element | null {
  const cells = row
    ? Array.from(row.children).filter((c) => c.tagName === 'TD' || c.tagName === 'TH')
    : []
  if (cells.length === 1) return cells[0]
  if (cells.length === 0 || !anchor) return null
  const anchorCell = anchorElement(anchor)?.closest('td, th')
  const host = anchorCell && cells.find((c) => c === anchorCell || c.contains(anchorCell))
  if (!host) return null
  const cutTop = anchorLineTop(anchor)
  if (cutTop === null) return null
  for (const c of cells) {
    if (c !== host && cellContentBottom(c) > cutTop + 1) return null
  }
  return host
}
