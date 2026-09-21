// DOM line-box sampling for page-crossing blocks (cached per element), the
// line-split re-slice, table row cut positions and line anchors.
import { rangeSlot } from './dom-range'
import { applyBlockMeta } from './pagination-measure'
import { CapacityWindows, SortedYs } from './pagination-index'
import { blockInlineExtraPx } from './pagination-sections'
import {
  columnLineSplits,
  computeSectionedSlicesF2,
  insertParityBlanks,
  sectionFirstPages,
  type SliceResume,
} from './pagination-slices'
import type {
  BlockBox,
  BlockMetaOf,
  ColWrapRequest,
  ColWrapTable,
  PageSlice,
  RowCellBox,
  RowSplitPatch,
  SectionGeom,
  SliceOutputs,
  TableRowBox,
} from './pagination-types'

const rowRange = rangeSlot()
const lineRange = rangeSlot()
const charRange = rangeSlot()
const anchorRange = rangeSlot()

/**
 * Two-pass slicing: slice by block first, then collect DOM line-box boundaries for
 * blocks crossing page bounds and re-slice. Only blocks that cross a page or exceed
 * one page get line collection (at most one per page, negligible cost).
 * metaOf: docxIndex → parse-layer pagination constraints (keepNext/widow/table row flags).
 */
/**
 * Resume a pass from the first page whose slicing can differ from the previous
 * pass: the pages above it are taken over unchanged and only the blocks from
 * that page on are sliced (Word starts repagination at the changed page too).
 */
export interface PassResume {
  /** the previous pass's slicing before parity blanks (SliceOutputs.preParity) */
  prevSlices: PageSlice[]
  /** index into prevSlices of the first page to slice again */
  page: number
  /** index into blocks of that page's first block (its top is the page start) */
  block: number
  /** the previous pass's outputs: patches for blocks above the page are kept */
  prevOut?: SliceOutputs
}

const PATCH_KEYS = [
  'rowFills',
  'rowSplits',
  'floatVShifts',
  'floatFlows',
  'floatSplits',
  'oversizeClips',
] as const

export function sliceWithLineSplit(
  blocks: BlockBox[],
  geoms: SectionGeom[],
  totalHeight: number,
  zoomFactor: number,
  metaOf?: BlockMetaOf,
  out?: SliceOutputs,
  resume?: PassResume,
): PageSlice[] {
  if (metaOf) applyBlockMeta(blocks, metaOf, zoomFactor)
  const outs: SliceOutputs = out ?? {}
  outs.colWrapRequests ??= []
  const prefix = resume ? resume.prevSlices.slice(0, resume.page) : []
  const run = resume ? blocks.slice(resume.block) : blocks
  const seed = resume ? resumeSeed(resume.prevSlices, resume.page) : undefined
  let slices = prefix.concat(computeSectionedSlicesF2(run, geoms, totalHeight, outs, seed))
  // re-slicing can surface new candidate blocks (a block pushed to a page top only
  // after an earlier block gained line data) — iterate to a fixed point, bounded.
  // The cascade can run one block per page boundary (a dense two-column grid doc
  // packs tighter on every pass: SAS prod_043 left whole paragraphs unsplit at
  // column bottoms with a bound of 3), so the bound follows the page count.
  const maxPasses = Math.max(3, Math.min(24, slices.length))
  outs.iterations = 1
  outs.fillMs = 0
  outs.sliceRunMs = 0
  sigMemo = new Map()
  try {
    for (let i = 0; i < maxPasses; i++) {
      const t0 = performance.now()
      const changed = fillLineBoxes(run, geoms, zoomFactor, slices, metaOf)
      const wrapped = fillColWraps(run, colWrapRequestsOf(run, slices, geoms, outs), zoomFactor)
      outs.fillMs += performance.now() - t0
      if (!changed && !wrapped) break
      const t1 = performance.now()
      slices = prefix.concat(computeSectionedSlicesF2(run, geoms, totalHeight, outs, seed))
      outs.sliceRunMs += performance.now() - t1
      outs.iterations++
    }
  } finally {
    sigMemo = null
  }
  if (resume?.prevOut) keepPatchesAbove(outs, resume.prevOut, resume.prevSlices[resume.page].start)
  outs.preParity = slices
  return insertParityBlanks(slices, geoms)
}

/**
 * Whether a pass can resume at page `page` of a previous slicing: a page of its
 * own (a blank sharing its start with the page before it is re-created by the
 * first block's extra breaks, so resuming there would double it) that opens
 * neither inside a table nor under a lifted block.
 */
export function isResumePage(prev: PageSlice[], page: number): boolean {
  const p = prev[page]
  if (page < 1 || !p || p.end <= p.start + 0.01) return false
  if (prev[page - 1].start >= p.start - 0.5) return false
  return !p.repeatHeader && !p.liftTop
}

function resumeSeed(prev: PageSlice[], page: number): SliceResume {
  const p = prev[page]
  const continued = !!p.continuedSection
  return {
    y: p.start,
    section: p.section,
    firstOfSection: !continued && (page === 0 || prev[page - 1].section !== p.section),
    continued,
  }
}

/** patches of the previous pass for blocks above `y` stay; the resumed run only produced the rest */
function keepPatchesAbove(outs: SliceOutputs, prev: SliceOutputs, y: number): void {
  type Patches = Record<string, Array<{ blockTop: number }> | undefined>
  const o = outs as Patches
  const p = prev as Patches
  for (const key of PATCH_KEYS) {
    const kept = p[key]?.filter((patch) => patch.blockTop < y - 0.5)
    if (kept?.length) o[key] = [...kept, ...(o[key] ?? [])]
  }
}

/** Balance requests plus the tables every line-cut paragraph between unequal
 *  columns needs for its split shape (tail height at the second width). */
function colWrapRequestsOf(
  blocks: BlockBox[],
  slices: PageSlice[],
  geoms: SectionGeom[],
  out: SliceOutputs,
): ColWrapRequest[] {
  const reqs = [...(out.colWrapRequests ?? [])]
  const colWidthsOf = (s: number) => geoms[Math.max(0, Math.min(s, geoms.length - 1))]?.colWidths
  for (const sp of columnLineSplits(blocks, slices, colWidthsOf)) {
    if (sp.tailBottom !== undefined) continue
    reqs.push({
      blockTop: blocks[sp.bi].top,
      headWidthPx: sp.headWidthPx,
      tailWidthPx: sp.tailWidthPx,
    })
  }
  return reqs
}

const colWrapCache = new WeakMap<HTMLElement, { sig: string; tables: ColWrapTable[] }>()

/** Formatting the probe depends on: the paragraph's own text properties and the
 *  markup of its runs (inline sizes, indents); pagination widgets (the split
 *  float, in-block gaps) are layout state, not formatting. Width and height are
 *  deliberately absent: the split itself changes both. */
function colWrapSig(el: HTMLElement, zoomFactor: number): string {
  const cs = getComputedStyle(el)
  const props = [
    cs.fontFamily,
    cs.fontSize,
    cs.lineHeight,
    cs.textIndent,
    cs.textAlign,
    cs.letterSpacing,
    cs.wordSpacing,
    cs.paddingLeft,
    cs.paddingRight,
    cs.paddingTop,
    cs.paddingBottom,
    cs.marginLeft,
    cs.marginRight,
    cs.direction,
  ].join('|')
  let h = 5381
  for (const n of Array.from(el.childNodes)) {
    if (
      n instanceof Element &&
      (n.classList.contains('doc-col-split-float') || n.classList.contains('page-gap'))
    )
      continue
    const text = n instanceof Element ? n.outerHTML : (n.textContent ?? '')
    for (let i = 0; i < text.length; i++) h = ((h << 5) + h + text.charCodeAt(i)) | 0
  }
  return `${lineSampleFontEpoch}:${Math.round(zoomFactor * 1000)}:${props}:${h}`
}

const sameWraps = (t: { headWidthPx: number; tailWidthPx: number }, r: ColWrapRequest) =>
  Math.abs(t.headWidthPx - r.headWidthPx) < 0.5 && Math.abs(t.tailWidthPx - r.tailWidthPx) < 0.5

/**
 * Attach the requested ColWrapTables (probeColWrap, cached per element and text
 * until the text or fonts change). Returns whether any block gained a table (the
 * caller re-slices).
 */
export function fillColWraps(
  blocks: BlockBox[],
  requests: ColWrapRequest[],
  zoomFactor: number,
  probe: typeof probeColWrap = probeColWrap,
): boolean {
  let changed = false
  for (const req of requests) {
    const b = blocks.find((x) => x.el && Math.abs(x.top - req.blockTop) < 0.5)
    if (!b?.el) continue
    if (b.colWraps?.some((t) => sameWraps(t, req))) continue
    const sig = colWrapSig(b.el, zoomFactor)
    let entry = colWrapCache.get(b.el)
    if (!entry || entry.sig !== sig) {
      entry = { sig, tables: [] }
      colWrapCache.set(b.el, entry)
    }
    let table = entry.tables.find((t) => sameWraps(t, req))
    if (!table) {
      const probed = probe(b.el, req.headWidthPx, req.tailWidthPx, zoomFactor)
      if (!probed) continue
      table = probed
      entry.tables.push(table)
    }
    b.colWraps = [...(b.colWraps ?? []), table]
    changed = true
  }
  return changed
}

/** Hidden probe copy of the canvas page: same classes/inline vars as the
 *  ProseMirror root (so the clone inherits every page rule), measuring state,
 *  zero-height overflow box that contains the probe's oversized float. */
function probeHost(pm: HTMLElement): HTMLElement | null {
  const parent = pm.parentElement
  if (!parent) return null
  const wrap = document.createElement('div')
  wrap.className = `${pm.className} measuring-columns doc-col-probe`
  const style = pm.getAttribute('style')
  if (style) wrap.setAttribute('style', style)
  const dir = pm.getAttribute('dir')
  if (dir) wrap.setAttribute('dir', dir)
  wrap.setAttribute('aria-hidden', 'true')
  Object.assign(wrap.style, {
    position: 'absolute',
    left: '0',
    top: '0',
    boxSizing: 'border-box',
    width: `${pm.offsetWidth}px`,
    height: '0',
    overflow: 'hidden',
    visibility: 'hidden',
    pointerEvents: 'none',
    columnCount: 'auto',
  })
  parent.appendChild(wrap)
  return wrap
}

/** The float that rewraps part of a straddling paragraph (same element in the
 *  probe and on the canvas, see setColumnLayout) */
export function splitFloatStyle(
  floatPx: number,
  heightPx: number,
  insetPx: number,
  rtl: boolean | undefined,
): string {
  const r = (n: number) => Math.round(n * 100) / 100
  return `float:${rtl ? 'left' : 'right'};width:${r(floatPx)}px;height:${r(heightPx)}px;shape-outside:inset(${r(insetPx)}px 0 0 0)`
}

const PROBE_FLOAT_H = 1e5

/**
 * Measure a paragraph split between two columns of different widths: a clone at
 * the wider width carries the same leading float shape the canvas will use, so
 * for every head line count k the probe reads the cut Y (lines wrapped at the
 * head width) and the bottom of the remaining lines rewrapped at the tail width.
 * Null when the paragraph has no measurable lines.
 */
export function probeColWrap(
  el: HTMLElement,
  headWidthPx: number,
  tailWidthPx: number,
  zoomFactor: number,
): ColWrapTable | null {
  const pm = el.parentElement
  if (!pm) return null
  const rtl = getComputedStyle(el).direction === 'rtl'
  const wrap = probeHost(pm)
  if (!wrap) return null
  try {
    const clone = el.cloneNode(true) as HTMLElement
    clone.removeAttribute('data-col-patch')
    clone.removeAttribute('data-idx')
    for (const g of Array.from(
      clone.querySelectorAll('.page-gap, .page-gap-inline, .page-float-host, .doc-col-split-float'),
    ))
      g.remove()
    clone.classList.add('doc-col-block')
    const wide = Math.max(headWidthPx, tailWidthPx)
    clone.style.setProperty('--col-w', `${Math.max(0, wide - blockInlineExtraPx(el))}px`)
    const fl = document.createElement('span')
    fl.className = 'doc-col-split-float'
    clone.insertBefore(fl, clone.firstChild)
    wrap.appendChild(clone)
    const cs = getComputedStyle(clone)
    const padTop = (parseFloat(cs.paddingTop) || 0) + (parseFloat(cs.borderTopWidth) || 0)
    const padBottom = (parseFloat(cs.paddingBottom) || 0) + (parseFloat(cs.borderBottomWidth) || 0)
    const contentBottom = () => clone.getBoundingClientRect().height / zoomFactor - padBottom
    const headNarrow = headWidthPx < tailWidthPx
    const floatPx = Math.abs(headWidthPx - tailWidthPx)
    const setFloat = (heightPx: number, insetPx: number) => {
      fl.style.cssText = splitFloatStyle(floatPx, heightPx, insetPx, rtl)
    }
    setFloat(headNarrow ? PROBE_FLOAT_H : 0, 0)
    const headLines = domLineRects(clone, zoomFactor)
    const n = headLines.length
    if (n === 0) return null
    const bounds = lineBreakBoundaries(headLines)
    if (bounds.length !== n - 1) return null
    const headH = [0, ...bounds, contentBottom()]
    // the float edge must not touch the line box on the other side of the cut
    const shapeY = headH.map((_, k) =>
      k === 0 ? 0 : k === n ? headH[n] : headNarrow ? headLines[k - 1].bottom : headLines[k].offset,
    )
    const tailH: number[] = new Array(n + 1).fill(0)
    const tailBottom: number[] = new Array(n + 1).fill(headH[n])
    for (let k = 0; k < n; k++) {
      const at = Math.max(shapeY[k] - padTop, 0)
      if (headNarrow) setFloat(at, 0)
      else setFloat(PROBE_FLOAT_H, at)
      const bottom = contentBottom()
      tailBottom[k] = bottom
      tailH[k] = Math.max(bottom - headH[k], 0)
    }
    return { headWidthPx, tailWidthPx, n, headH, tailH, tailBottom, shapeY }
  } finally {
    wrap.remove()
  }
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

// the DOM does not change inside one slicing call, so its fixed-point iterations
// share the per-element signature reads instead of repeating them
let sigMemo: Map<HTMLElement, string> | null = null

function lineSampleSig(el: HTMLElement, textH: number): string {
  const memoKey = Math.round(textH * 4)
  const memoHit = sigMemo?.get(el)
  if (memoHit && memoHit.startsWith(`${memoKey}:`)) return memoHit.slice(memoHit.indexOf(':') + 1)
  const sig = computeLineSampleSig(el, textH)
  sigMemo?.set(el, `${memoKey}:${sig}`)
  return sig
}

function computeLineSampleSig(el: HTMLElement, textH: number): string {
  // djb2 over the text: equal-length edits must still invalidate
  const text = el.textContent ?? ''
  let h = 5381
  for (let i = 0; i < text.length; i++) h = ((h << 5) + h + text.charCodeAt(i)) | 0
  // width guards width-only reflows; descendant count guards nested (e.g. table
  // cell) structure changes that keep the direct-child count. A stale miss only
  // costs one re-sample, so quantization errs toward invalidating.
  const w = el.getBoundingClientRect().width
  const nodes = el.getElementsByTagName('*').length
  // justify-shrink decorations move wrap points without changing text, height
  // or (when a decoration migrates between lines) the descendant count (r177)
  let js = 0
  for (const sp of el.querySelectorAll<HTMLElement>('.doc-jshrink')) {
    const per = Math.round((parseFloat(sp.style.wordSpacing) || 0) * -100)
    js = (js * 33 + per + (sp.textContent?.length ?? 0)) | 0
  }
  return `${lineSampleFontEpoch}:${Math.round(textH * 4)}:${Math.round(w * 4)}:${nodes}:${h}:${js}`
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
  // column windows of mixed-column pages and titlePg first pages, by start with
  // a running max end: a block top lies in at most the few windows the back
  // scan visits before the max end drops below it
  const windows: Array<{ start: number; end: number; cap: number }> = []
  ;(slices ?? []).forEach((s, i) => {
    if (i > 0) breaks.push(s.start)
    for (const r of s.regions ?? []) {
      for (const c of r.columns) {
        if (c.start > 0.5) breaks.push(c.start)
        windows.push({ start: c.start, end: c.end, cap: r.height })
      }
    }
  })
  const firsts = sectionFirstPages(slices ?? [])
  slices?.forEach((s, si) => {
    if (!firsts[si]) return
    const fc = geomOf(s.section)?.firstContentHeight
    if (fc !== undefined) windows.push({ start: s.start, end: s.end, cap: fc })
  })
  const capWindows = new CapacityWindows(windows)
  const breakYs = new SortedYs(breaks)
  let changed = false
  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i]
    if (!block.el || block.lineBoxes || block.tableRows) continue
    // Anchored textbox/shape blocks are atomic like Word shapes: their inner
    // text lines (usually inside fixed, clipped boxes) are not page-break
    // points. Left line-less, an over-page block places whole and overlaps the
    // bottom margin; the next block turns the page (Word-like shape overflow).
    if (
      block.el.classList.contains('doc-protected-textboxes') ||
      block.el.classList.contains('img-wrap-band')
    )
      continue
    const contentH = geomOf(block.section ?? 0)?.contentHeight ?? 0
    if (contentH <= 0) continue
    const bottom = block.top + block.height
    const crossing = breakYs.hasBetween(block.top, bottom)
    const atPageTop = breakYs.hasNear(block.top, 0.5)
    // region-aware capacity: a block in a mixed-column page's later region has
    // only the region's height, not the full page — a table there must get row
    // data even when its height fits a page (otherwise the first pass places
    // it whole as an "over-page" block and the collapse becomes a fixed point:
    // the ANSI table after a 3-col region). A block on a
    // section's FIRST page likewise has only the titlePg capacity.
    const capH = capWindows.capAt(block.top, contentH)
    // a block already carrying the renderer-baked oversize clip measures at the
    // clipped (fitting) height, so it must bypass the fit gates to re-qualify —
    // otherwise the flag drops, the clip clears, and the layout oscillates
    const clipMarked = block.el.dataset.oversizeClip !== undefined
    if (
      !clipMarked &&
      block.height <= capH &&
      !crossing &&
      !atPageTop &&
      // a mid-paragraph page break splits the block at a line boundary
      !block.innerBreaks?.length &&
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
            if (flags[i]?.keepNext) r.keepNext = true
            if (flags[i]?.minHPx) r.minHPx = flags[i].minHPx
          })
        const bands =
          block.docxIndex !== undefined ? metaOf?.(block.docxIndex)?.footnoteBands : undefined
        if (bands) applyRowNotes(block.el, rows, bands)
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
    // a leading page break's line has no text rect, so its cut (the first text
    // line's ink top) lands inside the first sampled line box: make it a boundary
    for (const y of block.innerBreaks ?? []) {
      if (y > 0.5 && y < textH - 0.5 && (boundaries.length === 0 || y < boundaries[0] - 1.5))
        boundaries.unshift(y)
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
/** the outer table's real rows: nested-table trs are in-row content, and
 *  decoration rows (page gaps / repeated tblHeader clones) are not page-split
 *  units — counting them would add phantom boundaries and shift the
 *  tableRowFlags index alignment */
export function outerTableRows(el: HTMLElement): HTMLElement[] {
  return Array.from(el.querySelectorAll<HTMLElement>('tr')).filter(
    (tr) =>
      !tr.closest('.doc-nested-table') &&
      !tr.classList.contains('page-gap') &&
      !tr.classList.contains('page-repeat-header'),
  )
}

/** Preview rules for Word-style split rows: stamp each split tr (data-pv-split, cloned
 *  with the block) and, on the page each fragment lands on, translate / clip / hide
 *  the cells' children so only that fragment's lines show, at its top, while the
 *  clone stays a plain clip of the canvas flow */
export function rowSplitCss(
  splits: RowSplitPatch[],
  blocks: BlockBox[],
  slices: Array<{ start: number }>,
): string {
  const root = blocks[0]?.el?.parentElement
  for (const el of root?.querySelectorAll('[data-pv-split]') ?? [])
    el.removeAttribute('data-pv-split')
  const pageOf = (y: number) => {
    let pg = 0
    while (pg + 1 < slices.length && slices[pg + 1].start <= y + 0.5) pg++
    return pg
  }
  const px = (v: number) => `${v.toFixed(1)}px`
  const out: string[] = []
  splits.forEach((split, n) => {
    const block = blocks.find((b) => b.tableRows && Math.abs(b.top - split.blockTop) < 0.5)
    const tr = block?.el ? outerTableRows(block.el)[split.row] : undefined
    if (!tr) return
    tr.dataset.pvSplit = String(n)
    for (const r of split.rules) {
      const decl =
        r.dy !== undefined
          ? `transform:translateY(${px(r.dy)})`
          : r.clipTop !== undefined || r.clipBottom !== undefined
            ? `clip-path:inset(${px(r.clipTop ?? 0)} 0 ${px(r.clipBottom ?? 0)} 0)`
            : 'visibility:hidden'
      out.push(
        `.pv-page[data-pv-page="${pageOf(r.from)}"] .pv-content tr[data-pv-split="${n}"] > :nth-child(${r.cell + 1}) > :nth-child(${r.tail ? 'n+' : ''}${r.child + 1}){${decl}}`,
      )
    }
  })
  return out.join('\n')
}

/** Charge each footnote's height to the row holding its reference (bands and
 *  refs are both in document order); a count mismatch keeps the block-level
 *  reservation placed after the table */
export function applyRowNotes(
  el: HTMLElement,
  rows: TableRowBox[],
  bands: Array<{ heightPx: number }>,
): void {
  const refs = Array.from(
    el.querySelectorAll('sup.doc-note-ref[data-note-kind="footnote"]'),
  ).filter((ref) => !ref.closest('.page-gap, .page-float-host, .page-repeat-header'))
  if (refs.length !== bands.length) return
  const trs = outerTableRows(el)
  refs.forEach((ref, i) => {
    const ri = trs.findIndex((tr) => tr.contains(ref))
    if (ri >= 0 && rows[ri]) rows[ri].notesPx = (rows[ri].notesPx ?? 0) + bands[i].heightPx
  })
}

function domTableRows(el: HTMLElement, blockHeight: number, zoomFactor: number): TableRowBox[] {
  const gaps = Array.from(el.querySelectorAll('.page-gap-inline')).map((g) =>
    g.getBoundingClientRect(),
  )
  const gapAbove = (top: number) => gaps.reduce((s, g) => (g.top <= top ? s + g.height : s), 0)
  const elTop = el.getBoundingClientRect().top
  const trs = outerTableRows(el)
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
    const { cuts, contentBottom, cells } = rowCutYs(
      trs[i],
      b.offsetInBlock,
      b.height,
      elTop,
      gapAbove,
      zoomFactor,
    )
    const splitExtra = parseFloat(trs[i].dataset.splitExtra ?? '')
    return {
      height: b.height,
      contentBottom,
      ...(trs[i].hasAttribute('data-repeat-header')
        ? { isHeader: trs[i].getAttribute('data-repeat-header') === '1' }
        : {}),
      ...(cuts.length > 0 ? { cutYs: cuts } : {}),
      ...(cells ? { cells } : {}),
      ...(splitExtra > 0 ? { splitExtra } : {}),
    }
  })
}

const PARA_SELECTOR = 'p, h1, h2, h3, h4, h5, h6, .doc-li'

/** In-row safe cut points (relative to row top, px, ascending): line-level candidates
 *  per cell (Word breaks between any two lines), rejecting cuts that would cross a
 *  line box in another cell. Also reports the lowest content-band bottom and, for
 *  multi-cell rows, each cell's own line geometry (per-cell Word-style splitting). */
function rowCutYs(
  tr: Element,
  rowTop: number,
  rowHeight: number,
  elTop: number,
  gapAbove: (top: number) => number,
  zoomFactor: number,
): { cuts: number[]; contentBottom: number; cells?: RowCellBox[] } {
  const cells = Array.from(tr.children).filter((c) => c.tagName === 'TD' || c.tagName === 'TH')
  const range = rowRange()
  const cellBands: Array<Array<[number, number]>> = []
  const paraBands: Array<Array<[number, number]>> = []
  const cellBoxes: RowCellBox[] = []
  // exact-height clip boxes and vertical text are not line-splittable content
  let cellsOk =
    cells.length >= 2 && !tr.querySelector(':scope > * > .cell-clip, :scope > * > .cell-vert')
  let paraSeq = 0
  for (const cell of cells) {
    const bands: Array<[number, number]> = []
    const byPara = new Map<Element, Array<[number, number]>>()
    const paraIds = new Map<Element, number>()
    const entries: Array<{ band: [number, number]; child: number; para: number }> = []
    const toBand = (r: DOMRect): [number, number] => [
      (r.top - elTop - gapAbove(r.top)) / zoomFactor - rowTop,
      (r.bottom - elTop - gapAbove(r.bottom)) / zoomFactor - rowTop,
    ]
    const childIndexOf = (node: Node | null): number => {
      let e: Node | null = node
      while (e && e.parentNode !== cell) e = e.parentNode
      return e ? Array.prototype.indexOf.call(cell.children, e) : -1
    }
    const paraIdOf = (para: Element | null, child: number) => {
      if (!para) return -1 - child
      let id = paraIds.get(para)
      if (id === undefined) {
        id = paraSeq++
        paraIds.set(para, id)
      }
      return id
    }
    const addEntry = (band: [number, number], node: Node, para: Element | null) => {
      const child = childIndexOf(node)
      if (child < 0) cellsOk = false
      else entries.push({ band, child, para: paraIdOf(para, child) })
    }
    const walker = document.createTreeWalker(cell, NodeFilter.SHOW_TEXT)
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      const parent = n.parentElement
      if (parent?.closest('.page-gap, .page-float-host')) continue
      const para = parent?.closest(PARA_SELECTOR) ?? null
      range.selectNodeContents(n)
      for (const r of range.getClientRects()) {
        if (r.height <= 0 || r.width <= 0) continue
        const band = toBand(r)
        bands.push(band)
        addEntry(band, n, para)
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
      if (r.height > 0 && r.width > 0) {
        const band = toBand(r)
        bands.push(band)
        addEntry(band, obj, obj.closest(PARA_SELECTOR))
      }
    }
    // empty paragraphs above the cell's content are lines Word breaks between;
    // trailing empty marks stay band-less (they drive the over-page tail fill)
    const contentB = bands.reduce((m, [, b]) => Math.max(m, b), -Infinity)
    for (const para of cell.querySelectorAll(PARA_SELECTOR)) {
      if (byPara.has(para) || para.closest('.page-gap, .page-float-host')) continue
      if (para.textContent?.trim() || para.querySelector('img, svg, canvas')) continue
      const r = para.getBoundingClientRect()
      if (r.height <= 0 || r.width <= 0) continue
      const band = toBand(r)
      if (band[0] >= contentB - 0.5) continue
      bands.push(band)
      byPara.set(para, [band])
      addEntry(band, para, para)
    }
    if (bands.length > 0) cellBands.push(bands)
    for (const list of byPara.values()) paraBands.push(list)
    if (cellsOk) cellBoxes.push(cellBoxOf(cell as HTMLElement, entries, zoomFactor, toBand))
  }
  const contentBottom = cellBands.reduce(
    (max, bands) => bands.reduce((m, [, b]) => Math.max(m, b), max),
    0,
  )
  return {
    cuts: cellCutYs(cellBands, rowHeight, paraBands),
    contentBottom,
    ...(cellsOk && cellBoxes.some((c) => c.lines.length > 0) ? { cells: cellBoxes } : {}),
  }
}

/** Per-cell line geometry in top-aligned coordinates: a middle/bottom-aligned cell's
 *  content is measured where the canvas centered it, so the alignment offset is
 *  removed here (Word aligns within each row fragment, not the whole row) */
function cellBoxOf(
  cell: HTMLElement,
  entries: Array<{ band: [number, number]; child: number; para: number }>,
  zoomFactor: number,
  toBand: (r: DOMRect) => [number, number],
): RowCellBox {
  const va = cell.style.verticalAlign
  let alignDy = 0
  let alignFrac = 0
  const first = cell.firstElementChild
  if ((va === 'middle' || va === 'bottom') && first && entries.length > 0) {
    alignFrac = va === 'middle' ? 0.5 : 1
    // rects are zoomed, computed lengths are not
    const cs = getComputedStyle(cell)
    const off =
      (first.getBoundingClientRect().top - cell.getBoundingClientRect().top) / zoomFactor -
      (parseFloat(getComputedStyle(first).marginTop) || 0) -
      (parseFloat(cs.paddingTop) || 0) -
      (parseFloat(cs.borderTopWidth) || 0)
    alignDy = Math.max(0, off)
  }
  const shifted = alignDy
    ? entries.map((e) => ({
        ...e,
        band: [e.band[0] - alignDy, e.band[1] - alignDy] as [number, number],
      }))
    : entries
  const childBox = Array.from(cell.children, (ch): [number, number] => {
    const [t, b] = toBand(ch.getBoundingClientRect())
    return [t - alignDy, b - alignDy]
  })
  return { ...cellLinesOf(shifted), childBox, alignDy, alignFrac }
}

/** Pure core of cellBoxOf: rect entries → clustered lines (same overlap rule as
 *  clusterLineBands), each line keeping the child index / paragraph id of its
 *  topmost rect */
export function cellLinesOf(
  entries: Array<{ band: [number, number]; child: number; para: number }>,
): Pick<RowCellBox, 'lines' | 'childOf' | 'paraOf'> {
  const sorted = [...entries].sort((a, b) => a.band[0] - b.band[0])
  const lines: Array<[number, number]> = []
  const childOf: number[] = []
  const paraOf: number[] = []
  for (const e of sorted) {
    const [top, bottom] = e.band
    const last = lines[lines.length - 1]
    const overlap = last ? last[1] - top : 0
    const minH = last ? Math.min(last[1] - last[0], bottom - top) : 0
    if (last && overlap > 1 && overlap > 0.4 * minH) last[1] = Math.max(last[1], bottom)
    else {
      lines.push([top, bottom])
      childOf.push(e.child)
      paraOf.push(e.para)
    }
  }
  return { lines, childOf, paraOf }
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
  const range = lineRange()
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
    // glyph boxes taller than the line pitch (Batang-class KR faces: 1.45em
    // content area under a 1.3029 line) overlap the next line by a few px; a
    // rect the open line covers less than half of starts a new line
    const overlap = lineBottom - r.top
    if (overlap <= 1 || overlap < r.height / 2) {
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
  const range = charRange()
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
    const range = anchorRange()
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
