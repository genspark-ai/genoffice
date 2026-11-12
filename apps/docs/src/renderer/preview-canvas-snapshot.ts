// What the print preview reads from the live canvas, and the clones it builds from it:
// comment / revision anchors, per-block clone geometry, the preview-only CSS derived from
// the pass blocks' DOM markers. The preview component and the export path read the canvas
// through this module only (tools/check-page-cut-readers.mjs).
import type { CommentInfo, SectionInfo } from '@genoffice/docx-engine'
import { decodeEntities } from '@genoffice/docx-engine'
import {
  anchorShiftPx,
  pageRelHang,
  outerTableRows,
  sectionVertical,
  verticalBlockShift,
  seamDecision,
  tableRowTops,
  ROW_SEAM_TOL_PX,
  type BlockBox,
  type FloatBox,
  type PageSlice,
  type SectionGeom,
} from './pagination'
import type { EditorView } from '@tiptap/pm/view'
import {
  anchorPointFor,
  revGroupsOf,
  visiblePointNear,
  type AnchorBlock,
  type AnchorPoint,
} from './editor/margin-annotations'

/** an open comment thread's anchor, in clone flow coordinates; `no` is the
 *  Word print number (document order over all open threads — including
 *  balloon-suppressed ones, which still consume their number) */
export type CommentSpot = { id: string; no: number; top: number; endX: number; endY: number }

/** a tracked revision that lives in a balloon (deletion / format change), in
 *  clone flow coordinates; the leader leaves the revision point (endX, endY) */
export type RevSpot = {
  key: string
  kind: 'del' | 'fmt'
  text: string
  author: string
  date: string | null
  top: number
  endX: number
  endY: number
}

/** one comment-range marker in the parsed block list: owning block, offset in
 *  its XML (document order within the block), and whether it sits inside a
 *  w:tbl (depth-aware — the block may be an SDT-wrapped table or a paragraph
 *  hosting a textbox table) */
export type MarkerSpot = { docxIndex: number; off: number; inTbl: boolean }

/**
 * Range markers per comment id, from one scan over the parsed blocks. `start`
 * prefers w:commentRangeStart (falling back to the bare reference), `end`
 * prefers w:commentRangeEnd, matching where Word aligns the balloon and drops
 * the leader.
 */
export function scanCommentMarkers(
  blocks: AnchorBlock[] | undefined,
): Map<string, { start?: MarkerSpot; end?: MarkerSpot; ref?: MarkerSpot }> {
  const map = new Map<string, { start?: MarkerSpot; end?: MarkerSpot; ref?: MarkerSpot }>()
  if (!blocks) return map
  const re = /<w:tbl[\s>]|<\/w:tbl>|<w:comment(RangeStart|RangeEnd|Reference)\b[^>]*w:id="([^"]+)"/g
  for (const b of blocks) {
    if (b.docxIndex == null || !b.originalXml) continue
    re.lastIndex = 0
    let depth = 0
    let m: RegExpExecArray | null
    while ((m = re.exec(b.originalXml)) !== null) {
      if (m[0].startsWith('</w:tbl')) depth = Math.max(0, depth - 1)
      else if (m[0].startsWith('<w:tbl')) depth += 1
      else {
        const id = m[2]
        const kind = m[1] === 'RangeStart' ? 'start' : m[1] === 'RangeEnd' ? 'end' : 'ref'
        const entry = map.get(id) ?? {}
        if (!entry[kind]) {
          entry[kind] = { docxIndex: b.docxIndex, off: m.index, inTbl: depth > 0 }
          map.set(id, entry)
        }
      }
    }
  }
  return map
}

/**
 * Open comment threads → print spots, measured against the neutralized canvas.
 * In-paragraph ranges anchor to their .doc-comment spans; ranges without a
 * mark (cross-paragraph, image blocks, table cells) fall back to the marker's
 * block. Numbering and stacking follow document order of the range start,
 * like Word — a geometric sort inverts same-line neighbors (the tie would
 * break on the range END), and dropped threads would shift later numbers.
 */
export function measureCommentSpots(
  pm: HTMLElement,
  comments: CommentInfo[],
  blocks: AnchorBlock[] | undefined,
  origin: number,
  factor: number,
  pmContentLeft: number,
): CommentSpot[] {
  // a thread marked across differently-formatted runs spans several
  // .doc-comment elements: the balloon aligns to the first one, the leader
  // leaves the last one (Word connects at the range end)
  const firstSpanOf = new Map<string, HTMLElement>()
  const lastSpanOf = new Map<string, HTMLElement>()
  for (const span of pm.querySelectorAll<HTMLElement>('.doc-comment')) {
    for (const id of (span.dataset.commentIds ?? '').split(' ')) {
      if (!id) continue
      if (!firstSpanOf.has(id)) firstSpanOf.set(id, span)
      lastSpanOf.set(id, span)
    }
  }
  const markers = scanCommentMarkers(blocks)
  const blockElOf = (spot: MarkerSpot | undefined): HTMLElement | null =>
    spot ? pm.querySelector<HTMLElement>(`[data-idx="${String(spot.docxIndex)}"]`) : null
  type Measured = CommentSpot & {
    spanEl: HTMLElement | null
    orderEl: HTMLElement
    /** range-start offset in the order block's XML; +Infinity until a session
     *  thread (no parsed marker) is interpolated among its block's file threads */
    pos: number
    balloon: boolean
  }
  const measured: Measured[] = []
  for (const c of comments) {
    if (c.parentId || c.done) continue
    const mk = markers.get(c.id)
    const startSpot = mk?.start ?? mk?.ref
    const endSpot = mk?.end ?? mk?.ref
    const spanEl = firstSpanOf.get(c.id) ?? null
    let el = spanEl
    let endEl = lastSpanOf.get(c.id) ?? el
    let balloon = true
    if (!el) {
      const startEl = blockElOf(startSpot)
      if (!startEl) continue
      el = startEl
      endEl = blockElOf(endSpot) ?? startEl
      // Word prints no balloon when the range touches table cells, but the
      // thread still consumes its number
      balloon = !(startSpot?.inTbl || endSpot?.inTbl)
    }
    const rects = [...el.getClientRects()].filter((r) => r.height > 0)
    if (rects.length === 0) continue
    const endRects = [...(endEl?.getClientRects() ?? [])].filter((r) => r.height > 0)
    const first = rects[0]
    const last = endRects[endRects.length - 1] ?? rects[rects.length - 1]
    // order key = owning block + range-start offset in its XML (exact document
    // order); session threads have no marker and group under their span's block
    const orderEl =
      blockElOf(startSpot) ?? spanEl?.closest<HTMLElement>('[data-idx]') ?? (el as HTMLElement)
    measured.push({
      id: c.id,
      no: 0,
      top: (first.top - origin) / factor,
      endX: (last.right - pmContentLeft) / factor,
      endY: (last.bottom - origin) / factor - 1,
      spanEl,
      orderEl,
      pos: startSpot?.off ?? Number.POSITIVE_INFINITY,
      balloon,
    })
  }
  const domBefore = (a: HTMLElement, b: HTMLElement) =>
    (a.compareDocumentPosition(b) & Node.DOCUMENT_POSITION_FOLLOWING) !== 0
  // session threads take a position interpolated among the same block's
  // marker-backed threads (span DOM order is range-start order): a pairwise
  // block-vs-span comparison instead would rank the ancestor block first —
  // and mixing keys per pair breaks the sort's transitivity
  for (const m of measured) {
    if (Number.isFinite(m.pos) || !m.spanEl) continue
    let prev = Number.NEGATIVE_INFINITY
    let next = Number.POSITIVE_INFINITY
    let pin = Number.POSITIVE_INFINITY
    // provable bounds: file threads whose own span orders against this one;
    // a shared first span (overlapping ranges on one mark) pins the session
    // range start at that span's document position
    for (const o of measured) {
      if (o === m || o.orderEl !== m.orderEl || !Number.isFinite(o.pos) || !o.spanEl) continue
      if (o.spanEl === m.spanEl) pin = Math.min(pin, o.pos)
      else if (domBefore(o.spanEl, m.spanEl)) prev = Math.max(prev, o.pos)
      else next = Math.min(next, o.pos)
    }
    if (Number.isFinite(pin)) {
      // just after the pinning thread: its loaded Word number stays stable,
      // and markers later in the block XML stay after the session thread
      m.pos = pin + 0.5
      continue
    }
    // spanless fallbacks it cannot be ordered against: keep their Word
    // numbers stable by sorting after them, within the provable window
    for (const o of measured) {
      if (o === m || o.orderEl !== m.orderEl || !Number.isFinite(o.pos) || o.spanEl) continue
      if (o.pos < next) prev = Math.max(prev, o.pos)
    }
    if (Number.isFinite(prev) && Number.isFinite(next)) m.pos = (prev + next) / 2
    else if (Number.isFinite(prev)) m.pos = prev + 0.5
    else if (Number.isFinite(next)) m.pos = next - 0.5
  }
  measured.sort((a, b) => {
    if (a.orderEl !== b.orderEl) return domBefore(a.orderEl, b.orderEl) ? -1 : 1
    if (a.pos !== b.pos) return a.pos - b.pos
    if (a.spanEl && b.spanEl && a.spanEl !== b.spanEl) return domBefore(a.spanEl, b.spanEl) ? -1 : 1
    return a.top - b.top || a.endX - b.endX
  })
  measured.forEach((m, i) => {
    m.no = i + 1
  })
  return measured
    .filter((m) => m.balloon)
    .map(({ id, no, top, endX, endY }) => ({ id, no, top, endX, endY }))
}

/**
 * Balloon revisions (All Markup print view: deletions and format changes have
 * left the text flow) → print spots. Marked text comes from the editor
 * document, anchored where the hidden run sat; protected blocks (TOC entries,
 * field results) carry no marks, so their deleted runs come from the XML,
 * anchored at the block's last line. A collapsed (fully deleted) block has no
 * box and prints no balloon — Word shows nothing for those either.
 */
export function measureRevisionSpots(
  view: EditorView | null,
  pm: HTMLElement,
  blocks: AnchorBlock[] | undefined,
  origin: number,
  factor: number,
  pmContentLeft: number,
  kinds: ReadonlySet<RevSpot['kind']> = new Set(['del', 'fmt']),
): RevSpot[] {
  const spots: RevSpot[] = []
  if (view) {
    for (const g of revGroupsOf(view.state.doc)) {
      if (!kinds.has(g.kind)) continue
      const p = anchorPointFor(view, pm, g.from)
      if (!p) continue
      spots.push({
        key: `m${g.from}`,
        kind: g.kind,
        text: g.text,
        author: g.author,
        date: g.date,
        top: (p.top - origin) / factor,
        endX: (p.left - pmContentLeft) / factor,
        endY: (p.bottom - origin) / factor - 1,
      })
    }
  }
  for (const b of blocks ?? []) {
    if (!kinds.has('del')) break
    if (b.type !== 'passthrough' || b.docxIndex == null || !b.originalXml?.includes('<w:delText'))
      continue
    const el = pm.querySelector<HTMLElement>(`[data-idx="${String(b.docxIndex)}"]`)
    if (!el) continue
    const rects = [...el.getClientRects()].filter((r) => r.height > 0)
    // a collapsed (fully deleted) block anchors where it sat, like a deleted paragraph
    const p: AnchorPoint | null =
      rects.length > 0
        ? {
            top: rects[0].top,
            bottom: rects[rects.length - 1].bottom,
            left: rects[rects.length - 1].right,
          }
        : visiblePointNear(el, pm)
    if (!p) continue
    const text = Array.from(
      b.originalXml.matchAll(/<w:delText(?:\s[^>]*)?>([\s\S]*?)<\/w:delText>|<w:tab\/>/g),
      (m) => (m[1] === undefined ? ' ' : m[1]),
    ).join('')
    const delTag = /<w:del\b[^>]*>/.exec(b.originalXml)?.[0] ?? ''
    const attr = (name: string) =>
      decodeEntities(new RegExp(`w:${name}="([^"]*)"`).exec(delTag)?.[1] ?? '')
    spots.push({
      key: `x${b.docxIndex}`,
      kind: 'del',
      text: decodeEntities(text.trim()),
      author: attr('author'),
      date: attr('date') || null,
      top: (p.top - origin) / factor,
      endX: (p.left - pmContentLeft) / factor,
      endY: (p.bottom - origin) / factor - 1,
    })
  }
  spots.sort((a, b) => a.top - b.top || a.endX - b.endX)
  return spots
}

/** Snapshot of one top-level canvas block for pruned per-page clones (virtual gapless coordinates, layout px) */
export interface CloneChild {
  html: string
  vTop: number
  vBottom: number
  /** CSS margins (layout px): spacer heights must exclude them to keep flow positions exact */
  mt: number
  mb: number
  /** zero-height marker (hidden bookmarks etc.): always kept, never worth pruning */
  zero: boolean
}

/**
 * Per-page full-document clones cost pages × doc DOM; past this budget (top-level
 * blocks × pages) a 300+-page document OOMs the renderer during preview/export
 * ("Promise was collected"), so pages switch to pruned clones: blocks outside the
 * page window collapse into fixed-height spacers.
 */
export const CLONE_PRUNE_BUDGET = 150_000
/**
 * Blocks under-count a document whose blocks are individually huge (page-long
 * tables, heavily run-split paragraphs): a 112-page report of 869 blocks but
 * 87k elements stayed under the block budget, so the preview carried 9.7 M
 * nodes and each chunked printToPDF took minutes. Elements are what Blink
 * pays for, so they get their own budget. 1.7 M nodes (107 pages of 16k
 * elements) still took 23 minutes to open against 3 pruned.
 */
export const CLONE_PRUNE_ELEMENT_BUDGET = 1_000_000
/** window slack around a page (px): keeps neighbours whose floats/overflow bleed into the page */
export const CLONE_PRUNE_PAD = 2000

/** Per-page clones of the whole document are affordable only within both budgets. */
export function shouldPruneClones(pages: number, blocks: number, elements: number): boolean {
  return pages * blocks >= CLONE_PRUNE_BUDGET || pages * elements >= CLONE_PRUNE_ELEMENT_BUDGET
}

/**
 * Canvas block → clone HTML. Phantom table rows (page-gap / repeated-header
 * widgets) are removed and rowspans restored to their source values
 * (data-base-rowspan): the canvas grows rowspans to bridge the phantom rows,
 * but the clone hides/drops them, so the grown spans would swallow real rows.
 */
export function cloneBlockHtml(el: HTMLElement): string {
  if (!el.querySelector('tr.page-gap, tr.page-repeat-header, [data-base-rowspan]')) {
    return el.outerHTML
  }
  const tmp = el.cloneNode(true) as HTMLElement
  for (const tr of Array.from(tmp.querySelectorAll('tr.page-gap, tr.page-repeat-header'))) {
    tr.remove()
  }
  for (const td of Array.from(tmp.querySelectorAll('[data-base-rowspan]'))) {
    td.setAttribute('rowspan', td.getAttribute('data-base-rowspan')!)
  }
  return tmp.outerHTML
}

/**
 * The canvas content box is `--page-w` minus two paddings, each snapped to layout
 * units by Blink; a clone sized from the twips value lands on another layout unit
 * and a line that just fit on the canvas wraps in the clone (every block below
 * then paints a line lower than the engine's cuts assume). The full-width canvas
 * block nearest the nominal width carries the snapped value.
 */
export function canvasWrapWidth(blockWidths: readonly number[], nominal: number): number | null {
  let best: number | null = null
  for (const w of blockWidths) {
    if (Math.abs(w - nominal) > 0.05) continue
    if (best === null || Math.abs(w - nominal) < Math.abs(best - nominal)) best = w
  }
  return best
}

/** a pruned clone's block geometry: the pass's block box when it measured one
 *  (the cuts were sliced on it, so the clone positions against the same
 *  numbers), else the raw rect (zero-height markers, joined paragraph tails) */
export function cloneChildGeom(
  b: BlockBox | undefined,
  raw: { vTop: number; h: number },
): { vTop: number; h: number } {
  if (!b) return raw
  // the lead fold moved the first block's space before out of its top; the
  // applied page-relative shift was taken out of it and the clone paints it
  return {
    vTop: b.top + (b.leadFoldPx ?? 0) + (b.shiftAppliedPx ?? 0),
    h: b.domHeight ?? raw.h,
  }
}

/** pruned clone for one page window: blocks intersecting [from-pad, to+pad] verbatim, pruned runs as spacers */
export function prunedCloneHtml(kids: CloneChild[], from: number, to: number): string {
  const lo = from - CLONE_PRUNE_PAD
  const hi = to + CLONE_PRUNE_PAD
  const parts: string[] = []
  let lastKept: CloneChild | null = null
  let pruned = false
  for (const c of kids) {
    if (!c.zero && (c.vBottom <= lo || c.vTop >= hi)) {
      pruned = true
      continue
    }
    if (pruned) {
      // spacer replaces the pruned run; its height re-derives the next block's
      // border-box top from the previous kept block's margin edge (spacers
      // suppress margin collapse, so both adjacent margins apply in full)
      const base = lastKept ? lastKept.vBottom + lastKept.mb : 0
      const h = Math.max(0, c.vTop - c.mt - base)
      parts.push(
        `<div class="pv-prune-spacer" style="margin:0;border:0;padding:0;height:${h}px"></div>`,
      )
      pruned = false
    }
    parts.push(c.html)
    // zero-height markers anchor positions too: skipping them here made every
    // following marker's spacer re-span the full distance from the last real
    // block, inflating the clone flow (blank pages past the drift)
    lastKept = c
  }
  return parts.join('')
}

/**
 * Page-pinned boxes (cover art) ride the full-document clone onto every page at
 * the same page coordinates; only the copy on the owning page (data-pin-page,
 * stamped on the canvas wrapper before cloning) may stay visible. Expressed as
 * CSS rules because React re-assigns each clone's innerHTML on every re-render
 * (the {__html} wrapper is a fresh object), wiping any imperative DOM fixup.
 * visibility (not display): a stray-run wrapper carries flow height the slices
 * were measured with, and hidden ink must not emit glyphs into the PDF layer.
 */
/** Word keeps a paragraph-relative box on the paper: one whose top would land
 *  above the page edge is pushed down to it (boxOffsets = box top - anchor top) */
export function hoistPaperClamp(dy: number, topPx: number, boxOffsets: number[]): number {
  const top = Math.min(...boxOffsets.map((o) => topPx + dy + o))
  return top < 0 ? dy - top : dy
}

/** anchor-relative boxes the hoist re-pins. A wrapper mixing page-relative V
 *  boxes with anchor-relative ones passes neither the all-page-relative nor the
 *  spill gate and stays inside the pv-clip, which cut every box at the flow end
 *  (a page holding only that anchor line exported blank): un-position it and
 *  hoist the anchor-relative boxes whether or not they spill. */
export function hoistPlan(wfs: FloatBox[]): { boxes: FloatBox[]; mixed: boolean } | null {
  if (wfs.some((f) => f.pinned)) return null
  const boxes = wfs.filter((f) => !f.pageRelV)
  if (boxes.length === 0) return null
  return { boxes, mixed: boxes.length < wfs.length }
}

/** horizontal slot of a floating box's inline position (hoist CSS corrects each slot from column to page) */
export function hoistSlotOf(f: { el: HTMLElement }): 'left' | 'center' | 'right' | null {
  // the float slot style (left/clamp/right) lives on the .box-rotator that
  // owns the box, not on the box itself (same carrier lookup as absBoxes)
  const carrier = f.el.parentElement?.classList.contains('box-rotator') ? f.el.parentElement : f.el
  const st = carrier.getAttribute('style') ?? ''
  // a paper-edge X renders as calc(<px> - var(--doc-margin-left)), a column X
  // clamped onto the paper as clamp(<floor>, <px>, <cap>): both are left slots
  if (/(?:^|;)\s*left:\s*(?:calc\(|clamp\([^,]*,\s*)?-?[\d.]+px/.test(st)) return 'left'
  if (/(?:^|;)\s*left:\s*50%/.test(st)) return 'center'
  if (/(?:^|;)\s*right:\s*0(?:px)?\s*(?:;|$)/.test(st)) return 'right'
  return null
}

/**
 * Vertical-text pages: every window clone carries the canvas' own-page translate
 * (--col-dx/dy), so per-page rules re-place each block against THIS page's
 * start (a paragraph continued from the previous page gets a negative offset and
 * its already-shown lines fall outside the clip) and hide the ride-along copies
 * that belong to other pages. The measured horizontal height stands in for the
 * sideways extent (same line count, same line pitch).
 */
export function verticalPageCss(
  blocks: BlockBox[],
  slices: PageSlice[],
  sections: SectionInfo[],
  geoms: SectionGeom[],
): string {
  const rules: string[] = []
  slices.forEach((slice, i) => {
    const si = Math.max(0, Math.min(slice.section, sections.length - 1))
    const mode = sectionVertical(sections[si]?.settings)
    const g = geoms[Math.min(si, geoms.length - 1)]
    if (!mode || !g || slice.regions) return
    const page = `.pv-page[data-pv-page="${i}"]`
    // the clip is the full body height (lines span it), so the window pad's
    // ride-along neighbours must not show: hide the addressable blocks and
    // reveal those intersecting this page (floated frames and other
    // horizontal blocks included; blocks without an index stay visible).
    // Canvas gap chrome (page / hf / notes gaps, cut marks, float hosts) has
    // no index and would paint inside the sheet: hide it too.
    rules.push(
      `${page} .pv-content > [data-idx],` +
        `${page} .pv-content > :is([class*="page-gap"],.page-float-host){visibility:hidden;}`,
    )
    for (const b of blocks) {
      if (!b.el || b.top >= slice.end - 0.5 || b.top + b.height <= slice.start + 0.5) continue
      const idx = b.el.dataset.idx
      if (idx === undefined) continue
      const sel = `${page} .pv-content > [data-idx="${idx}"]`
      if (!b.el.classList.contains('doc-vert-block')) {
        rules.push(`${sel}{visibility:visible;}`)
        continue
      }
      const { dx, dy } = verticalBlockShift(mode, g.contentHeight, b.top - slice.start, b.height)
      rules.push(
        `${sel}{visibility:visible;--pv-vdx:${dx.toFixed(2)}px;--pv-vdy:${dy.toFixed(2)}px;}`,
      )
    }
  })
  return rules.join('\n')
}

/**
 * Glyph ink may overflow its line box (CJK or bold faces whose content area
 * exceeds a tight line height), so the block that opens the next page paints a
 * few pixels above its own top: inside the window of the page before it, which
 * ends exactly there. Word paints nothing of a page's first line on the page
 * before it. Hide that lead block on the preceding page (visibility keeps the
 * flow height); blocks further down cannot reach the window, straddling
 * blocks must stay visible, and floated / anchored carriers are left alone
 * because their boxes may own the page through the pin rules.
 */
export const EDGE_LEAK_PX = 12

export function leadLeakCss(blocks: BlockBox[], slices: PageSlice[]): string {
  const rules: string[] = []
  const pinCarrier = '[data-pin-page],[data-pv-hoist]'
  const root = blocks[0]?.el?.parentElement
  for (const el of root?.querySelectorAll('[data-pv-lead]') ?? [])
    el.removeAttribute('data-pv-lead')
  let leadRows = 0
  slices.forEach((slice, i) => {
    if (i === slices.length - 1 || slice.regions) return
    for (const b of blocks) {
      if (b.top < slice.end - 0.5 || b.top >= slice.end + EDGE_LEAK_PX) continue
      if (b.floated || b.floatTable || b.liftPx || b.pageRelVyPx !== undefined || b.isFloatSpill)
        continue
      const idx = b.el?.dataset.idx
      if (idx === undefined || b.el!.matches(pinCarrier) || b.el!.querySelector(pinCarrier))
        continue
      rules.push(
        `.pv-page[data-pv-page="${i}"] .pv-content > [data-idx="${idx}"]{visibility:hidden;}`,
      )
    }
    // a table cut at a row seam: the seam allowance keeps the collapsed border
    // whole and with it the next row's glyph tops; hide that row's cell
    // contents (its borders and fills stay for the seam)
    for (const b of blocks) {
      if (
        !b.tableRows ||
        !b.el ||
        b.top >= slice.end - ROW_SEAM_TOL_PX ||
        b.top + b.height <= slice.end + ROW_SEAM_TOL_PX
      )
        continue
      const seam = seamDecision(tableRowTops(b), slice.end)
      const tr = seam && seam.row > 0 ? outerTableRows(b.el)[seam.row] : undefined
      if (!tr) continue
      tr.dataset.pvLead = String(leadRows)
      rules.push(
        `.pv-page[data-pv-page="${i}"] .pv-content tr[data-pv-lead="${leadRows}"] > * > *{visibility:hidden;}`,
      )
      leadRows++
    }
    // a page-anchored table lifted into the next page's top margin starts
    // inside this page's window (a negative margin pulls a float over the
    // flow before it); it belongs to the page whose window opens for it
    for (const b of blocks) {
      const hang = pageRelHang(b)
      if (hang >= 0 || b.top < slice.end - 0.5 || b.top + hang >= slice.end - 0.5) continue
      const next = slices[i + 1]
      if (b.top >= next.end - 0.5 && next.end > next.start) continue
      const idx = b.el?.dataset.idx
      if (idx === undefined) continue
      rules.push(
        `.pv-page[data-pv-page="${i}"] .pv-content > [data-idx="${idx}"]{visibility:hidden;}`,
      )
    }
  })
  // a window opened into the top margin for a hanging table also shows the
  // previous page's blocks ending in that band: hide them (floats included);
  // only lifted tables hang there by design and pinned carriers are owned by
  // the pin rules
  slices.forEach((slice, i) => {
    if (i === 0 || !slice.liftTop || slice.regions) return
    for (const b of blocks) {
      if (b.top + b.height > slice.start + 0.5 || b.top + b.height <= slice.start - slice.liftTop)
        continue
      if (b.liftPx) continue
      const idx = b.el?.dataset.idx
      if (idx === undefined || b.el!.matches(pinCarrier) || b.el!.querySelector(pinCarrier))
        continue
      rules.push(
        `.pv-page[data-pv-page="${i}"] .pv-content > [data-idx="${idx}"]{visibility:hidden;}`,
      )
    }
  })
  return rules.join('\n')
}

/**
 * Height of a page's (or column's) body window (px): the slice span capped at the
 * content height (`full` opens the whole capacity: last page, vertical text), grown by
 * the kept last line's spilled leading so its glyphs are not clipped — unless
 * footnotes own that band.
 */
export function bodyWindowHeight(
  slice: Pick<PageSlice, 'start' | 'end' | 'repeatHeader' | 'leadSpill'>,
  contentH: number,
  opts: { full: boolean; hasNotes: boolean; seamExtend: number },
): number {
  const cap =
    contentH - (slice.repeatHeader?.height ?? 0) + (opts.hasNotes ? 0 : (slice.leadSpill ?? 0))
  if (opts.full) return cap
  return Math.max(0, Math.min(slice.end - slice.start, cap) + opts.seamExtend)
}

/** layoutInCell="0" page-positioned cell box (pinCellBoxesToPage stamped its
 *  page): the preview draws a copy at page coordinates on the owning page and
 *  hides the in-flow clones — inside a positioned table the box could neither
 *  escape the pv-clip window nor resolve against the page box */
export interface CellPageBox {
  page: number
  offX: number
  offY: number
  /** X measures from the paper edge (relativeFrom="page"), else from the column start */
  fromPaperEdge: boolean
  fromPage: boolean
  /** behindDoc: the overlay sinks under the page clone like the in-flow box */
  behind: boolean
  html: string
}

export function collectCellPageBoxes(pm: HTMLElement): CellPageBox[] {
  const out: CellPageBox[] = []
  for (const box of pm.querySelectorAll<HTMLElement>(
    '.doc-cell-boxes-page > .doc-textbox[data-cell-page]',
  )) {
    const page = parseInt(box.dataset.pinPage ?? '', 10)
    if (!Number.isFinite(page)) continue
    out.push({
      page,
      offX: parseFloat(box.dataset.cellPageX ?? '') || 0,
      offY: parseFloat(box.dataset.cellPageY ?? '') || 0,
      fromPaperEdge: box.dataset.pageRelXFrom === 'page',
      fromPage: box.dataset.pageRelFrom === 'page',
      behind: (parseInt(box.style.zIndex, 10) || 0) < 0,
      html: box.outerHTML,
    })
  }
  return out
}

/** run-anchored front/behind picture with a page/margin-relative X, drawn once
 *  per owning page at page coordinates (the column clone would clip it) */
export interface PageRelXPicture {
  page: number
  offX: number
  fromPaperEdge: boolean
  /** content-area Y of the picture top on its page (unclamped; the render pins it to the paper) */
  topY: number
  behind: boolean
  html: string
}

/** the inline transform's leading translate(...) (anchor offset, may nest calc()) */
export function stripLeadingTranslate(transform: string): string {
  const m = /^\s*translate\(/.exec(transform)
  if (!m) return transform
  let depth = 0
  for (let i = m[0].length - 1; i < transform.length; i++) {
    if (transform[i] === '(') depth++
    else if (transform[i] === ')' && --depth === 0) return transform.slice(i + 1).trim()
  }
  return ''
}

export function pageRelXPictureHtml(img: HTMLElement): string {
  const copy = img.cloneNode(false) as HTMLElement
  // the overlay sits at the page position: drop the anchor offset, keep rotation/flip
  const xf = stripLeadingTranslate(copy.style.transform)
  if (xf) copy.style.transform = xf
  else copy.style.removeProperty('transform')
  copy.style.left = '0px'
  copy.style.top = '0px'
  delete copy.dataset.pvOverlaid
  return copy.outerHTML
}

/** hoisted boxes show in the body window owning their anchor only (data-pin-win
 *  stamped at hoist time, data-pv-win on each body clone of a regioned page) */
export function hoistWindowCss(wins: ReadonlyMap<number, ReadonlySet<string>>): string {
  const rules: string[] = []
  for (const [pg, keys] of wins) {
    for (const key of keys) {
      const w = `.pv-page[data-pv-page="${pg}"] .pv-content[data-pv-win]:not([data-pv-win="${key}"]) [data-pv-hoist='1'][data-pin-win="${key}"]`
      rules.push(`${w} > .doc-textbox,${w} > .box-rotator,${w} > .doc-img-wrap{visibility:hidden;}`)
    }
  }
  return rules.join('\n')
}

export function pinnedCloneCss(pageCount: number): string {
  const rules: string[] = []
  for (let i = 0; i < pageCount; i++) {
    rules.push(
      `.pv-page[data-pv-page="${i}"] .doc-protected-pagepinned[data-pin-page]:not([data-pin-page="${i}"]){visibility:hidden;}`,
      // page-relative V boxes: same ride-along duplicates, stamped per box
      `.pv-page[data-pv-page="${i}"] [data-page-rel-v='1'][data-pin-page]:not([data-pin-page="${i}"]){visibility:hidden;}`,
      `.pv-page[data-pv-page="${i}"] [data-anchor-dy][data-pin-page]:not([data-pin-page="${i}"]){visibility:hidden;}`,
      // hoisted spill floats (data-pv-hoist wrappers): boxes escape the
      // pv-clip, so their ride-along copies need the same per-page hiding
      `.pv-page[data-pv-page="${i}"] [data-pv-hoist='1']:not([data-pin-page="${i}"]) > .doc-textbox,` +
        `.pv-page[data-pv-page="${i}"] [data-pv-hoist='1']:not([data-pin-page="${i}"]) > .box-rotator,` +
        `.pv-page[data-pv-page="${i}"] [data-pv-hoist='1']:not([data-pin-page="${i}"]) > .doc-img-wrap{visibility:hidden;}`,
    )
  }
  return rules.join('\n')
}

/** the live editor root the preview snapshots */
export function canvasEditorEl(): HTMLElement | null {
  return document.querySelector('.editor-scroll .ProseMirror') as HTMLElement | null
}

/** viewport Y of the flow's origin: the editor top plus the canvas offset */
export function canvasFlowOrigin(pm: HTMLElement, canvasTop: number, factor: number): number {
  return pm.getBoundingClientRect().top + canvasTop * factor
}

/** vertical extent (flow px) of the shape boxes of an anchor-shifted wrapper painted at paintTop */
export function shapeBoxExtent(
  el: HTMLElement,
  paintTop: number,
  factor: number,
): { top: number; bottom: number } {
  const wrapTop = el.getBoundingClientRect().top
  let top = Infinity
  let bottom = -Infinity
  for (const box of el.querySelectorAll<HTMLElement>(
    // a single shape's box is mirrored by the chrome rotator layer
    ':scope > .doc-textbox, :scope > .box-rotator > .doc-textbox',
  )) {
    const r = box.getBoundingClientRect()
    const boxTop = paintTop + (r.top - wrapTop) / factor
    top = Math.min(top, boxTop)
    bottom = Math.max(bottom, boxTop + r.height / factor)
  }
  return { top, bottom }
}

/** the drawings a float wrapper carries: text boxes, picture wraps, anchored inline images */
export function anchoredBoxesOf(wrap: HTMLElement): HTMLElement[] {
  return Array.from(
    wrap.querySelectorAll<HTMLElement>(
      ':scope > .doc-textbox, :scope > .box-rotator > .doc-textbox, ' +
        ':scope > .doc-img-wrap, :scope .doc-inline-img-anchor > img',
    ),
  )
}

/** an absolutely positioned box with a height (its float slot may sit on the chrome rotator) */
export function isLaidOutAbsoluteBox(b: HTMLElement): boolean {
  const carrier = b.parentElement?.classList.contains('box-rotator') ? b.parentElement : b
  return (
    /position:\s*absolute/.test(carrier.getAttribute('style') ?? '') &&
    b.getBoundingClientRect().height > 0
  )
}

/** wrappers an earlier snapshot marked as hoisted */
export function hoistMarkedEls(pm: HTMLElement): HTMLElement[] {
  return Array.from(pm.querySelectorAll<HTMLElement>('[data-pv-hoist]'))
}

/** viewport X of the editor's content box (the padded .doc-page less its left padding) */
export function canvasContentLeft(pm: HTMLElement): number {
  return pm.getBoundingClientRect().left + (parseFloat(getComputedStyle(pm).paddingLeft) || 0)
}

/** layout widths of the top-level canvas blocks */
export function canvasChildWidths(kidEls: HTMLElement[], factor: number): number[] {
  return kidEls.map((el) => el.getBoundingClientRect().width / factor)
}

/** per-block geometry and HTML of the top-level canvas blocks, for pruned per-page clones */
export function cloneChildrenOf(
  kidEls: HTMLElement[],
  blocks: BlockBox[],
  origin: number,
  factor: number,
): CloneChild[] {
  const metas: CloneChild[] = []
  const blockOf = new Map(blocks.flatMap((b) => (b.el ? [[b.el, b] as const] : [])))
  let gapAccum = 0
  for (const el of kidEls) {
    const rect = el.getBoundingClientRect()
    if (el.classList.contains('page-gap') || el.classList.contains('page-float-host')) {
      gapAccum += rect.height
      continue
    }
    if (el.classList.contains('page-float-carry')) continue
    let innerGap = 0
    for (const g of el.querySelectorAll('.page-gap-inline'))
      innerGap += g.getBoundingClientRect().height
    const cs = window.getComputedStyle(el)
    const { vTop, h } = cloneChildGeom(blockOf.get(el), {
      vTop: (rect.top - anchorShiftPx(el) * factor - origin - gapAccum) / factor,
      h: (rect.height - innerGap) / factor,
    })
    // a CSS float's in-table gaps grow only the float's own box (measureBlocks)
    const cssFloat =
      (el.classList.contains('doc-table-float-left') ||
        el.classList.contains('doc-table-float-right')) &&
      !el.classList.contains('doc-table-float-flow')
    if (!cssFloat) gapAccum += innerGap
    metas.push({
      html: cloneBlockHtml(el),
      vTop,
      vBottom: vTop + h,
      mt: parseFloat(cs.marginTop) || 0,
      mb: parseFloat(cs.marginBottom) || 0,
      zero: rect.height <= 0,
    })
  }
  return metas
}

/** the scroll root of the mounted preview sheets */
export function previewScrollEl(): HTMLElement | null {
  return document.querySelector<HTMLElement>('.pv-scroll')
}

/** probes of the mounted clones for the export dump (__previewDebug): the repeated-header
 *  windows and the clone's drift from the pass's block tops */
export function publishPreviewProbes(blockTops: Array<[string, number]>): void {
  const dbg = (window as unknown as { __previewDebug?: Record<string, unknown> }).__previewDebug
  if (!dbg) return
  const out: Array<Record<string, number | string | undefined>> = []
  document.querySelectorAll<HTMLElement>('.pv-repeat-header').forEach((clip) => {
    const idx = clip.dataset.pvTable
    const table =
      (idx ? clip.querySelector<HTMLElement>(`table[data-idx="${idx}"]`) : null) ??
      clip.querySelector<HTMLElement>('table')
    if (!table) return
    const reserved = parseFloat(clip.dataset.pvReserved ?? '') || 0
    const lift = parseFloat(clip.dataset.pvLift ?? '') || 0
    // page-cut: allow — probe reads of the clone geometry
    const clipRect = clip.getBoundingClientRect()
    // the sheets may be scaled: the clip's known height gives the factor
    const scale = clipRect.height / (reserved + lift || 1) || 1
    const row = table.querySelector<HTMLElement>('tr')
    out.push({
      page: Number(clip.closest<HTMLElement>('.pv-page')?.dataset.pvPage ?? out.length),
      table: idx,
      // page-cut: allow — probe
      offsetPx: (table.getBoundingClientRect().top - clipRect.top) / scale - lift,
      // page-cut: allow — probe
      firstRowPx: row ? row.getBoundingClientRect().height / scale : undefined,
      reservedPx: reserved,
    })
  })
  dbg.repeatHeaderOffset = out
  // where the clone's flow departs from the pass's block tops: the first page's
  // full clone against every block's expected top, reported at each step. Both
  // are taken from the first block: its space before collapses through the
  // clone's box and moves `base` with it, which is no drift
  const content = document.querySelector<HTMLElement>(
    '.pv-page[data-pv-page="0"] .pv-clip:not(.pv-repeat-header) .pv-content',
  )
  const bodyClip = content?.closest<HTMLElement>('.pv-clip')
  if (!content || !bodyClip) return
  const clipH = parseFloat(bodyClip.style.height)
  // page-cut: allow — probe
  const scale = clipH > 0 ? bodyClip.getBoundingClientRect().height / clipH || 1 : 1
  // page-cut: allow — probe
  const base = content.getBoundingClientRect().top
  const expected = new Map(blockTops)
  const steps: Array<Record<string, string | number>> = []
  let prevDelta = 0
  let prev: HTMLElement | null = null
  let maxAbs = 0
  let origin: number | null = null
  for (const child of Array.from(content.children) as HTMLElement[]) {
    const idx = child.dataset.idx ?? ''
    const want = idx ? expected.get(idx) : undefined
    if (want === undefined) continue
    // page-cut: allow — probe
    const raw = (child.getBoundingClientRect().top - base) / scale - want
    origin ??= raw
    const delta = raw - origin
    maxAbs = Math.max(maxAbs, Math.abs(delta))
    if (Math.abs(delta - prevDelta) > 0.5 && steps.length < 24) {
      steps.push({
        idx,
        tag: child.tagName,
        cls: child.className.slice(0, 80),
        prevTag: prev?.tagName ?? '',
        prevCls: prev?.className.slice(0, 80) ?? '',
        prevIdx: prev?.dataset.idx ?? '',
        delta: Math.round(delta * 100) / 100,
        prevDelta: Math.round(prevDelta * 100) / 100,
      })
    }
    prevDelta = delta
    prev = child
  }
  dbg.cloneDrift = {
    maxAbs: Math.round(maxAbs * 100) / 100,
    origin: Math.round((origin ?? 0) * 100) / 100,
    steps,
  }
}

/** whether the print preview is mounted */
export function previewIsOpen(): boolean {
  return !!document.querySelector('.pagination-preview')
}

/** the app shell root (export state classes hang on it) */
export function appRootEl(): Element | null {
  return document.querySelector('.app')
}

/** the mounted preview sheets, in page order */
export function previewPageEls(): HTMLElement[] {
  return [...document.querySelectorAll('.pv-page')] as HTMLElement[]
}

/** the section text width the canvas publishes on the editor root (px) */
export function sectionContentWidthPx(root: HTMLElement): number {
  return Number.parseFloat(getComputedStyle(root).getPropertyValue('--section-content-w'))
}
