import { useEffect, useMemo, useState } from 'react'
import type {
  CommentInfo,
  HeaderFooter,
  HfImage,
  HfPartInfo,
  SectionInfo,
  SectionSettings,
} from '@genoffice/docx-engine'
import {
  appendEndnotesBlock,
  appendFloatSpillBlock,
  assignSections,
  effectiveBottomPx,
  effectiveHfRefs,
  effectiveTopPx,
  formatPageNumber,
  liveSections,
  measureBlocks,
  pageNumbers,
  pinnedFloatPage,
  sectionBidi,
  sectionColGeom,
  sectionFirstPages,
  sectionGeoms,
  sectionPageBox,
  sliceWithLineSplit,
  type BlockBox,
  type BlockMetaOf,
  type FloatBox,
  type PageNoteItem,
  type PageSlice,
  type SectionHfHeights,
} from '../pagination'
import { hfHeaderGeom, FOOTNOTE_SEPARATOR_H } from '../line-metrics'
import type { AnchorBlock } from '../editor/margin-annotations'
import { pageBorderStyleOf } from '../editor/pagination-gaps'
import { toRoman } from '../note-format'
import { useI18n } from '../i18n/locale'
import { HF_WASHOUT_FILTER, hfFloatPagePos, hfReservedHeightPx } from '../editor/hf-dom'
import { HeaderFooterArea } from './HeaderFooterArea'

const twipsToPx = (twips: number) => (twips / 1440) * 96

/**
 * Word's print-markup geometry (measured against Word for Mac PDF output):
 * the sheet is laid out on a virtual page widened by a markup strip, the whole
 * thing is scaled uniformly to fit the paper width and centered vertically,
 * and balloons keep their anchor's unscaled Y. On A4 that comes out to
 * k ≈ 0.745 with the gray strip running from the content's right edge.
 */
const MARKUP_EXTRA_W = 272
const MARKUP_BAND_GUTTER = 8
const BUBBLE_ENTRY = 40
const BUBBLE_RIGHT_PAD = 24
const BUBBLE_STACK_GAP = 8
const BUBBLE_FONT_PX = 12
const BUBBLE_LINE_H = 16
const BUBBLE_PAD_H = 12

/** an open comment thread's anchor, in clone flow coordinates; `no` is the
 *  Word print number (document order over all open threads — including
 *  balloon-suppressed ones, which still consume their number) */
export type CommentSpot = { id: string; no: number; top: number; endX: number; endY: number }

/** one comment-range marker in the parsed block list: owning block, offset in
 *  its XML (document order within the block), and whether it sits inside a
 *  w:tbl (depth-aware — the block may be an SDT-wrapped table or a paragraph
 *  hosting a textbox table) */
type MarkerSpot = { docxIndex: number; off: number; inTbl: boolean }

/**
 * Range markers per comment id, from one scan over the parsed blocks. `start`
 * prefers w:commentRangeStart (falling back to the bare reference), `end`
 * prefers w:commentRangeEnd, matching where Word aligns the balloon and drops
 * the leader.
 */
function scanCommentMarkers(
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

/** width-weighted wrap estimate: bubbles are absolutely stacked before layout */
function estimateBubbleHeight(text: string, innerW: number): number {
  let w = 0
  let lines = 1
  for (const ch of text) {
    const wide = (ch.codePointAt(0) ?? 0) > 0x2e7f
    const cw = ch === '\n' ? Infinity : wide ? BUBBLE_FONT_PX : BUBBLE_FONT_PX * 0.55
    if (w + cw > innerW) {
      lines += 1
      w = ch === '\n' ? 0 : cw
    } else {
      w += cw
    }
  }
  return lines * BUBBLE_LINE_H + BUBBLE_PAD_H
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
const CLONE_PRUNE_BUDGET = 150_000
/** window slack around a page (px): keeps neighbours whose floats/overflow bleed into the page */
const CLONE_PRUNE_PAD = 2000

/**
 * Canvas block → clone HTML. Phantom table rows (page-gap / repeated-header
 * widgets) are removed and rowspans restored to their source values
 * (data-base-rowspan): the canvas grows rowspans to bridge the phantom rows,
 * but the clone hides/drops them, so the grown spans would swallow real rows.
 */
function cloneBlockHtml(el: HTMLElement): string {
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
export function pinnedCloneCss(pageCount: number): string {
  const rules: string[] = []
  for (let i = 0; i < pageCount; i++) {
    rules.push(
      `.pv-page[data-pv-page="${i}"] .doc-protected-pagepinned[data-pin-page]:not([data-pin-page="${i}"]){visibility:hidden;}`,
      // page-relative V boxes: same ride-along duplicates, stamped per box
      `.pv-page[data-pv-page="${i}"] [data-page-rel-v='1'][data-pin-page]:not([data-pin-page="${i}"]){visibility:hidden;}`,
      // hoisted spill floats (data-pv-hoist wrappers): boxes escape the
      // pv-clip, so their ride-along copies need the same per-page hiding
      `.pv-page[data-pv-page="${i}"] [data-pv-hoist='1']:not([data-pin-page="${i}"]) > .doc-textbox,` +
        `.pv-page[data-pv-page="${i}"] [data-pv-hoist='1']:not([data-pin-page="${i}"]) > .doc-img-wrap{visibility:hidden;}`,
    )
  }
  return rules.join('\n')
}

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

/**
 * Pagination preview: a read-only snapshot of real page slicing over the canvas's continuous
 * flow. Each page = a full content clone + overflow clipping + negative-margin offset; the
 * clone is fixed at the canvas content width (line breaks from measurement must not change),
 * and paper size/margins render per each page's section (mixed portrait/landscape across
 * sections is real). Headers/footers render per page by Word variant rules (first page /
 * odd-even), with real page numbers.
 */
export function PaginationPreview({
  section,
  sections,
  delSectBreaks,
  hfParts,
  colFlow,
  colMode,
  hf,
  watermark,
  blockMetaOf,
  pageFootnotesOf,
  endnoteItems,
  sectionHfOverride,
  clearPageGaps,
  comments,
  anchorBlocks,
  onExportPdf,
  onClose,
  suppressEscape,
}: {
  /** Canvas geometry (final section): for the measurement origin / clone width */
  section: SectionSettings
  /** All sections: for per-page paper geometry (empty array = single section per `section`) */
  sections: SectionInfo[]
  /** section-break paragraphs whose mark is a tracked deletion (no break in markup views) */
  delSectBreaks?: Set<number>
  /** rId → header/footer parts (multi-section picks by each section's references) */
  hfParts: Record<string, HfPartInfo>
  /** Canvas column-flow geometry (non-null when the canvas column CSS is active): shared by the measuring state / clone wrap width */
  colFlow: { cols: number; colWidthPx: number; gapPx: number } | null
  /** canvas column mode: 'uniform' = whole-page CSS multicol, 'mixed' = per-block layout decorations */
  colMode: 'none' | 'uniform' | 'mixed'
  hf: HfSet
  watermark: string | null
  /** docxIndex → parse-layer pagination constraints (keepNext/widow/table-row flags) */
  blockMetaOf?: BlockMetaOf
  /** Per-page footnote collection (referencing page → entry list), for page-bottom rendering */
  pageFootnotesOf?: (blocks: BlockBox[], slices: PageSlice[]) => PageNoteItem[][]
  /** Endnote entries (placed together at the document end, take part in slicing, may continue across pages) */
  endnoteItems?: PageNoteItem[]
  /** Multi-section: unsaved per-section header/footer edit overrides (default variant) */
  sectionHfOverride?: (sectionIndex: number, kind: 'header' | 'footer') => HeaderFooter | null
  /**
   * Clears the canvas page-gap decorations before the snapshot measure. In-table
   * gap/repeated-header widgets are extra <tr>s that consume rowspan slots, so a
   * vMerge-heavy table measures with collapsed columns (exploding row heights)
   * while they are present; the canvas rebuilds them on its next debounced
   * remeasure after the snapshot.
   */
  clearPageGaps?: () => void
  /** Comment threads: open threads print as Word-style margin balloons (scaled sheet + markup strip) */
  comments?: CommentInfo[]
  /** parsed blocks: anchors for comment ranges that never produced a text mark (cross-paragraph / image / table ranges) */
  anchorBlocks?: AnchorBlock[]
  onExportPdf: () => void
  onClose: () => void
  /** While true (e.g. the print dialog is stacked on top), Escape must not close the preview */
  suppressEscape?: boolean
}) {
  const { t } = useI18n()
  const [slices, setSlices] = useState<PageSlice[]>([])
  const [pageNotes, setPageNotes] = useState<PageNoteItem[][]>([])
  /** Top Y of the endnote area (virtual coordinates); null = no endnotes */
  const [endnotesTop, setEndnotesTop] = useState<number | null>(null)
  const [html, setHtml] = useState('')
  /** non-null = pruned-clone mode (large documents): per-page windows instead of full clones */
  const [cloneKids, setCloneKids] = useState<CloneChild[] | null>(null)
  /** Live section list: a section whose break block was deleted (unsaved) merges into the next, matching the canvas */
  const [secs, setSecs] = useState<SectionInfo[]>(sections)
  /** open comment threads' anchors (clone flow coordinates), in document order */
  const [commentSpots, setCommentSpots] = useState<CommentSpot[]>([])

  const canvasContentW = twipsToPx(section.pageWidth - section.marginLeft - section.marginRight)
  // canvas content-area top = effective top margin after header push-down (matches --page-pad)
  const canvasMTop = effectiveTopPx(
    section,
    hfReservedHeightPx(
      'header',
      hf.header,
      canvasContentW,
      hf.images?.header,
      hfHeaderGeom(section),
    ),
  )
  /** Settings of the page's section (single-section documents fall back to the canvas geometry) */
  const settingsOf = (slice: PageSlice): SectionSettings =>
    secs[Math.min(slice.section, secs.length - 1)]?.settings ?? section
  /** Clone wrap width = the section's measurement width (columned canvas = column width);
   *  differing-width sections wrap at their own content width (per-block width decorations ride the clone) */
  const wrapWOf = (sectionIdx: number): number =>
    colFlow?.colWidthPx ??
    sectionPageBox(secs[Math.min(sectionIdx, Math.max(secs.length - 1, 0))]?.settings ?? section)
      .contentWidth

  useEffect(() => {
    const pm = document.querySelector('.editor-scroll .ProseMirror') as HTMLElement | null
    if (!pm) return
    clearPageGaps?.()
    // Measure at zoom 1: CSS zoom rounds every box to device pixels, so dividing
    // zoomed rects by the factor drifts from the zoom-1 clones the pages render
    // (a long table accumulates rows of error — repeated headers overprint the
    // first data row and page cuts land mid-row). Neutralizing the canvas zoom
    // for the snapshot makes measurement and clone layout share one geometry.
    const zoomHost = pm.closest<HTMLElement>('.doc-zoom')
    const savedZoom = zoomHost?.style.zoom ?? ''
    if (zoomHost) zoomHost.style.zoom = '1'
    const factor = 1
    // switch the columned canvas to the single-flow measuring state (uniform: CSS columns
    // off, width = column width; mixed: block translates off), matching engine column-flow
    // coordinates. vAlign documents carry the same visual translates on the canvas
    // (vAlignShiftSpecs) and the preview applies its own vOffset, so they must be
    // neutralized here too or the shifted rects double-apply.
    const measureNeutralize =
      colMode !== 'none' ||
      section.vAlign === 'center' ||
      section.vAlign === 'bottom' ||
      sections.some((s) => s.settings.vAlign === 'center' || s.settings.vAlign === 'bottom')
    if (measureNeutralize) pm.classList.add('measuring-columns')
    try {
      const origin = pm.getBoundingClientRect().top + canvasMTop * factor
      const { blocks, totalHeight, floats, sectBreaks } = measureBlocks(pm, origin, factor)
      const live = liveSections(sections, blocks, sectBreaks, delSectBreaks)
      setSecs(live)
      if (live.length > 0) assignSections(blocks, live)
      const withEndnotes = appendEndnotesBlock(
        blocks,
        totalHeight,
        endnoteItems ?? [],
        FOOTNOTE_SEPARATOR_H,
      )
      // floating boxes below the flow end still need pages to land on; bottom-
      // margin overhang stays on the page (same allowance as the canvas)
      const lastSec = live.length > 0 ? live[live.length - 1].settings : section
      const flowWithFloats = appendFloatSpillBlock(
        blocks,
        withEndnotes?.totalHeight ?? totalHeight,
        floats,
        lastSec ? twipsToPx(lastSec.marginBottom) : 0,
      )
      const flowH = flowWithFloats ?? withEndnotes?.totalHeight ?? totalHeight
      setEndnotesTop(withEndnotes?.top ?? null)
      let computed: PageSlice[]
      if (live.length > 0) {
        // each section's default-variant header/footer estimated heights → body push-down (matching the canvas)
        const refs = effectiveHfRefs(live)
        const hfHs: SectionHfHeights[] = live.map((s, i) => {
          const set = s.settings
          const w = twipsToPx(set.pageWidth - set.marginLeft - set.marginRight)
          const pick = (kind: 'header' | 'footer'): HeaderFooter | null => {
            if (i === live.length - 1) return kind === 'header' ? hf.header : hf.footer
            const ov = sectionHfOverride?.(i, kind)
            if (ov) return ov
            const rId = refs[i]?.[kind]?.default
            const part = rId ? hfParts[rId] : undefined
            return part
              ? { text: part.text, pageNumber: part.hasPageNumber, paras: part.paras }
              : null
          }
          const imagesOf = (kind: 'header' | 'footer') => {
            const rId = refs[i]?.[kind]?.default
            const fromPart = rId ? hfParts[rId]?.images : undefined
            if (fromPart?.length) return fromPart
            return i === live.length - 1 ? hf.images?.[kind] : undefined
          }
          // titlePg first-page variant heights: the section's first page renders
          // these strips (hfFor), so its slice capacity must match or the taller
          // variant's push-down clips slice content off the page (prod100r4/43)
          const firstPart = (kind: 'header' | 'footer') => {
            const rId = refs[i]?.[kind]?.first
            const part = rId ? hfParts[rId] : undefined
            return part
              ? { text: part.text, pageNumber: part.hasPageNumber, paras: part.paras }
              : null
          }
          const firstImagesOf = (kind: 'header' | 'footer') => {
            const rId = refs[i]?.[kind]?.first
            return rId ? hfParts[rId]?.images : undefined
          }
          return {
            headerPx: hfReservedHeightPx(
              'header',
              pick('header'),
              w,
              imagesOf('header'),
              hfHeaderGeom(set),
            ),
            footerPx: hfReservedHeightPx('footer', pick('footer'), w, imagesOf('footer')),
            ...(s.titlePg
              ? {
                  firstHeaderPx: hfReservedHeightPx(
                    'header',
                    firstPart('header'),
                    w,
                    firstImagesOf('header'),
                    hfHeaderGeom(set),
                  ),
                  firstFooterPx: hfReservedHeightPx(
                    'footer',
                    firstPart('footer'),
                    w,
                    firstImagesOf('footer'),
                  ),
                }
              : {}),
          }
        })
        const geoms = sectionGeoms(live, hfHs)
        // when the canvas column layout is inactive, measure as full-width single flow; the geometry drops column flow to match
        if (colMode === 'none') for (const g of geoms) if (g.cols) g.cols = undefined
        computed = sliceWithLineSplit(blocks, geoms, flowH, factor, blockMetaOf)
      } else {
        const contentH =
          twipsToPx(section.pageHeight) -
          effectiveTopPx(
            section,
            hfReservedHeightPx(
              'header',
              hf.header,
              canvasContentW,
              hf.images?.header,
              hfHeaderGeom(section),
            ),
          ) -
          effectiveBottomPx(
            section,
            hfReservedHeightPx('footer', hf.footer, canvasContentW, hf.images?.footer),
          )
        // titlePg: the first page renders the first-page header/footer variant
        // (hfFor), so it gets its own capacity
        const firstContentH = hf.titlePg
          ? twipsToPx(section.pageHeight) -
            effectiveTopPx(
              section,
              hfReservedHeightPx(
                'header',
                hf.headerFirst,
                canvasContentW,
                hf.images?.headerFirst,
                hfHeaderGeom(section),
              ),
            ) -
            effectiveBottomPx(
              section,
              hfReservedHeightPx('footer', hf.footerFirst, canvasContentW, hf.images?.footerFirst),
            )
          : undefined
        computed = sliceWithLineSplit(
          blocks,
          [
            {
              contentHeight: contentH,
              forceBreak: false,
              ...(firstContentH !== undefined ? { firstContentHeight: firstContentH } : {}),
              ...(colFlow ? { cols: colFlow.cols } : {}),
            },
          ],
          flowH,
          factor,
          blockMetaOf,
        )
      }
      // stamp each page-pinned box's owning page on its canvas wrapper before
      // cloning so the per-page CSS rules can hide the copies on other pages
      for (const f of floats) {
        if (f.pinned) {
          const wrap = f.el.closest<HTMLElement>('.doc-protected-pagepinned')
          if (wrap) wrap.dataset.pinPage = String(pinnedFloatPage(computed, f.anchorTop))
        } else if (f.pageRelV) {
          // page-relative V boxes ride every clone too (absolute boxes ignore
          // the pv-clip overflow): nearby pages showed visible duplicates —
          // a cover band re-painted over the TOC two pages later (real_run2/69).
          // Stamped per box: sibling boxes of one anchor can own different pages.
          f.el.dataset.pinPage = String(pinnedFloatPage(computed, f.anchorTop))
        }
      }
      // wrappers whose every box is page-relative V get un-positioned in the
      // preview (like pinned covers): the boxes escape the pv-clip — sized to
      // the page's flow content, it cut a page-bottom cover band clean off —
      // and resolve against the page box at their page-relative offsets
      for (const f of floats) {
        if (!f.pageRelV) continue
        const wrap = f.el.closest<HTMLElement>('.doc-protected-floating, .doc-img-float')
        if (!wrap) continue
        const boxes = Array.from(
          wrap.querySelectorAll<HTMLElement>(':scope > .doc-textbox, :scope > .doc-img-wrap'),
        )
        // static siblings (an inline drawing sharing the paragraph) neither
        // need the shared containing block nor ride the page-margin translate:
        // only absolutely positioned siblings gate the un-positioning
        const absBoxes = boxes.filter((b) =>
          /position:\s*absolute/.test(b.getAttribute('style') ?? ''),
        )
        if (absBoxes.length > 0 && absBoxes.every((b) => b.dataset.pageRelV === '1')) {
          wrap.dataset.pvPagerel = '1'
        } else {
          delete wrap.dataset.pvPagerel
        }
      }
      // paragraph-anchored floats spilling past their page's flow window: the
      // pv-clip cuts them on the owning page and the next page's window
      // repaints the spilled part at its top (all windows share one flow
      // clone). Un-position the wrapper so the boxes escape the clip and
      // re-pin to the page box at the anchor's content-area Y; the stamped
      // owning page lets other clones hide the ride-along copies.
      for (const el of pm.querySelectorAll<HTMLElement>('[data-pv-hoist]')) {
        delete el.dataset.pvHoist
        delete el.dataset.pinPage
        el.style.removeProperty('--pv-hoist-dy')
      }
      const wrapFloats = new Map<HTMLElement, FloatBox[]>()
      for (const f of floats) {
        const wrap = f.el.closest<HTMLElement>('.doc-protected-floating, .doc-img-float')
        if (!wrap) continue
        const list = wrapFloats.get(wrap)
        if (list) list.push(f)
        else wrapFloats.set(wrap, [f])
      }
      for (const [wrap, wfs] of wrapFloats) {
        // pinned / page-relative wrappers already position against the page
        if (wrap.dataset.pvPagerel === '1' || wrap.classList.contains('doc-protected-pagepinned'))
          continue
        if (!wfs.every((f) => !f.pinned && !f.pageRelV)) continue
        // the hoist CSS re-pins by inline left/top: boxes on the right/center
        // slots (right:0 / left:50%) resolve against the column and must stay
        if (
          !wfs.every((f) => /(?:^|;)\s*left:\s*-?[\d.]+px/.test(f.el.getAttribute('style') ?? ''))
        )
          continue
        const pg = pinnedFloatPage(computed, wfs[0].anchorTop)
        const slice = computed[pg]
        if (!slice || slice.regions || slice.repeatHeader) continue
        const spills = wfs.some((f) => f.top + f.height > slice.end + 1 || f.top < slice.start - 1)
        if (!spills) continue
        wrap.dataset.pvHoist = '1'
        wrap.dataset.pinPage = String(pg)
        wrap.style.setProperty('--pv-hoist-dy', `${wfs[0].anchorTop - slice.start}px`)
      }
      setSlices(computed)
      setPageNotes(pageFootnotesOf ? pageFootnotesOf(blocks, computed) : [])
      // comment anchors, measured in the same neutralized geometry as the blocks;
      // content-relative X: the pm element is the padded .doc-page, but the
      // preview clones strip that padding and sit inside the sheet's own
      const pmContentLeft =
        pm.getBoundingClientRect().left + (parseFloat(getComputedStyle(pm).paddingLeft) || 0)
      setCommentSpots(
        comments && comments.length > 0
          ? measureCommentSpots(pm, comments, anchorBlocks, origin, factor, pmContentLeft)
          : [],
      )
      // Per-page full clones explode on large documents (pages × doc DOM →
      // renderer OOM / "Promise was collected" during printToPDF). Past the
      // budget, snapshot per-block geometry and render pruned windows instead.
      const kidEls = Array.from(pm.children) as HTMLElement[]
      if (computed.length * kidEls.length >= CLONE_PRUNE_BUDGET) {
        const metas: CloneChild[] = []
        let gapAccum = 0
        for (const el of kidEls) {
          const rect = el.getBoundingClientRect()
          if (el.classList.contains('page-gap') || el.classList.contains('page-float-host')) {
            gapAccum += rect.height
            continue
          }
          let innerGap = 0
          for (const g of el.querySelectorAll('.page-gap-inline'))
            innerGap += g.getBoundingClientRect().height
          const cs = window.getComputedStyle(el)
          const vTop = (rect.top - origin - gapAccum) / factor
          const h = (rect.height - innerGap) / factor
          gapAccum += innerGap
          metas.push({
            html: cloneBlockHtml(el),
            vTop,
            vBottom: vTop + h,
            mt: parseFloat(cs.marginTop) || 0,
            mb: parseFloat(cs.marginBottom) || 0,
            zero: rect.height <= 0,
          })
        }
        setCloneKids(metas)
        setHtml('')
      } else {
        setCloneKids(null)
        setHtml(Array.from(pm.children, (c) => cloneBlockHtml(c as HTMLElement)).join(''))
      }
    } finally {
      if (measureNeutralize) pm.classList.remove('measuring-columns')
      if (zoomHost) zoomHost.style.zoom = savedZoom
    }
    // snapshot: measure once on open; deps intentionally empty
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  useEffect(() => {
    // stopPropagation cannot shield this window-level listener from the print
    // dialog's own Escape handler (same target, same phase), so the dialog
    // suppresses it via prop while it is stacked on top
    if (suppressEscape) return
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.stopPropagation()
        onClose()
      }
    }
    window.addEventListener('keydown', onKey, true)
    return () => window.removeEventListener('keydown', onKey, true)
  }, [onClose, suppressEscape])

  const multiSection = secs.length > 1
  const commentById = useMemo(() => new Map((comments ?? []).map((c) => [c.id, c])), [comments])
  const effRefs = useMemo(() => effectiveHfRefs(secs), [secs])
  // single-section also uses pageNumbers: pgNumType w:start renumbering applies to single-section documents too
  const nums = useMemo(
    () => (secs.length > 0 ? pageNumbers(slices, secs) : slices.map((_, i) => i + 1)),
    [slices, secs],
  )
  const firsts = useMemo(() => sectionFirstPages(slices), [slices])

  // line positions of endnote entries in virtual coordinates (matching appendEndnotesBlock's line boxes)
  const endnoteRows = useMemo(() => {
    if (endnotesTop === null || !endnoteItems || endnoteItems.length === 0) return []
    let off = endnotesTop
    return endnoteItems.map((item, i) => {
      const height = (i === 0 ? FOOTNOTE_SEPARATOR_H : 0) + item.height
      const row = { item, top: off, height, withSeparator: i === 0 }
      off += height
      return row
    })
  }, [endnotesTop, endnoteItems])

  const toHf = (rId: string | undefined): HeaderFooter | null => {
    const part = rId ? hfParts[rId] : undefined
    if (!part) return null
    return {
      text: part.text,
      pageNumber: part.hasPageNumber,
      paras: part.paras.length > 0 ? part.paras : undefined,
    }
  }

  /** Single section: reuse the editing state (unsaved header edits are visible); multi-section: pick parts by each section's references */
  const hfFor = (
    i: number,
  ): {
    header: HeaderFooter | null
    footer: HeaderFooter | null
    headerImages?: HfImage[]
    footerImages?: HfImage[]
  } => {
    const pageNo = nums[i]
    if (!multiSection) {
      if (hf.titlePg && i === 0) {
        return {
          header: hf.headerFirst,
          footer: hf.footerFirst,
          headerImages: hf.images?.headerFirst,
          footerImages: hf.images?.footerFirst,
        }
      }
      if (hf.evenOddHf && pageNo % 2 === 0) {
        return {
          header: hf.headerEven,
          footer: hf.footerEven,
          headerImages: hf.images?.headerEven,
          footerImages: hf.images?.footerEven,
        }
      }
      return {
        header: hf.header,
        footer: hf.footer,
        headerImages: hf.images?.header,
        footerImages: hf.images?.footer,
      }
    }
    const slice = slices[i]
    const sec = secs[Math.min(slice.section, secs.length - 1)]
    const refs = effRefs[Math.min(slice.section, effRefs.length - 1)]
    const variant =
      sec.titlePg && firsts[i] ? 'first' : hf.evenOddHf && pageNo % 2 === 0 ? 'even' : 'default'
    // unsaved per-section header/footer edits take priority over document parts (default variant)
    const ovHeader = variant === 'default' ? sectionHfOverride?.(slice.section, 'header') : null
    const ovFooter = variant === 'default' ? sectionHfOverride?.(slice.section, 'footer') : null
    const headerRId = refs.header[variant]
    const footerRId = refs.footer[variant]
    return {
      header: ovHeader ?? toHf(headerRId),
      footer: ovFooter ?? toHf(footerRId),
      headerImages: headerRId ? hfParts[headerRId]?.images : undefined,
      footerImages: footerRId ? hfParts[footerRId]?.images : undefined,
    }
  }

  return (
    <div className="pagination-preview">
      <div className="pv-toolbar">
        <span className="pv-title">{t('appPaginationPreview')}</span>
        <span className="pv-count">{t('appTotalPagesN', { n: slices.length })}</span>
        <span className="pv-hint">{t('appPvHint')}</span>
        <button className="pv-close" data-tip={t('appPvExportTip')} onClick={onExportPdf}>
          {t('appExportPdf')}
        </button>
        <button className="pv-close" onClick={onClose}>
          {t('appClose')}
        </button>
      </div>
      <style>{pinnedCloneCss(slices.length)}</style>
      {/* aria-hidden: the cloned page stack is a visual print preview; exposing
          its full-document DOM to the accessibility tree overflows Blink's AX
          update queue on long documents and crashes the renderer */}
      <div className="pv-scroll" aria-hidden="true">
        {slices.map((slice, i) => {
          const parts = hfFor(i)
          const s = settingsOf(slice)
          const pageBox = sectionPageBox(s)
          const pageW = pageBox.width
          const pageH = pageBox.height
          const secContentW = pageBox.contentWidth
          // effective margins after this page's variant header/footer push-down (an over-tall header pushes the body down)
          const mTop = effectiveTopPx(
            s,
            hfReservedHeightPx(
              'header',
              parts.header,
              secContentW,
              parts.headerImages,
              hfHeaderGeom(s),
            ),
          )
          const mBottom = effectiveBottomPx(
            s,
            hfReservedHeightPx('footer', parts.footer, secContentW, parts.footerImages),
          )
          const contentH = pageH - mTop - mBottom
          // page vertical alignment (sectPr w:vAlign): content of non-full pages shifts down as a whole
          const usedH = Math.min(slice.end - slice.start, contentH)
          const vSpare = Math.max(0, contentH - usedH)
          const vOffset = s.vAlign === 'center' ? vSpare / 2 : s.vAlign === 'bottom' ? vSpare : 0
          // page numbers display in the owning section's number format (w:pgNumType w:fmt)
          const pageNoText = formatPageNumber(
            nums[i],
            secs[Math.min(slice.section, secs.length - 1)]?.pageNumberFmt,
          )
          // page border (w:pgBorders): drawn per sheet; w:display counts pages within the section
          const pageBorder = pageBorderStyleOf(s)
          const firstOfSection = i === 0 || slices[i - 1].section !== slice.section
          const drawPageBorder =
            pageBorder &&
            !(pageBorder.display === 'firstPage' && !firstOfSection) &&
            !(pageBorder.display === 'notFirstPage' && firstOfSection)
          // Word print-markup: uniform scale to make room for the markup strip,
          // scaled sheet centered vertically (see MARKUP_EXTRA_W)
          const markupOn = commentSpots.length > 0
          const markupK = pageW / (pageW + MARKUP_EXTRA_W)
          const markupOffY = (pageH - pageH * markupK) / 2
          const markupTransform = {
            transform: `translateY(${markupOffY}px) scale(${markupK})`,
            transformOrigin: '0 0',
          } as const
          return (
            <div
              key={i}
              className="pv-page"
              data-pv-page={i}
              // inert: cloned pages carry natively focusable nodes (links,
              // contenteditable cells); keyboard focus must not enter the
              // aria-hidden subtree. Kept off .pv-scroll so it stays scrollable.
              inert
              style={
                {
                  width: pageW,
                  height: pageH,
                  '--pv-page-h': `${pageH}px`,
                  '--page-w': `${pageW}px`,
                  '--page-h': `${pageH}px`,
                  '--section-content-w': `${secContentW}px`,
                  '--header-dist': `${pageBox.headerDist}px`,
                  '--footer-dist': `${pageBox.footerDist}px`,
                  '--pv-mr': `${twipsToPx(s.marginRight)}px`,
                  '--pv-ml': `${twipsToPx(s.marginLeft)}px`,
                  '--pv-mt': `${mTop}px`,
                } as React.CSSProperties
              }
            >
              <div
                className="pv-sheet"
                style={{
                  padding: `${mTop}px ${twipsToPx(s.marginRight)}px ${mBottom}px ${twipsToPx(s.marginLeft)}px`,
                  ...(markupOn ? markupTransform : {}),
                }}
              >
                {watermark && (
                  <div className="page-watermark" aria-hidden="true">
                    {watermark}
                  </div>
                )}
                {drawPageBorder && (
                  <div
                    className="pv-page-border"
                    aria-hidden="true"
                    style={{
                      top: pageBorder.sides.top?.insetPx ?? 0,
                      right: pageBorder.sides.right?.insetPx ?? 0,
                      bottom: pageBorder.sides.bottom?.insetPx ?? 0,
                      left: pageBorder.sides.left?.insetPx ?? 0,
                      borderTop: pageBorder.sides.top?.css,
                      borderRight: pageBorder.sides.right?.css,
                      borderBottom: pageBorder.sides.bottom?.css,
                      borderLeft: pageBorder.sides.left?.css,
                    }}
                  />
                )}
                {(parts.headerImages ?? [])
                  .filter((img) => img.floating)
                  .map((img, k) => {
                    // picture watermark (anchored image in the header): drawn once
                    // per page behind the body (negative z-index; .pv-page isolates)
                    const pos = hfFloatPagePos(img, {
                      pageW,
                      pageH,
                      marginLeft: twipsToPx(s.marginLeft),
                      marginRight: twipsToPx(s.marginRight),
                      marginTop: mTop,
                      marginBottom: mBottom,
                      headerDist: pageBox.headerDist,
                      sectMarginTop: twipsToPx(s.marginTop),
                    })
                    return (
                      <img
                        key={`wm${k}`}
                        className="pv-watermark-img"
                        src={img.dataUrl}
                        alt=""
                        aria-hidden="true"
                        style={{
                          left: pos.x,
                          top: pos.y,
                          transform: `translate(${pos.translateX}%, ${pos.translateY}%)`,
                          ...(img.widthPx ? { width: img.widthPx } : {}),
                          ...(img.heightPx ? { height: img.heightPx } : {}),
                          ...(img.washout ? { filter: HF_WASHOUT_FILTER } : {}),
                        }}
                      />
                    )
                  })}
                {/* image-only parts (logo headers) have a null text value but must still render */}
                {(parts.header || parts.headerImages?.some((img) => !img.floating)) && (
                  <HeaderFooterArea
                    kind="header"
                    value={parts.header ?? { text: '' }}
                    images={parts.headerImages?.filter((img) => !img.floating)}
                    readOnly
                    onCommit={() => {}}
                    pageNo={pageNoText}
                    pageTotal={slices.length}
                  />
                )}
                {slice.repeatHeader && !slice.regions && (
                  // tblHeader repeated headers: a broken table's page first renders a clone of the source table's header rows
                  // (the engine already reserved repeatHeader.height on this page)
                  <div className="pv-clip" style={{ height: slice.repeatHeader.height }}>
                    <div
                      className="pv-offset"
                      style={{ marginTop: -slice.repeatHeader.top, width: wrapWOf(slice.section) }}
                    >
                      <div
                        className="doc-page pv-content"
                        dangerouslySetInnerHTML={{
                          __html: cloneKids
                            ? prunedCloneHtml(
                                cloneKids,
                                slice.repeatHeader.top,
                                slice.repeatHeader.top + slice.repeatHeader.height,
                              )
                            : html,
                        }}
                      />
                    </div>
                  </div>
                )}
                {slice.regions ? (
                  // column flow: regions stack vertically; within a region, columns are narrow-clipped side by side (column-leading repeated headers follow their column)
                  slice.regions.map((region, ri) => {
                    const rSec = secs[Math.min(region.section, secs.length - 1)]
                    const rg = rSec
                      ? sectionColGeom(rSec)
                      : (colFlow ?? { cols: 1, colWidthPx: canvasContentW, gapPx: 0 })
                    const extent =
                      ri + 1 < slice.regions!.length
                        ? slice.regions![ri + 1].top - region.top
                        : undefined
                    const multi = rg.cols > 1
                    const rtl = multi && rSec != null && sectionBidi(rSec)
                    const geo = rg as Partial<{ widths: number[]; gaps: number[] }> & typeof rg
                    // per-column width/gap (w:equalWidth="0" lists differ per column);
                    // gaps ride the columns as margins so unequal spaces work too
                    const widthOf = (ci: number) =>
                      multi ? (geo.widths?.[ci] ?? rg.colWidthPx) : undefined
                    const gapAfter = (ci: number) =>
                      multi && ci < region.columns.length - 1 ? (geo.gaps?.[ci] ?? rg.gapPx) : 0
                    return (
                      <div
                        key={ri}
                        className="pv-region"
                        style={{
                          ...(extent !== undefined ? { height: extent } : {}),
                          // RTL section (w:bidi): columns fill right-to-left
                          ...(rtl ? { flexDirection: 'row-reverse' as const } : {}),
                        }}
                      >
                        {region.columns.map((col, ci) => {
                          // the document's very last column window opens to the region's
                          // full capacity: slice bounds can drift a few lines short of the
                          // clone's real height (same allowance as the single-flow branch),
                          // silently cutting the document tail mid-line from export/print
                          const tailWindow =
                            i === slices.length - 1 &&
                            ri === slice.regions!.length - 1 &&
                            ci === region.columns.length - 1 &&
                            vOffset <= 0.5
                          return (
                            <div
                              key={ci}
                              className="pv-col"
                              style={{
                                width: widthOf(ci),
                                ...(rtl
                                  ? { marginLeft: gapAfter(ci) }
                                  : { marginRight: gapAfter(ci) }),
                              }}
                            >
                              {col.repeatHeader && (
                                <div
                                  className="pv-clip"
                                  style={{ height: col.repeatHeader.height }}
                                >
                                  <div
                                    className="pv-offset"
                                    style={{
                                      marginTop: -col.repeatHeader.top,
                                      width: wrapWOf(region.section),
                                    }}
                                  >
                                    <div
                                      className="doc-page pv-content"
                                      dangerouslySetInnerHTML={{
                                        __html: cloneKids
                                          ? prunedCloneHtml(
                                              cloneKids,
                                              col.repeatHeader.top,
                                              col.repeatHeader.top + col.repeatHeader.height,
                                            )
                                          : html,
                                      }}
                                    />
                                  </div>
                                </div>
                              )}
                              <div
                                className="pv-clip"
                                style={{
                                  height: tailWindow
                                    ? region.height - (col.repeatHeader?.height ?? 0)
                                    : Math.min(
                                        col.end - col.start,
                                        region.height - (col.repeatHeader?.height ?? 0),
                                      ),
                                }}
                              >
                                <div
                                  className="pv-offset"
                                  style={{ marginTop: -col.start, width: wrapWOf(region.section) }}
                                >
                                  <div
                                    className="doc-page pv-content"
                                    dangerouslySetInnerHTML={{
                                      __html: cloneKids
                                        ? prunedCloneHtml(
                                            cloneKids,
                                            col.start,
                                            tailWindow ? col.start + region.height : col.end,
                                          )
                                        : html,
                                    }}
                                  />
                                </div>
                              </div>
                            </div>
                          )
                        })}
                      </div>
                    )
                  })
                ) : (
                  <div
                    className="pv-clip"
                    style={{
                      // last page opens to full capacity: slice bounds can drift a few
                      // lines short of the clone's real height (page-crossing tables),
                      // silently dropping the document tail from export/print;
                      // past the real content bottom the window is empty anyway
                      height:
                        i === slices.length - 1 && vOffset <= 0.5
                          ? contentH - (slice.repeatHeader?.height ?? 0)
                          : Math.min(
                              slice.end - slice.start,
                              contentH - (slice.repeatHeader?.height ?? 0),
                            ),
                      ...(vOffset > 0.5 ? { marginTop: vOffset } : {}),
                    }}
                  >
                    {/* the offset lives on a separate wrapper: print rules zero out .doc-page's margin;
                      width is fixed to the section's wrap width so the clone never reflows against the paper */}
                    <div
                      className="pv-offset"
                      style={{ marginTop: -slice.start, width: wrapWOf(slice.section) }}
                    >
                      <div
                        className="doc-page pv-content"
                        dangerouslySetInnerHTML={{
                          __html: cloneKids
                            ? prunedCloneHtml(
                                cloneKids,
                                slice.start,
                                // last page opens its clip to full capacity; the window must cover it
                                i === slices.length - 1 ? slice.start + contentH : slice.end,
                              )
                            : html,
                        }}
                      />
                    </div>
                  </div>
                )}
                {(pageNotes[i]?.length ?? 0) > 0 && (
                  // page-bottom footnotes (Word behavior: placed at the bottom of the page's content area, separator on top)
                  <div
                    className="pv-footnotes"
                    style={{
                      left: twipsToPx(s.marginLeft),
                      width: pageW - twipsToPx(s.marginLeft) - twipsToPx(s.marginRight),
                      bottom: twipsToPx(s.marginBottom),
                      height:
                        pageNotes[i]!.reduce((sum, n) => sum + n.height, 0) + FOOTNOTE_SEPARATOR_H,
                    }}
                  >
                    {pageNotes[i]!.map((n) => (
                      // entries get reserved heights (DOM-measured at these exact styles);
                      // min-height lets a residual long entry spill into the bottom margin
                      // instead of overprinting the next entry
                      <div
                        key={n.id}
                        className="pv-footnote"
                        style={{
                          minHeight: n.height,
                          ...(n.lineHeightPx ? { lineHeight: `${n.lineHeightPx}px` } : {}),
                          ...(n.fontSizePt ? { fontSize: `${n.fontSizePt}pt` } : {}),
                        }}
                      >
                        {!n.noRefMark && <sup>{n.no}</sup>}
                        {n.richParas
                          ? n.richParas.map((para, pi) => (
                              <span key={pi}>
                                {pi > 0 && <br />}
                                {para.map((run, ri) => (
                                  <span
                                    key={ri}
                                    style={{
                                      fontWeight: run.bold ? 600 : undefined,
                                      fontStyle: run.italic ? 'italic' : undefined,
                                      textDecoration:
                                        [run.underline && 'underline', run.strike && 'line-through']
                                          .filter(Boolean)
                                          .join(' ') || undefined,
                                      color: run.color ? `#${run.color}` : undefined,
                                      fontSize: run.sizeHalfPoints
                                        ? `${run.sizeHalfPoints / 2}pt`
                                        : undefined,
                                      textTransform: run.caps === 'all' ? 'uppercase' : undefined,
                                      fontVariantCaps:
                                        run.caps === 'small' ? 'small-caps' : undefined,
                                    }}
                                  >
                                    {run.text}
                                  </span>
                                ))}
                              </span>
                            ))
                          : n.text}
                      </div>
                    ))}
                  </div>
                )}
                {(() => {
                  // endnotes: immediately after the body's end, placed on pages per the slices, may continue across pages
                  const rows = endnoteRows.filter(
                    (r) => r.top >= slice.start - 0.5 && r.top < slice.end - 0.5,
                  )
                  if (rows.length === 0) return null
                  return (
                    <div
                      className={`pv-endnotes${rows[0].withSeparator ? ' with-separator' : ''}`}
                      style={{
                        left: twipsToPx(s.marginLeft),
                        width: pageW - twipsToPx(s.marginLeft) - twipsToPx(s.marginRight),
                        top: mTop + (slice.repeatHeader?.height ?? 0) + (rows[0].top - slice.start),
                      }}
                    >
                      {rows.map(({ item: n, height, withSeparator }) => (
                        <div
                          key={n.id}
                          className="pv-footnote"
                          style={{
                            minHeight: height,
                            ...(n.lineHeightPx ? { lineHeight: `${n.lineHeightPx}px` } : {}),
                            ...(n.fontSizePt ? { fontSize: `${n.fontSizePt}pt` } : {}),
                          }}
                        >
                          {withSeparator && <div className="pv-endnote-separator" />}
                          {!n.noRefMark && <sup>{toRoman(n.no)}</sup>}
                          {n.richParas
                            ? n.richParas.map((para, pi) => (
                                <span key={pi}>
                                  {pi > 0 && <br />}
                                  {para.map((run, ri) => (
                                    <span
                                      key={ri}
                                      style={{
                                        fontWeight: run.bold ? 600 : undefined,
                                        fontStyle: run.italic ? 'italic' : undefined,
                                        textDecoration:
                                          [
                                            run.underline && 'underline',
                                            run.strike && 'line-through',
                                          ]
                                            .filter(Boolean)
                                            .join(' ') || undefined,
                                        color: run.color ? `#${run.color}` : undefined,
                                        fontSize: run.sizeHalfPoints
                                          ? `${run.sizeHalfPoints / 2}pt`
                                          : undefined,
                                        textTransform: run.caps === 'all' ? 'uppercase' : undefined,
                                        fontVariantCaps:
                                          run.caps === 'small' ? 'small-caps' : undefined,
                                      }}
                                    >
                                      {run.text}
                                    </span>
                                  ))}
                                </span>
                              ))
                            : n.text}
                        </div>
                      ))}
                    </div>
                  )
                })()}
                {(parts.footer || parts.footerImages?.some((img) => !img.floating)) && (
                  <HeaderFooterArea
                    kind="footer"
                    value={parts.footer ?? { text: '' }}
                    images={parts.footerImages?.filter((img) => !img.floating)}
                    readOnly
                    onCommit={() => {}}
                    pageNo={pageNoText}
                    pageTotal={slices.length}
                  />
                )}
                <div className="pv-pageno">{i + 1}</div>
              </div>
              {markupOn &&
                (() => {
                  const contentRight = pageW - twipsToPx(s.marginRight)
                  const bubbleLeft = contentRight + BUBBLE_ENTRY
                  const bubbleW = pageW + MARKUP_EXTRA_W - BUBBLE_RIGHT_PAD - bubbleLeft
                  const headerH = slice.repeatHeader?.height ?? 0
                  // the last page opens its clip to full capacity (slice bounds can
                  // land short); its tail anchors must still balloon
                  const sliceEnd =
                    i === slices.length - 1 && vOffset <= 0.5
                      ? Math.max(slice.end, slice.start + contentH - headerH)
                      : slice.end
                  const inPageTopOf = (sp: CommentSpot): number | null => {
                    if (!slice.regions) {
                      if (sp.top < slice.start - 0.5 || sp.top >= sliceEnd - 0.5) return null
                      return mTop + vOffset + headerH + (sp.top - slice.start)
                    }
                    const base = slice.regions[0]?.top ?? slice.start
                    for (const region of slice.regions) {
                      for (const col of region.columns) {
                        // the document's last column window opens to the region's full
                        // capacity (see the pv-clip tail allowance): match its anchors too
                        const colEnd =
                          i === slices.length - 1 &&
                          region === slice.regions.at(-1) &&
                          col === region.columns.at(-1) &&
                          vOffset <= 0.5
                            ? Math.max(
                                col.end,
                                col.start + region.height - (col.repeatHeader?.height ?? 0),
                              )
                            : col.end
                        if (sp.top >= col.start - 0.5 && sp.top < colEnd - 0.5)
                          return (
                            mTop +
                            (region.top - base) +
                            (col.repeatHeader?.height ?? 0) +
                            (sp.top - col.start)
                          )
                      }
                    }
                    return null
                  }
                  const placed: {
                    sp: CommentSpot
                    no: number
                    top: number
                    anchorTop: number
                    text: string
                  }[] = []
                  let prevBottom = -Infinity
                  for (const sp of commentSpots) {
                    const anchorTop = inPageTopOf(sp)
                    if (anchorTop === null) continue
                    const root = commentById.get(sp.id)
                    const replies = (comments ?? []).filter((r) => r.parentId === sp.id)
                    const text = [root?.text ?? '', ...replies.map((r) => `${r.author}: ${r.text}`)]
                      .filter(Boolean)
                      .join('\n')
                    const top = Math.max(anchorTop, prevBottom + BUBBLE_STACK_GAP)
                    // the label prefix wraps with the text, so it counts toward the height
                    const label = `${t('appPvCommented')} [${sp.no}]: `
                    prevBottom = top + estimateBubbleHeight(label + text, bubbleW - 16)
                    placed.push({
                      sp,
                      no: sp.no,
                      top,
                      anchorTop,
                      text,
                    })
                  }
                  return (
                    <div
                      className="pv-markup"
                      style={{
                        width: pageW + MARKUP_EXTRA_W,
                        height: pageH,
                        ...markupTransform,
                      }}
                    >
                      <div
                        className="pv-markup-band"
                        style={{
                          left: contentRight + MARKUP_BAND_GUTTER,
                          width: pageW + MARKUP_EXTRA_W - contentRight - MARKUP_BAND_GUTTER * 2,
                        }}
                      />
                      {placed.length > 0 && (
                        <svg
                          className="pv-comment-leaders"
                          width={pageW + MARKUP_EXTRA_W}
                          height={pageH}
                        >
                          {placed.map((p) => {
                            // leader runs from the range end; a range that ends on
                            // another page falls back to the anchor's first line
                            const endsHere =
                              !slice.regions &&
                              p.sp.endY >= slice.start - 0.5 &&
                              p.sp.endY < sliceEnd + 0.5
                            const ex = endsHere ? twipsToPx(s.marginLeft) + p.sp.endX : contentRight
                            const ey = endsHere
                              ? mTop + vOffset + headerH + (p.sp.endY - slice.start)
                              : p.anchorTop + BUBBLE_LINE_H - 4
                            return (
                              <path
                                key={p.sp.id}
                                d={`M ${ex} ${ey} L ${contentRight + BUBBLE_ENTRY / 2} ${ey} L ${bubbleLeft} ${p.top + 10}`}
                              />
                            )
                          })}
                        </svg>
                      )}
                      {placed.map((p) => (
                        <div
                          key={p.sp.id}
                          className="pv-comment-bubble"
                          style={{ left: bubbleLeft, top: p.top, width: bubbleW }}
                        >
                          <span className="pv-comment-label">
                            {t('appPvCommented')} [{p.no}]:{' '}
                          </span>
                          {p.text}
                        </div>
                      ))}
                    </div>
                  )
                })()}
            </div>
          )
        })}
      </div>
    </div>
  )
}
