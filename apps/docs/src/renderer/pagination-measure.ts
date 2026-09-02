// DOM measurement of the continuous-flow render: block boxes, floats, note
// anchors, endnote / float-spill synthetic blocks and block metadata.
import type {
  BlockBox,
  BlockMetaOf,
  FloatBox,
  MeasuredContent,
  PageNoteItem,
} from './pagination-types'

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
  const floats: FloatBox[] = []
  const sectBreaks = new Set<number>()
  let totalHeight = 0
  let gapAccum = 0
  for (const el of Array.from(pm.children) as HTMLElement[]) {
    const rect = el.getBoundingClientRect()
    if (el.classList.contains('page-gap') || el.classList.contains('page-float-host')) {
      gapAccum += rect.height
      continue
    }
    // floating-anchor boxes: absolute children of a zero-height wrapper; record
    // shift-neutral virtual positions so pages can be extended to contain them
    let wrapFloatBottom: number | undefined
    if (el.classList.contains('doc-protected-floating') || el.classList.contains('doc-img-float')) {
      const anchorTop = (rect.top - origin - gapAccum) / zoomFactor
      const pinned = el.classList.contains('doc-protected-pagepinned')
      for (const box of Array.from(
        el.querySelectorAll(':scope > .doc-textbox, :scope > .doc-img-wrap'),
      )) {
        const b = (box as HTMLElement).getBoundingClientRect()
        if (b.height <= 0) continue
        const applied = parseFloat((box as HTMLElement).dataset.pageFloatDy ?? '0') || 0
        const top = (b.top - origin - (pinned ? 0 : gapAccum)) / zoomFactor - applied
        const height = b.height / zoomFactor
        if (!pinned) wrapFloatBottom = Math.max(wrapFloatBottom ?? 0, top + height)
        floats.push({
          el: box as HTMLElement,
          // pinned boxes position against the page box: gaps never move them
          top,
          height,
          anchorTop,
          pinned,
          pageRelV: (box as HTMLElement).dataset.pageRelV === '1',
        })
      }
    }
    // sectPr-only paragraph: the section-break mark itself has no height in Word
    // (its editor chip must not occupy a page or hold a forced break's page open),
    // but its boundary must stay visible to liveSections or the section merges away
    if (el.classList.contains('doc-protected-sectbreak')) {
      const idx = el.getAttribute('data-idx')
      if (idx) sectBreaks.add(parseInt(idx, 10))
      continue
    }
    // Word ignores page-type w:br inside table cells, and breaks inside a
    // textbox lay out that box's own text — neither may break the body flow
    const breakEls = Array.from(el.querySelectorAll('.doc-field-pagebreak, .doc-page-br')).filter(
      (b) => !b.closest('td, th, .doc-textbox'),
    )
    const hasBreak = breakEls.length > 0
    const hasColBreak = Array.from(el.querySelectorAll('.doc-col-br')).some(
      (b) => !b.closest('td, th, .doc-textbox'),
    )
    // zero-height blocks are skipped, except a break carrier (e.g. a floating
    // textbox whose anchor paragraph holds a page-type w:br) must still be seen
    if (rect.height <= 0 && !hasBreak) continue
    // in-block gaps from mid-paragraph page breaks: subtract from block height and add to the gap accumulator for later blocks
    const innerGap = innerGapHeight(el)
    const top = (rect.top - origin - gapAccum) / zoomFactor
    const height = (rect.height - innerGap) / zoomFactor
    const idxAttr = el.getAttribute('data-idx')
    // break-only paragraph (br line + ProseMirror trailing-break phantom line): marked
    // for dedicated placement — Word pushes it into a deliberate blank page when its
    // line doesn't fit at the page bottom. Word renders a single break line, but the
    // DOM height spans one line box per <br> (a text-less paragraph lays out exactly
    // brCount line boxes), so the fit height is one line's share. Word only charges
    // the line's natural single-spacing extent at the page bottom (probe 20260901: a
    // double-spaced Calibri 11pt break line absorbs at 14pt remaining, while an exact
    // line demands its full exact height), so auto multiples above 1 are divided out.
    const breakOnly = hasBreak && !(el.textContent ?? '').trim() && !el.querySelector('img')
    const brLines = breakOnly ? el.querySelectorAll('br').length : 0
    let breakOnlyLineH: number | undefined
    if (breakOnly) {
      const box = brLines > 1 ? height / brLines : height
      // per-paragraph declarations live in the inline style; the document-level
      // multiple cascades through the computed style. Fixed-height lines demand
      // their full box: direct exact/atLeast carries the doc-lh-fixed class, a
      // style-level exact/atLeast is marked by --doc-line-fixed (doc-style-css)
      // — unless a direct auto override re-declares the inline multiple.
      const inlineMult = el.style.getPropertyValue('--doc-line-mult')
      const fixed =
        el.classList.contains('doc-lh-fixed') ||
        (!inlineMult &&
          (
            el.style.getPropertyValue('--doc-line-fixed') ||
            getComputedStyle(el).getPropertyValue('--doc-line-fixed')
          ).trim() === '1')
      const mult = fixed
        ? 1
        : parseFloat(inlineMult || getComputedStyle(el).getPropertyValue('--doc-line-mult')) || 1
      breakOnlyLineH = box / Math.max(1, mult)
    }
    // a single break with no text before it: Word starts this block's own content
    // on a new page, so it maps to breakBefore (breakAfter only pushes the next block)
    let leadingBreak = false
    if (breakEls.length === 1 && (el.textContent ?? '').trim()) {
      const r = document.createRange()
      r.setStart(el, 0)
      r.setEndBefore(breakEls[0])
      leadingBreak = !r.toString().trim()
    }
    const floated =
      /(?:^|\s)img-wrap-(?:square|tight|through)-(?:left|right)(?:\s|$)/.test(el.className) ||
      el.classList.contains('doc-table-float-left') ||
      el.classList.contains('doc-table-float-right')
    // page/margin-anchored floated table: strip the applied --tblp-dy shift so
    // the engine sees the natural flow position (float margins move only the
    // float's own box, so no gapAccum contribution)
    const relVy = floated ? parseFloat(el.dataset.tblpVy ?? '') : NaN
    const relVAnchor = el.dataset.tblpVanchor
    const relVApplied = Number.isFinite(relVy) ? parseFloat(el.dataset.tblpDy ?? '') || 0 : 0
    const emptyPara = !(el.textContent ?? '').trim() && !el.querySelector('img')
    // non-reflowable blocks keep their rendered width in any column (tables,
    // anchored/inline textbox shapes; protected text paragraphs still reflow)
    const fixedWidth =
      el.tagName === 'TABLE' ||
      el.classList.contains('doc-protected-textboxes') ||
      !!el.querySelector('table')
    const bandKeep = el.dataset.bandKeep === '1' && el.classList.contains('doc-protected-floating')
    blocks.push({
      top: top - relVApplied,
      height,
      ...(floated ? { floated: true } : {}),
      ...(bandKeep ? { bandKeep: true } : {}),
      ...(bandKeep && wrapFloatBottom !== undefined ? { floatBottom: wrapFloatBottom } : {}),
      ...(Number.isFinite(relVy) && (relVAnchor === 'page' || relVAnchor === 'margin')
        ? { pageRelVyPx: relVy, pageRelVAnchor: relVAnchor }
        : {}),
      ...(emptyPara ? { emptyPara: true } : {}),
      ...(fixedWidth ? { fixedWidthPx: rect.width / zoomFactor } : {}),
      breakBefore: el.classList.contains('page-break-before') || leadingBreak || undefined,
      breakBeforeBr: leadingBreak || undefined,
      breakAfter: (hasBreak && !leadingBreak) || undefined,
      colBreakAfter: hasColBreak || undefined,
      breakForce: (hasBreak && rect.height <= 0) || undefined,
      el,
      ...(breakOnlyLineH !== undefined ? { breakOnlyLineH } : {}),
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
  // leading offset before the first block (first-paragraph space-before): Word
  // consumes page capacity with it, so fold it in like the inter-block margins
  // (space-before semantics: counted before the block's own lines)
  const first = blocks[0]
  const firstIsTable =
    !!first?.el && (first.el.matches('table') || !!first.el.querySelector('table'))
  if (blocks.length > 0 && first.top > 0.5 && !firstIsTable) {
    const lead = blocks[0].top
    blocks[0].spaceBeforePx = (blocks[0].spaceBeforePx ?? 0) + lead
    blocks[0].height += lead
    blocks[0].top = 0
  }
  return { blocks, totalHeight, floats, sectBreaks }
}

/**
 * Canvas anchor for the endnote area: display-state bottom (layout px, relative to
 * baseTop) of the last visible in-flow block. Word places endnotes right after the
 * last body line, but the canvas paper is padded to a full page, so the area cannot
 * just stack after the editor — it is absolutely positioned at this Y instead.
 */
export function endnotesAnchorY(pm: HTMLElement, baseTop: number, factor: number): number | null {
  for (let i = pm.children.length - 1; i >= 0; i--) {
    const el = pm.children[i] as HTMLElement
    if (el.classList.contains('page-gap') || el.classList.contains('page-float-host')) continue
    const rect = el.getBoundingClientRect()
    if (rect.height <= 0) continue
    return (rect.bottom - baseTop) / factor
  }
  return null
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
 * Floating boxes extending past the flow end (Word: an anchored object that
 * does not fit on its page moves to the next page) need pages to exist there:
 * append a virtual zero-content block spanning to the lowest float bottom so
 * the slicer materializes the trailing page(s). Fine-grained line boxes let it
 * split at any page boundary without widow constraints.
 */
export function appendFloatSpillBlock(
  blocks: BlockBox[],
  totalHeight: number,
  floats: FloatBox[],
  /** allowed overhang into the landing page's bottom margin (px): Word draws
   *  anchored boxes over the margin instead of opening a page for them */
  bottomOverhangPx = 0,
): number | null {
  let bottom = 0
  // page-absolute boxes (pinned / page-relative V) draw on their anchor's
  // page — Word never opens a page for them, and their measured tops are not
  // flow extents (pinned = page coords, pageRelV = anchor + page offset)
  for (const f of floats) {
    if (!f.pinned && !f.pageRelV) bottom = Math.max(bottom, f.top + f.height)
  }
  bottom -= bottomOverhangPx
  if (bottom <= totalHeight + 1) return null
  const top = totalHeight
  const spill = bottom - totalHeight
  const STEP = 24
  const lineBoxes: Array<{ offsetInBlock: number; height: number }> = []
  for (let off = 0; off < spill; off += STEP) {
    lineBoxes.push({ offsetInBlock: off, height: Math.min(STEP, spill - off) })
  }
  blocks.push({
    top,
    height: spill,
    lineBoxes,
    widowControl: false,
    isFloatSpill: true,
    ...(blocks.length > 0 && blocks[blocks.length - 1].section !== undefined
      ? { section: blocks[blocks.length - 1].section }
      : {}),
  })
  return top + spill
}

/** Extract each tr's tblHeader/cantSplit/atLeast-trHeight flags from table XML (header repetition across breaks / unsplittable rows / reserved row heights) */
export function tableRowFlags(
  tableXml: string,
): Array<{ isHeader: boolean; cantSplit: boolean; minHPx?: number }> {
  const flags: Array<{ isHeader: boolean; cantSplit: boolean; minHPx?: number }> = []
  for (const m of tableXml.matchAll(/<w:tr[\s>][\s\S]*?(?=<w:tr[\s>]|<\/w:tbl>)/g)) {
    const trPr = m[0].match(/<w:trPr>[\s\S]*?<\/w:trPr>/)?.[0] ?? ''
    // non-exact w:trHeight = atLeast (parse.ts semantics); exact rows keep the
    // split path (deliberate clip deviation, see _placeTable). Clamp mirrors
    // parse.ts (MS-OI29500 2.1.51: 31680 twips / 22in).
    const trH = trPr.match(/<w:trHeight\b[^>]*>/)?.[0] ?? ''
    const val = Number(/w:val="(\d+)"/.exec(trH)?.[1])
    const atLeast = Number.isFinite(val) && val > 0 && !/w:hRule="exact"/.test(trH)
    flags.push({
      isHeader: /<w:tblHeader(?!\s+w:val="(?:0|false)")/.test(trPr),
      cantSplit: /<w:cantSplit(?!\s+w:val="(?:0|false)")/.test(trPr),
      ...(atLeast ? { minHPx: Math.min(val, 31680) / 15 } : {}),
    })
  }
  return flags
}

/** Extract each tr's tblHeader flag from table XML (header repeated at page top after a table break) */
export function tableHeaderFlags(tableXml: string): boolean[] {
  return tableRowFlags(tableXml).map((f) => f.isHeader)
}

/** In-block virtual offsets of the block's footnote markers (document order),
 *  same space as lineBoxes offsets (in-block inline gaps subtracted). */
export function noteRefOffsets(el: HTMLElement, zoomFactor: number): number[] {
  const gaps = Array.from(el.querySelectorAll('.page-gap-inline')).map((g) =>
    g.getBoundingClientRect(),
  )
  const gapAbove = (top: number) => gaps.reduce((s, g) => (g.top <= top ? s + g.height : s), 0)
  const elTop = el.getBoundingClientRect().top
  const out: number[] = []
  for (const sup of Array.from(
    el.querySelectorAll('sup.doc-note-ref[data-note-kind="footnote"]'),
  )) {
    if (sup.closest('.page-gap, .page-float-host')) continue
    const r = sup.getBoundingClientRect()
    out.push((r.top - elTop - gapAbove(r.top)) / zoomFactor)
  }
  return out
}

/** Inject parse-layer constraints into measured blocks (call before slicing; table row flags are applied by fillLineBoxes) */
export function applyBlockMeta(blocks: BlockBox[], metaOf: BlockMetaOf, zoomFactor = 1): void {
  for (const b of blocks) {
    if (b.docxIndex === undefined) continue
    const meta = metaOf(b.docxIndex)
    if (!meta) continue
    if (meta.keepNext) b.keepNext = true
    if (meta.modernTableHeaders) b.modernTableHeaders = true
    if (meta.keepLines) b.keepLines = true
    if (meta.breakBefore) b.breakBefore = true
    if (meta.widowControl === false) b.widowControl = false
    if (meta.footnoteExtraPx) {
      // the reservation consumes page capacity through the block height only;
      // it must never ride spaceAfterPx (the page-bottom trailing-space
      // exemption would hand the note area back to body text — prod100r4/038
      // printed three footnotes over the last body lines)
      b.height += meta.footnoteExtraPx
      b.footnoteExtraPx = (b.footnoteExtraPx ?? 0) + meta.footnoteExtraPx
      // marker offsets pair with the per-ref heights (both in document order):
      // a count mismatch keeps the block-level fallback (whole reservation on
      // the paragraph's last line)
      if (meta.footnoteBands && b.el && !b.tableRows && !b.el.querySelector('tr')) {
        const offsets = noteRefOffsets(b.el, zoomFactor)
        if (offsets.length === meta.footnoteBands.length) {
          b.noteBands = meta.footnoteBands.map((band, i) => ({
            offset: offsets[i],
            height: band.heightPx,
          }))
        }
      }
    }
  }
}

/** Total height of in-block inline gaps (mid-paragraph page-break decorations) (screen px) */
function innerGapHeight(el: HTMLElement): number {
  let sum = 0
  for (const g of el.querySelectorAll('.page-gap-inline')) sum += g.getBoundingClientRect().height
  return sum
}
