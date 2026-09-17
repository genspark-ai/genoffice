// DOM measurement of the continuous-flow render: block boxes, floats, note
// anchors, endnote / float-spill synthetic blocks and block metadata.
import type {
  BlockBox,
  BlockMetaOf,
  FloatBox,
  MeasuredContent,
  PageNoteItem,
  PageSlice,
} from './pagination-types'

/**
 * Collect the editor's top-level block boxes (relative to the content-area top,
 * converted back to 100% zoom). origin is the content-area top's screen Y (page
 * rect.top + top margin × zoom). Page-gap decorations (.page-gap) are not content:
 * they are skipped and subtracted from subsequent block coordinates, yielding
 * "gapless continuous flow" virtual coordinates so slicing is independent of the gaps.
 */
/** anchor offset painted as a wrapper translate (px): display-only, the flow slot is unshifted */
export function anchorShiftPx(el: HTMLElement): number {
  return parseFloat(el.dataset.anchorDy ?? '') || 0
}

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
  // float-carry spacer (split floating table): real flow space that moves the
  // anchor paragraph beside the last portion, so it is not a gap; its height
  // is reported on the block it precedes
  let carryApplied = 0
  for (const el of Array.from(pm.children) as HTMLElement[]) {
    const rect = el.getBoundingClientRect()
    if (el.classList.contains('page-gap') || el.classList.contains('page-float-host')) {
      gapAccum += rect.height
      continue
    }
    if (el.classList.contains('page-float-carry')) {
      carryApplied += rect.height / zoomFactor
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
          ...((box as HTMLElement).dataset.pageRelFrom === 'page' ? { pageRelFromPage: true } : {}),
          ...((box as HTMLElement).dataset.noSpill === '1' ? { noSpill: true } : {}),
        })
      }
    }
    // run-level page-relative pictures re-pin like floating boxes (origin = hosting paragraph)
    for (const img of Array.from(
      el.querySelectorAll<HTMLElement>('.doc-inline-img-anchor > img[data-page-rel-v="1"]'),
    )) {
      const b = img.getBoundingClientRect()
      if (b.height <= 0) continue
      const applied = parseFloat(img.dataset.pageFloatDy ?? '0') || 0
      floats.push({
        el: img,
        top: (b.top - origin - gapAccum) / zoomFactor - applied,
        height: b.height / zoomFactor,
        anchorTop: (rect.top - origin - gapAccum) / zoomFactor,
        pinned: false,
        pageRelV: true,
        ...(img.dataset.pageRelFrom === 'page' ? { pageRelFromPage: true } : {}),
      })
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
    const colBreakEls = Array.from(el.querySelectorAll('.doc-col-br')).filter(
      (b) => !b.closest('td, th, .doc-textbox'),
    )
    const hasColBreak = colBreakEls.length > 0
    // zero-height blocks are skipped, except a break carrier (e.g. a floating
    // textbox whose anchor paragraph holds a page-type w:br) must still be seen
    if (rect.height <= 0 && !hasBreak) continue
    // in-block gaps from mid-paragraph page breaks: subtract from block height and add to the gap accumulator for later blocks
    const innerGap = innerGapHeight(el)
    const top = (rect.top - anchorShiftPx(el) * zoomFactor - origin - gapAccum) / zoomFactor
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
    // breaks with no text before them lead the block: the break line stays on
    // the current page and the block's text starts the next one (Word), so the
    // last leading break cuts like a mid-paragraph break; a field break has no
    // line of its own and maps to breakBefore. Breaks with no text after it
    // trail (breakAfter pushes the next block). Every break character turns the
    // page, so a run of N adjacent breaks leaves N-1 blank sheets. Text is
    // measured outside the break nodes (a field break carries its own label).
    let hasText = false
    let leadingCount = 0
    let trailingCount = 0
    let leadingCut = false
    // a column break with nothing before it moves the whole paragraph (its own
    // line included) to the next column top; a break-only paragraph likewise
    let leadingColBreak = false
    if (colBreakEls.length === 1) {
      const r = document.createRange()
      r.setStart(el, 0)
      r.setEndBefore(colBreakEls[0])
      leadingColBreak = !r.toString().trim()
    }
    // breaks with text on both sides split the paragraph across the page turn
    // (Word keeps the text before on this page); only a trailing break pushes
    // the next block. Y = the post-break line's ink top, element-relative and
    // net of inline gaps a previous pass already inserted there
    const innerBreaks: number[] = []
    if (hasBreak) {
      const gapRects = Array.from(el.querySelectorAll('.page-gap-inline')).map((g) =>
        g.getBoundingClientRect(),
      )
      const gapAbove = (y: number) => gapRects.reduce((s, g) => (g.top <= y ? s + g.height : s), 0)
      const r = document.createRange()
      const brs = breakEls.map((b) => {
        r.setStart(el, 0)
        r.setEndBefore(b)
        const before = r.toString().trim() !== ''
        r.setStartAfter(b)
        r.setEnd(el, el.childNodes.length)
        const after = r.toString().trim() !== ''
        return { b, before, after }
      })
      hasText = brs.some((x) => x.before || x.after)
      leadingCount = brs.filter((x) => !x.before).length
      trailingCount = brs.filter((x) => !x.after).length
      brs.forEach(({ b, before, after }, i) => {
        if (!after || !b.classList.contains('doc-page-br')) return
        if (!before && i !== leadingCount - 1) return
        r.setStartAfter(b)
        r.setEnd(el, el.childNodes.length)
        const first = Array.from(r.getClientRects?.() ?? []).find(
          (c) => c.height > 0 && c.width > 0,
        )
        if (!first) return
        innerBreaks.push((first.top - rect.top - gapAbove(first.top)) / zoomFactor)
        if (!before) leadingCut = true
      })
    }
    const leadingBreak = hasText && leadingCount > 0 && !leadingCut
    const floatTable =
      el.classList.contains('doc-table-float-left') ||
      el.classList.contains('doc-table-float-right')
    const floatFlowed = floatTable && el.classList.contains('doc-table-float-flow')
    const floated =
      /(?:^|\s)img-wrap-(?:square|tight|through)-(?:left|right)(?:\s|$)/.test(el.className) ||
      el.classList.contains('doc-para-frame-float') ||
      (floatTable && !floatFlowed)
    // page/margin-anchored floated table: strip the applied --tblp-dy shift so
    // the engine sees the natural flow position (float margins move only the
    // float's own box, so no gapAccum contribution)
    const relVy = floated ? parseFloat(el.dataset.tblpVy ?? '') : NaN
    const relVAnchor = el.dataset.tblpVanchor
    const relVSpec = el.dataset.tblpVspec
    const relVApplied = Number.isFinite(relVy) ? parseFloat(el.dataset.tblpDy ?? '') || 0 : 0
    const emptyPara = !(el.textContent ?? '').trim() && !el.querySelector('img')
    // non-reflowable blocks keep their rendered width in any column (tables,
    // anchored/inline textbox shapes; protected text paragraphs still reflow)
    const fixedWidth =
      el.tagName === 'TABLE' ||
      el.classList.contains('doc-protected-textboxes') ||
      !!el.querySelector('table')
    const bandKeep = el.dataset.bandKeep === '1' && el.classList.contains('doc-protected-floating')
    const liftPx = parseFloat(el.dataset.tblpLift ?? '')
    blocks.push({
      top: top - relVApplied,
      height,
      ...(floated ? { floated: true } : {}),
      ...(floatTable ? { floatTable: true } : {}),
      ...(floatFlowed ? { floatFlowed: true } : {}),
      ...(liftPx > 0 ? { liftPx } : {}),
      ...(bandKeep ? { bandKeep: true } : {}),
      ...(bandKeep && wrapFloatBottom !== undefined ? { floatBottom: wrapFloatBottom } : {}),
      ...(Number.isFinite(relVy) && (relVAnchor === 'page' || relVAnchor === 'margin')
        ? {
            pageRelVyPx: relVy,
            pageRelVAnchor: relVAnchor,
            ...(relVSpec === 'top' || relVSpec === 'center' || relVSpec === 'bottom'
              ? { pageRelVSpec: relVSpec }
              : {}),
          }
        : {}),
      ...(emptyPara ? { emptyPara: true } : {}),
      ...(fixedWidth
        ? { fixedWidthPx: rect.width / zoomFactor }
        : { widthPx: rect.width / zoomFactor }),
      breakBefore: el.classList.contains('page-break-before') || leadingBreak || undefined,
      breakBeforeBr: leadingBreak || undefined,
      ...(hasText && leadingCount > 1 ? { extraBreaksBefore: leadingCount - 1 } : {}),
      breakAfter: trailingCount > 0 || undefined,
      ...(trailingCount > 1 ? { extraBreaksAfter: trailingCount - 1 } : {}),
      ...(innerBreaks.length > 0 ? { innerBreaks } : {}),
      colBreakBefore: leadingColBreak || undefined,
      colBreakAfter: (hasColBreak && !leadingColBreak) || undefined,
      breakForce: (hasBreak && rect.height <= 0) || undefined,
      el,
      ...(breakOnlyLineH !== undefined ? { breakOnlyLineH } : {}),
      ...(idxAttr ? { docxIndex: parseInt(idxAttr, 10) } : {}),
      ...(carryApplied > 0 ? { carryAppliedPx: carryApplied } : {}),
    })
    carryApplied = 0
    // a CSS float's in-table gaps grow only the float's own box: the blocks
    // after it stack where they would without them
    if (!floated) gapAccum += innerGap
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
    blocks[0].leadFoldPx = lead
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
    if (
      el.classList.contains('page-gap') ||
      el.classList.contains('page-float-host') ||
      el.classList.contains('page-float-carry')
    )
      continue
    const rect = el.getBoundingClientRect()
    if (rect.height <= 0) continue
    return (rect.bottom - baseTop) / factor + (parseFloat(getComputedStyle(el).marginBottom) || 0)
  }
  return null
}

/**
 * Space after of the flow's last text block (px). measureBlocks folds inter-block
 * margins into the preceding block, but nothing follows the last one, while Word
 * still lays a note area out below that paragraph's space after.
 */
export function trailingSpaceAfterPx(blocks: BlockBox[]): number {
  let last: BlockBox | undefined
  for (const b of blocks) {
    if (!b.el || b.height <= 0) continue
    if (!last || b.top + b.height > last.top + last.height) last = b
  }
  if (!last?.el || last.spaceAfterPx) return 0
  return parseFloat(getComputedStyle(last.el).marginBottom) || 0
}

/**
 * Per-page flow-coordinate bottom of the body text (w:footnotePr w:pos="beneathText":
 * Word starts the note area right under the last line instead of at the page bottom).
 * A block's footnote reservation is virtual space below its text and is excluded;
 * virtual blocks (endnote area, float spill) are not text. The flow's last block
 * contributes its space after (the page's window is clamped by the caller).
 */
export function pageTextEnds(blocks: BlockBox[], slices: PageSlice[]): number[] {
  const text = blocks.filter((b) => b.el || b.docxIndex !== undefined)
  const bottomOf = (b: BlockBox) => b.top + b.height - (b.footnoteExtraPx ?? 0)
  const flowEnd = text.reduce((m, b) => Math.max(m, bottomOf(b)), 0)
  const trailing = trailingSpaceAfterPx(blocks)
  return slices.map((s) => {
    let end = s.start
    for (const b of text) {
      const bottom = bottomOf(b)
      if (b.top >= s.end || bottom <= s.start) continue
      end = Math.max(end, Math.min(bottom, s.end))
    }
    if (flowEnd > s.start && flowEnd <= s.end + 0.5) end = Math.max(end, flowEnd + trailing)
    return end
  })
}

/**
 * Page placement of the footnote area (height 0 = none on the page). top is null
 * for Word's default bottom anchor; w:pos="beneathText" puts it right under the
 * last body line (regioned pages keep the bottom anchor, their text end is per
 * column). vOffset is the page's w:vAlign shift, which moves the body and the
 * area alike. Endnote rows on a beneath-text page start below that footnote area
 * (Word stacks footnotes, then endnotes): endnoteShift is added to their top; an
 * area clamped to the content bottom shifts them only as far as its real bottom.
 */
export function noteAreaPlacement(
  height: number,
  textEnd: number | undefined,
  slice: Pick<PageSlice, 'start' | 'regions'>,
  geom: { pageH: number; mTop: number; mBottom: number; headerH: number; vOffset: number },
): { top: number | null; endnoteShift: number } {
  if (height <= 0 || textEnd === undefined || slice.regions) return { top: null, endnoteShift: 0 }
  const beneath = geom.mTop + geom.headerH + geom.vOffset + (textEnd - slice.start)
  const top = Math.min(beneath, geom.pageH - geom.mBottom - height)
  return { top, endnoteShift: top + height - beneath }
}

/**
 * Endnote layout: endnotes gather at the end of the document
 * (or section) right after the body (below its last space after), flowing to later
 * pages when they don't fit.
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
  const top = totalHeight + trailingSpaceAfterPx(blocks)
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
    if (!f.pinned && !f.pageRelV && !f.noSpill) bottom = Math.max(bottom, f.top + f.height)
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

/** Extract each tr's tblHeader/cantSplit/keepNext/atLeast-trHeight flags from table XML (header
 *  repetition across breaks / unsplittable rows / rows kept with the next row / reserved row heights).
 *  `styleKeepNext` resolves a cell paragraph's pStyle to its keepNext (Word probe 2026-09-17: one
 *  keepNext paragraph anywhere in the row chains it to the next row). */
export function tableRowFlags(
  tableXml: string,
  styleKeepNext?: (styleId: string) => boolean,
): Array<{ isHeader: boolean; cantSplit: boolean; keepNext?: boolean; minHPx?: number }> {
  const flags: Array<{
    isHeader: boolean
    cantSplit: boolean
    keepNext?: boolean
    minHPx?: number
  }> = []
  for (const m of tableXml.matchAll(/<w:tr[\s>][\s\S]*?(?=<w:tr[\s>]|<\/w:tbl>)/g)) {
    const trPr = m[0].match(/<w:trPr>[\s\S]*?<\/w:trPr>/)?.[0] ?? ''
    let keepNext = false
    for (const pp of m[0].matchAll(/<w:pPr>([\s\S]*?)<\/w:pPr>/g)) {
      const direct = /<w:keepNext\b[^>]*>/.exec(pp[1])?.[0]
      if (direct) {
        if (!/w:val="(?:0|false)"/.test(direct)) keepNext = true
        continue
      }
      const styleId = /<w:pStyle w:val="([^"]+)"/.exec(pp[1])?.[1]
      if (styleId && styleKeepNext?.(styleId)) keepNext = true
    }
    // non-exact w:trHeight = atLeast (parse.ts semantics); exact rows keep the
    // split path (deliberate clip deviation, see _placeTable). Clamp mirrors
    // parse.ts (MS-OI29500 2.1.51: 31680 twips / 22in).
    const trH = trPr.match(/<w:trHeight\b[^>]*>/)?.[0] ?? ''
    const val = Number(/w:val="(\d+)"/.exec(trH)?.[1])
    const atLeast = Number.isFinite(val) && val > 0 && !/w:hRule="exact"/.test(trH)
    flags.push({
      isHeader: /<w:tblHeader(?!\s+w:val="(?:0|false)")/.test(trPr),
      cantSplit: /<w:cantSplit(?!\s+w:val="(?:0|false)")/.test(trPr),
      ...(keepNext ? { keepNext } : {}),
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
    if (sup.closest('.page-gap, .page-float-host, .page-repeat-header')) continue
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
      // the paragraph's last line; a table's notes all on its first page)
      if (meta.footnoteBands && b.el) {
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
