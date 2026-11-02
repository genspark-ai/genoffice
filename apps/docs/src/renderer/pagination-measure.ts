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
import { effectiveParaFlags } from './editor/para-flags'
import type { BlockPlan, PlanOf } from './pagination-plan'
import { floatKeepWholeOf, wrapperObstacleFacts } from './editor/obstacle-plan'

/** the break facts of one block, from the plan or scraped from its DOM */
interface BreakFacts {
  breakOnly: boolean
  brLines: number
  hasText: boolean
  leadingCount: number
  trailingCount: number
  leadingCut: boolean
  leadingColBreak: boolean
  innerBreaks: number[]
  hasCellBreak: boolean
}

/** the engine flags a scrape would have produced (BlockBox.shadowScrape) */
function scrapeFlags(
  f: BreakFacts,
  pageBreakBefore: boolean,
  hasColBreak: boolean,
  zeroHeight: boolean,
) {
  const leadingBreak = f.hasText && f.leadingCount > 0 && !f.leadingCut
  return {
    breakBefore: pageBreakBefore || leadingBreak || undefined,
    breakBeforeBr: leadingBreak || undefined,
    ...(f.hasText && f.leadingCount > 1 ? { extraBreaksBefore: f.leadingCount - 1 } : {}),
    breakAfter: f.trailingCount > 0 || undefined,
    ...(f.trailingCount > 1 ? { extraBreaksAfter: f.trailingCount - 1 } : {}),
    innerBreakCount: f.innerBreaks.length,
    ...(f.leadingCut ? { leadingBreakCut: true as const } : {}),
    ...(f.hasCellBreak ? { hasCellBreak: true as const } : {}),
    colBreakBefore: f.leadingColBreak || undefined,
    colBreakAfter: (hasColBreak && !f.leadingColBreak) || undefined,
    breakForce: (f.brLines + f.leadingCount + f.trailingCount > 0 && zeroHeight) || undefined,
    breakOnly: f.breakOnly,
  }
}
import { additiveParagraphSpacing, autoLineMultOf } from './line-metrics'
import { appliedFloatShift } from './editor/layout-health'

/**
 * Collect the editor's top-level block boxes (relative to the content-area top,
 * converted back to 100% zoom). origin is the content-area top's screen Y (page
 * rect.top + top margin × zoom). Page-gap decorations (.page-gap) are not content:
 * they are skipped and subtracted from subsequent block coordinates, yielding
 * "gapless continuous flow" virtual coordinates so slicing is independent of the gaps.
 */
/** display-only offsets of a block from its flow slot (px): the anchor wrapper
 *  translate and the leading shift's relative top (styles.css --doc-lead-shift) */
export function anchorShiftPx(el: HTMLElement): number {
  const anchor = parseFloat(el.dataset.anchorDy ?? '') || 0
  return anchor + (parseFloat(getComputedStyle(el).top) || 0)
}

/**
 * Extra leading of one line of an auto-multiple paragraph (px, unzoomed like
 * every computed length here). Word charges only the single-spacing extent at
 * the page bottom (probe 2026-09-23); exact and atLeast lines demand their full
 * box, space before/after always counts.
 */
export function lineLeadPx(el: HTMLElement): number {
  // a protected field block carries its line rule on the inner field element
  const inner = el.classList.contains('doc-protected-field')
    ? el.querySelector<HTMLElement>(':scope > .doc-toc-line, :scope > .doc-field-text')
    : null
  if (inner) el = inner
  const mult = autoLineMultOf(el)
  if (mult <= 1) return 0
  const cs = getComputedStyle(el)
  // typed line grids centre the multiple's extra in the snapped cell (unprobed at
  // the page bottom): keep charging the full box there
  if (cs.getPropertyValue('--doc-grid-pitch').trim()) return 0
  // a picture line's extra is the bottom padding past its effectExtent (styles.css --doc-pic-extra)
  if (el.classList.contains('doc-img-para')) {
    const effect = parseFloat(el.style.getPropertyValue('--doc-pic-effect-b')) || 0
    return Math.max(0, (parseFloat(cs.paddingBottom) || 0) - effect)
  }
  const lh = parseFloat(cs.lineHeight)
  if (!Number.isFinite(lh) || lh <= 0) return 0
  return lh * (1 - 1 / mult)
}

const ZERO_SLOT_PX = 0.05

/**
 * Per-top-level-block buckets for the whole-flow scans below: one
 * querySelectorAll per selector over the flow instead of a subtree scan per
 * block — the scan returns only matches, so a document of plain paragraphs
 * pays a handful of whole-DOM walks instead of one per block (genoffice#526).
 * Each hit climbs to its top-level owner (a direct child of the flow root);
 * hits outside any block (page-gap widgets) land under the widget element and
 * are never looked up. Bucket lists keep document order, matching what the
 * per-block scans returned.
 */
function bucketByBlock(pm: HTMLElement, selector: string): Map<HTMLElement, HTMLElement[]> {
  const byOwner = new Map<HTMLElement, HTMLElement[]>()
  for (const hit of pm.querySelectorAll(selector)) {
    let owner = hit.parentElement
    while (owner && owner.parentElement !== pm) owner = owner.parentElement
    if (!owner || owner === pm) continue
    const el = owner as HTMLElement
    const list = byOwner.get(el)
    if (list) list.push(hit as HTMLElement)
    else byOwner.set(el, [hit as HTMLElement])
  }
  return byOwner
}

/**
 * `planOf` (pagination-plan.ts) supplies the model facts of a block — break
 * classification, row keep policy; with it only the geometry is read here
 * (break-line Ys, boxes). Blocks without a plan (decoration widgets, callers
 * without an editor) fall back to the DOM scrape of the same facts.
 */
const isJoinHead = (el: HTMLElement) => getComputedStyle(el).display === 'inline'

function planKeepWhole(
  el: HTMLElement,
  plan: BlockPlan,
  widthPx: number,
  zoomFactor: number,
): { bottom: number; maxH: number } | undefined {
  if (plan.kind !== 'protected' || !el.classList.contains('doc-protected-floating'))
    return undefined
  const facts = wrapperObstacleFacts(el, plan, widthPx, zoomFactor)
  return facts ? floatKeepWholeOf(facts) : undefined
}

export function measureBlocks(
  pm: HTMLElement,
  origin: number,
  zoomFactor: number,
  planOf?: PlanOf,
  /** also run the scrape for planned blocks and keep it on `shadowScrape` (I9) */
  shadow = false,
  /** float-carry spacer height before a block, by its position (the engine's
   *  decoration spec); plan-less blocks read the spacer element */
  carryOf?: (pos: number) => number | undefined,
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
  // consecutive zero-height break carriers collapse their margins onto one DOM
  // top; each gets its own virtual slot so its page can be told apart
  let lastZeroTop = -Infinity
  const showDeleted = !!pm.closest('.rev-display-original')
  const breaksByBlock = bucketByBlock(pm, '.doc-field-pagebreak, .doc-page-br')
  const colBreaksByBlock = bucketByBlock(pm, '.doc-col-br')
  const gapInlinesByBlock = bucketByBlock(pm, '.page-gap-inline')
  const relImgsByBlock = bucketByBlock(pm, '.doc-inline-img-anchor > img[data-page-rel-v="1"]')
  const relXImgsByBlock = bucketByBlock(
    pm,
    '.doc-inline-img-anchor > img[data-page-rel-x="1"]:not([data-page-rel-v="1"])',
  )
  const imgsByBlock = bucketByBlock(pm, 'img')
  const tablesByBlock = bucketByBlock(pm, 'table')
  for (const el of Array.from(pm.children) as HTMLElement[]) {
    const rect = el.getBoundingClientRect()
    if (el.classList.contains('page-gap') || el.classList.contains('page-float-host')) {
      gapAccum += rect.height
      continue
    }
    // a paragraph joined onto the previous one (pending paragraph-mark deletion
    // in a final-text view) shares its line: it extends that block, not a new
    // one. A tail with break or float facts of its own stays a block.
    if (el.classList.contains('doc-para-join-tail') && rect.height > 0) {
      const prev = blocks[blocks.length - 1]
      const liveBreak = (b: Element) => showDeleted || !b.closest('.doc-del')
      const plain =
        !breaksByBlock.get(el)?.some(liveBreak) &&
        !colBreaksByBlock.get(el)?.some(liveBreak) &&
        [gapInlinesByBlock, relImgsByBlock, relXImgsByBlock].every((m) => !m.get(el)?.length)
      if (plain && prev?.el && isJoinHead(prev.el)) {
        const bottom = (rect.bottom - origin - gapAccum) / zoomFactor
        prev.height = Math.max(prev.height, bottom - prev.top)
        prev.domHeight = prev.height
        continue
      }
    }
    if (el.classList.contains('page-float-carry')) {
      carryApplied += rect.height / zoomFactor
      continue
    }
    // floating-anchor boxes: absolute children of a zero-height wrapper; record
    // shift-neutral virtual positions so pages can be extended to contain them
    let wrapFloatBottom: number | undefined
    const plan = planOf?.(el)
    const pageBreakBefore = plan
      ? plan.pageBreakBeforeDirect
      : el.classList.contains('page-break-before')
    let nudge = 0
    if (rect.height <= 0 && pageBreakBefore) {
      const raw = (rect.top - origin - gapAccum) / zoomFactor
      if (raw <= lastZeroTop + ZERO_SLOT_PX) nudge = lastZeroTop + ZERO_SLOT_PX - raw
      lastZeroTop = raw + nudge
    }
    if (el.classList.contains('doc-protected-floating') || el.classList.contains('doc-img-float')) {
      const anchorTop = (rect.top - origin - gapAccum) / zoomFactor + nudge
      const pinned = el.classList.contains('doc-protected-pagepinned')
      for (const box of Array.from(
        el.querySelectorAll(
          ':scope > .doc-textbox, :scope > .box-rotator > .doc-textbox, :scope > .doc-img-wrap',
        ),
      )) {
        const b = (box as HTMLElement).getBoundingClientRect()
        if (b.height <= 0) continue
        const applied = appliedFloatShift(box as HTMLElement)
        const top = (b.top - origin - (pinned ? 0 : gapAccum)) / zoomFactor - applied + nudge
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
    for (const img of relImgsByBlock.get(el) ?? []) {
      const b = img.getBoundingClientRect()
      if (b.height <= 0) continue
      const applied = appliedFloatShift(img)
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
    // page-relative X only: the preview hoists them out of the column window
    for (const img of relXImgsByBlock.get(el) ?? []) {
      const b = img.getBoundingClientRect()
      if (b.height <= 0) continue
      floats.push({
        el: img,
        top: (b.top - origin - gapAccum) / zoomFactor,
        height: b.height / zoomFactor,
        anchorTop: (rect.top - origin - gapAccum) / zoomFactor,
        pinned: false,
        pageRelV: false,
        noSpill: true,
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
    // a tracked-deleted break (w:del) only breaks in the Original view
    const breakEls = (breaksByBlock.get(el) ?? []).filter(
      (b) => !b.closest('td, th, .doc-textbox') && (showDeleted || !b.closest('.doc-del')),
    )
    const hasBreak = breakEls.length > 0
    const colBreakEls = (colBreaksByBlock.get(el) ?? []).filter(
      (b) => !b.closest('td, th, .doc-textbox') && (showDeleted || !b.closest('.doc-del')),
    )
    const hasColBreak = colBreakEls.length > 0
    // zero-height blocks are skipped, except a break carrier (e.g. a floating
    // textbox whose anchor paragraph holds a page-type w:br or pageBreakBefore)
    // must still be seen
    if (rect.height <= 0 && !hasBreak && !pageBreakBefore) continue
    // in-block gaps from mid-paragraph page breaks: subtract from block height and add to the gap accumulator for later blocks
    let innerGap = 0
    for (const g of gapInlinesByBlock.get(el) ?? []) innerGap += g.getBoundingClientRect().height
    let top = (rect.top - anchorShiftPx(el) * zoomFactor - origin - gapAccum) / zoomFactor + nudge
    // a page-opening section clears the previous section's page-bottom float:
    // it starts where the shifted float ends, without the float's text distance
    const prev = blocks[blocks.length - 1]
    if (el.classList.contains('doc-sect-page-start') && prev?.pageRelVyPx !== undefined) {
      const prevDy = prev.floatShiftPx ?? 0
      if (prevDy > 0) top = Math.min(top, prev.top + prevDy + prev.height)
    }
    const height = (rect.height - innerGap) / zoomFactor
    const idxAttr = el.getAttribute('data-idx')
    // the plan's break facts apply when it saw the same break characters the
    // DOM holds (the shadow comparison keeps the two in step)
    const planned = plan && plan.breakCount === breakEls.length ? plan : undefined
    const gapRects = (gapInlinesByBlock.get(el) ?? []).map((g) => g.getBoundingClientRect())
    const gapAbove = (y: number) => gapRects.reduce((s, g) => (g.top <= y ? s + g.height : s), 0)
    // Y of the line starting after a break: its ink top, element-relative and
    // net of inline gaps a previous pass already inserted there
    const lineAfter = (r: Range, b: Element): number | undefined => {
      r.setStartAfter(b)
      r.setEnd(el, el.childNodes.length)
      const first = Array.from(r.getClientRects?.() ?? []).find((c) => c.height > 0 && c.width > 0)
      return first ? (first.top - rect.top - gapAbove(first.top)) / zoomFactor : undefined
    }
    // break-only paragraph: marked for dedicated placement — Word pushes it into a
    // deliberate blank page when its line doesn't fit at the page bottom. Word
    // renders a single break line; the DOM lays out one line box per break (the
    // trailing-break phantom after the last one is hidden by CSS, as its mark
    // shares the break's line), so the fit height is one line's share. Word only
    // charges the line's natural single-spacing extent at the page bottom (probe
    // 20260901: a double-spaced Calibri 11pt break line absorbs at 14pt remaining,
    // while an exact line demands its full exact height), so auto multiples above
    // 1 are divided out.
    // ProseMirror's separator img (a layout stub before the trailing break) is not content
    // the DOM scrape of the plan's break facts (text is measured outside the
    // break nodes: a field break carries its own label); the fallback for
    // plan-less blocks, and the shadow twin the I9 comparison reads
    const scrape = (): BreakFacts => {
      const breakOnly =
        hasBreak &&
        !(el.textContent ?? '').trim() &&
        !(imgsByBlock.get(el) ?? []).some((i) => !i.classList.contains('ProseMirror-separator'))
      const f: BreakFacts = {
        breakOnly,
        brLines: breakOnly ? el.querySelectorAll('br:not(.ProseMirror-trailingBreak)').length : 0,
        hasText: false,
        leadingCount: 0,
        trailingCount: 0,
        leadingCut: false,
        leadingColBreak: false,
        innerBreaks: [],
        hasCellBreak: !!el.querySelector('tr p.page-break-before'),
      }
      if (colBreakEls.length === 1) {
        const r = document.createRange()
        r.setStart(el, 0)
        r.setEndBefore(colBreakEls[0])
        f.leadingColBreak = !r.toString().trim()
      }
      if (hasBreak) {
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
        f.hasText = brs.some((x) => x.before || x.after)
        f.leadingCount = brs.filter((x) => !x.before).length
        f.trailingCount = brs.filter((x) => !x.after).length
        brs.forEach(({ b, before, after }, i) => {
          if (!after || !b.classList.contains('doc-page-br')) return
          if (!before && i !== f.leadingCount - 1) return
          const y = lineAfter(r, b)
          if (y === undefined) return
          f.innerBreaks.push(y)
          if (!before) f.leadingCut = true
        })
      }
      return f
    }
    const fromPlan = (p: BlockPlan): BreakFacts => {
      const innerBreaks: number[] = []
      if (p.innerBreakIdx.length > 0) {
        const r = document.createRange()
        for (const i of p.innerBreakIdx) {
          const y = lineAfter(r, breakEls[i])
          if (y !== undefined) innerBreaks.push(y)
        }
      }
      return {
        breakOnly: p.breakOnly,
        brLines: p.breakOnly ? p.hardBreaks : 0,
        hasText: p.hasText,
        leadingCount: p.leadingBreaks,
        trailingCount: p.trailingBreaks,
        leadingCut: p.leadingBreakCut,
        leadingColBreak: p.colBreakBefore,
        innerBreaks,
        hasCellBreak: p.hasCellBreak,
      }
    }
    const scraped = !planned || shadow ? scrape() : undefined
    const {
      breakOnly,
      brLines,
      hasText,
      leadingCount,
      trailingCount,
      leadingCut,
      leadingColBreak,
      innerBreaks,
      hasCellBreak,
    } = planned ? fromPlan(planned) : scraped!
    let breakOnlyLineH: number | undefined
    if (breakOnly) {
      const box = brLines > 1 ? height / brLines : height
      breakOnlyLineH = box / Math.max(1, autoLineMultOf(el))
    }
    const leadingBreak = hasText && leadingCount > 0 && !leadingCut
    const floatTable =
      el.classList.contains('doc-table-float-left') ||
      el.classList.contains('doc-table-float-right')
    const floatFlowed = floatTable && el.classList.contains('doc-table-float-flow')
    const lifted =
      el.dataset.anchorLiftBase !== undefined && (parseFloat(el.style.marginTop) || 0) < 0
    const floated =
      /(?:^|\s)img-wrap-(?:square|tight|through)-(?:left|right)(?:\s|$)/.test(el.className) ||
      el.classList.contains('doc-para-frame-float') ||
      (floatTable && !floatFlowed)
    // page/margin-anchored floated table: strip the applied --tblp-dy shift so
    // the engine sees the natural flow position (float margins move only the
    // float's own box, so no gapAccum contribution)
    // an in-flow wrapTopAndBottom band carries the same page-relative Y
    const pageRelBand = !floated && el.classList.contains('img-wrap-topBottom')
    const relVy = floated || pageRelBand ? parseFloat(el.dataset.tblpVy ?? '') : NaN
    const relVAnchor = el.dataset.tblpVanchor
    const relVSpec = el.dataset.tblpVspec
    const scrapedShift = Number.isFinite(relVy) ? parseFloat(el.dataset.tblpDy ?? '') || 0 : 0
    // plan-driven block: the float facts come from the model and the applied
    // shift from the engine's own plugin state; the dataset is CSS only
    const planGeom = plan ? (plan.tableGeom ?? plan.anchorGeom) : undefined
    const planShift = plan?.floatShift && !plan.floatShift.flow ? plan.floatShift.dyPx : 0
    const relVApplied = plan ? (planGeom?.pageRelV ? planShift : 0) : scrapedShift
    const blockImgs = imgsByBlock.get(el) ?? []
    const emptyPara = !(el.textContent ?? '').trim() && blockImgs.length === 0
    const pictures = blockImgs.filter((i) => !i.classList.contains('ProseMirror-separator'))
    const anchorOnly =
      el.tagName === 'P' &&
      !(el.textContent ?? '').trim() &&
      pictures.length > 0 &&
      pictures.every(
        (i) => i.closest('.doc-inline-img-anchor') || /(?:^|\s)doc-inline-img--/.test(i.className),
      )
    // non-reflowable blocks keep their rendered width in any column (tables,
    // anchored/inline textbox shapes; protected text paragraphs still reflow)
    const fixedWidth =
      el.tagName === 'TABLE' ||
      el.classList.contains('doc-protected-textboxes') ||
      !!tablesByBlock.get(el)?.length
    // a paragraph whose own column-spanning side-wrap picture runs past the
    // page bottom keeps its line on that page (Word draws it below the paper
    // edge); only a paragraph with no more than one line below the picture
    let paraFloatBottom: number | undefined
    if (el.tagName === 'P') {
      const img = (imgsByBlock.get(el) ?? []).find((i) =>
        i.matches(
          'img.doc-inline-img--wrap-square-left, img.doc-inline-img--wrap-square-right, ' +
            'img.doc-inline-img--wrap-tight-left, img.doc-inline-img--wrap-tight-right, ' +
            'img.doc-inline-img--wrap-through-left, img.doc-inline-img--wrap-through-right',
        ),
      )
      if (img) {
        const b = img.getBoundingClientRect()
        const lh = parseFloat(getComputedStyle(el).lineHeight) || 0
        if (
          b.height > 0 &&
          b.width >= rect.width - 36 * zoomFactor &&
          (rect.bottom - b.bottom) / zoomFactor <= 1.6 * lh
        ) {
          paraFloatBottom = (b.bottom - origin - gapAccum) / zoomFactor
        }
      }
    }
    const scrapedBandKeep =
      el.dataset.bandKeep === '1' && el.classList.contains('doc-protected-floating')
    const bandKeep =
      paraFloatBottom !== undefined || (plan ? !!planGeom?.bandKeep : scrapedBandKeep)
    const keepFloatBottom = wrapFloatBottom ?? paraFloatBottom
    // side-wrapped drawings as page obstacles: from the plan's obstacle rects
    // (the data-obs-* twins stay the fallback of a plan-less block)
    const keepWhole = bandKeep
      ? undefined
      : plan
        ? planKeepWhole(el, plan, rect.width / zoomFactor, zoomFactor)
        : el.dataset.obsBottom !== undefined && el.dataset.obsMaxh !== undefined
          ? {
              bottom: parseFloat(el.dataset.obsBottom) || 0,
              maxH: parseFloat(el.dataset.obsMaxh) || 0,
            }
          : undefined
    const liftPx = parseFloat(el.dataset.tblpLift ?? '')
    // the plan's space before: padded under additive spacing, otherwise a
    // margin the folds below consume (page lead / forced section start)
    const planSpaceBeforePx =
      plan && plan.kind !== 'table' && !plan.spaceBeforeInner && plan.spaceBeforePt !== undefined
        ? (plan.spaceBeforePt * 96) / 72
        : undefined
    const additive = additiveParagraphSpacing()
    const padBeforePx =
      additive && !el.matches('table')
        ? (planSpaceBeforePx ?? (parseFloat(getComputedStyle(el).paddingTop) || 0))
        : 0
    const marginBeforePx =
      planSpaceBeforePx === undefined ? undefined : additive ? 0 : planSpaceBeforePx
    // the data-tblp-* / band facts: from the plan when the block has one, the
    // scrape otherwise (and as the shadow twin of a planned block)
    const scrapedFloat = {
      ...(floatFlowed ? { floatFlowed: true as const } : {}),
      ...(liftPx > 0 ? { liftPx } : {}),
      ...(el.dataset.tblpPush === '1' ? { floatPushWhole: true as const } : {}),
      ...(paraFloatBottom !== undefined || scrapedBandKeep ? { bandKeep: true } : {}),
      ...(Number.isFinite(relVy) && (relVAnchor === 'page' || relVAnchor === 'margin')
        ? {
            pageRelVyPx: relVy,
            pageRelVAnchor: relVAnchor as 'page' | 'margin',
            ...(relVSpec === 'top' || relVSpec === 'center' || relVSpec === 'bottom'
              ? { pageRelVSpec: relVSpec as 'top' | 'center' | 'bottom' }
              : {}),
            ...(pageRelBand ? { pageRelVApplied: scrapedShift } : {}),
            floatShiftPx: scrapedShift,
          }
        : {}),
    }
    const plannedFloat = plan
      ? {
          ...(floatFlowed ? { floatFlowed: true as const } : {}),
          ...(planGeom?.liftPx ? { liftPx: planGeom.liftPx } : {}),
          ...(planGeom?.pushWhole ? { floatPushWhole: true as const } : {}),
          ...(bandKeep ? { bandKeep: true } : {}),
          ...(planGeom?.pageRelV
            ? {
                pageRelVyPx: planGeom.pageRelV.yPx,
                pageRelVAnchor: planGeom.pageRelV.band,
                ...(planGeom.pageRelV.spec ? { pageRelVSpec: planGeom.pageRelV.spec } : {}),
                ...(!planGeom.floated ? { pageRelVApplied: planShift } : {}),
                floatShiftPx: planShift,
              }
            : {}),
        }
      : scrapedFloat
    const carryAppliedPx = plan && carryOf ? (carryOf(plan.pos) ?? 0) : carryApplied
    blocks.push({
      top: top - relVApplied,
      height,
      domHeight: height,
      ...(padBeforePx > 0.5 ? { padBeforePx } : {}),
      ...(marginBeforePx !== undefined ? { marginBeforePx } : {}),
      ...(plan ? { dropAtPageTop: plan.dropAtPageTop } : {}),
      ...(floated ? { floated: true } : {}),
      ...(floated && lifted ? { lifted: true } : {}),
      ...(floatTable ? { floatTable: true } : {}),
      ...plannedFloat,
      ...(bandKeep && keepFloatBottom !== undefined ? { floatBottom: keepFloatBottom } : {}),
      ...(keepWhole
        ? {
            floatKeepWhole: true,
            floatBottom: top - relVApplied + keepWhole.bottom,
            floatMaxH: keepWhole.maxH,
          }
        : {}),
      ...(emptyPara ? { emptyPara: true } : {}),
      ...(anchorOnly ? { anchorOnly: true } : {}),
      ...(fixedWidth
        ? { fixedWidthPx: rect.width / zoomFactor }
        : { widthPx: rect.width / zoomFactor }),
      breakBefore: pageBreakBefore || leadingBreak || undefined,
      breakBeforeBr: leadingBreak || undefined,
      ...(hasText && leadingCount > 1 ? { extraBreaksBefore: leadingCount - 1 } : {}),
      breakAfter: trailingCount > 0 || undefined,
      ...(trailingCount > 1 ? { extraBreaksAfter: trailingCount - 1 } : {}),
      ...(innerBreaks.length > 0 ? { innerBreaks } : {}),
      ...(leadingCut ? { leadingBreakCut: true } : {}),
      ...(hasCellBreak ? { hasCellBreak: true } : {}),
      ...(planned?.rowPlan ? { rowPlan: planned.rowPlan } : {}),
      ...(plan && scraped
        ? {
            shadowScrape: {
              ...scrapeFlags(scraped, pageBreakBefore, hasColBreak, rect.height <= 0),
              ...scrapedFloat,
            },
          }
        : {}),
      colBreakBefore: leadingColBreak || undefined,
      colBreakAfter: (hasColBreak && !leadingColBreak) || undefined,
      breakForce: (hasBreak && rect.height <= 0) || undefined,
      el,
      ...(breakOnlyLineH !== undefined ? { breakOnlyLineH } : {}),
      ...(idxAttr ? { docxIndex: parseInt(idxAttr, 10) } : {}),
      ...(carryAppliedPx > 0 ? { carryAppliedPx } : {}),
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
    // a page/margin-anchored float takes no flow space: the distance to the
    // next block is its own applied shift (or that block clearing it), and
    // folding it in would cancel the shift on the next pass
    if (blocks[i].pageRelVyPx !== undefined) continue
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
    !!first?.el && (first.el.matches('table') || !!tablesByBlock.get(first.el)?.length)
  if (first && !firstIsTable) {
    // the plan's margin for a planned block, the DOM offset otherwise; what the
    // DOM adds beyond the plan is editor chrome (break indicator) and stays an
    // offset Word reserves nothing for
    const lead = first.marginBeforePx ?? first.top
    if (first.shadowScrape) first.shadowScrape.leadPx = first.top
    if (lead > 0.5) {
      const rest = first.top - lead
      first.spaceBeforePx = (first.spaceBeforePx ?? 0) + lead
      first.height += lead
      first.top = rest > 0.1 ? rest : 0
      first.leadFoldPx = lead
    }
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
    return (
      (rect.bottom - anchorShiftPx(el) * factor - baseTop) / factor +
      (parseFloat(getComputedStyle(el).marginBottom) || 0)
    )
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
 *  `styleKeepNext` resolves a cell paragraph's pStyle to its keepNext. Only the FIRST paragraph
 *  of the row's FIRST cell counts (Word probes 2026-09-17 / 2026-10-13: its keepNext chains the
 *  row to the next row; a keepNext on a later paragraph or in a later cell does nothing). */
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
  for (const m of outerRowsXml(tableXml).matchAll(/<w:tr[\s>][\s\S]*?(?=<w:tr[\s>]|<\/w:tbl>)/g)) {
    const trPr = m[0].match(/<w:trPr>[\s\S]*?<\/w:trPr>/)?.[0] ?? ''
    let keepNext = false
    const firstCell = /<w:tc[\s>]([\s\S]*?)(?=<w:tc[\s>]|<\/w:tr>|$)/.exec(m[0])?.[1] ?? ''
    const firstP =
      /<w:p(?:\s[^>]*)?\/>|<w:p[\s>][\s\S]*?(?=<\/w:p>|<w:p[\s>])/.exec(firstCell)?.[0] ?? ''
    const pPr = /<w:pPr>([\s\S]*?)<\/w:pPr>/.exec(firstP)?.[1]
    if (pPr) {
      const direct = /<w:keepNext\b[^>]*>/.exec(pPr)?.[0]
      if (direct) keepNext = !/w:val="(?:0|false)"/.test(direct)
      else {
        const styleId = /<w:pStyle w:val="([^"]+)"/.exec(pPr)?.[1]
        if (styleId && styleKeepNext?.(styleId)) keepNext = true
      }
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

/** The table XML with every nested table removed: nested rows are cell content,
 *  counting them shifts every later flag onto the wrong outer row. */
export function outerRowsXml(tableXml: string): string {
  let depth = 0
  let out = ''
  for (const m of tableXml.matchAll(/<w:tbl[\s>]|<\/w:tbl>|[^<]+|<[^>]*>/g)) {
    const tok = m[0]
    if (/^<w:tbl[\s>]/.test(tok)) {
      depth++
      if (depth === 1) out += tok
      continue
    }
    if (tok === '</w:tbl>') {
      if (depth === 1) out += tok
      depth--
      continue
    }
    if (depth <= 1) out += tok
  }
  return out
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
export function applyBlockMeta(
  blocks: BlockBox[],
  metaOf: BlockMetaOf | undefined,
  zoomFactor = 1,
): void {
  for (const b of blocks) {
    // flags edited in the dialog live on the element: they win over the parsed
    // paragraph's meta so a change re-paginates before the document is saved
    const meta = b.docxIndex === undefined ? undefined : metaOf?.(b.docxIndex)
    const flags = effectiveParaFlags(b.el, meta)
    if (flags.keepNext) b.keepNext = true
    if (flags.keepLines) b.keepLines = true
    if (flags.widowControl === false) b.widowControl = false
    if (!meta) continue
    if (meta.modernTableHeaders) b.modernTableHeaders = true
    if (meta.cellWidowOff) b.cellWidowOff = true
    if (meta.breakBefore) b.breakBefore = true
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
