import { Extension } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import type { EditorView } from '@tiptap/pm/view'
import type { LineAnchor } from '../pagination'
import { rangeSlot } from '../dom-range'

const anchorRange = rangeSlot()

const key = new PluginKey<DecorationSet>('paginationGaps')

/**
 * Always-on pagination in the canvas: renders a "page gap" widget before each
 * page-leading block (previous page's bottom margin + gray inter-page band +
 * next page's top margin). Decorations are visual only; document content and
 * saving are unaffected. Positions are driven by App's pagination measurement;
 * gaps map along transactions while editing and are rebuilt wholesale after a debounce.
 */
export const PaginationGapsExtension = Extension.create({
  name: 'paginationGaps',

  addProseMirrorPlugins() {
    return [
      new Plugin<DecorationSet>({
        key,
        state: {
          init: () => DecorationSet.empty,
          apply(tr, set) {
            const next = tr.getMeta(key) as DecorationSet | undefined
            if (next) return next
            return set.map(tr.mapping, tr.doc)
          },
        },
        props: {
          decorations(state) {
            return key.getState(state)
          },
        },
      }),
    ]
  },
})

const rowFillKey = new PluginKey<DecorationSet>('paginationRowFills')

/**
 * Split declared-height table rows: node decorations stretching the tr to the
 * engine's target height (Word re-honors an atLeast trHeight on the continuation
 * page fragment). Separate from the page-gap set: the pagination preview clears
 * the gaps while measuring, but these heights are real layout that must persist.
 */
export const RowFillsExtension = Extension.create({
  name: 'paginationRowFills',
  addProseMirrorPlugins() {
    return [
      new Plugin<DecorationSet>({
        key: rowFillKey,
        state: {
          init: () => DecorationSet.empty,
          apply(tr, set) {
            const next = tr.getMeta(rowFillKey) as DecorationSet | undefined
            if (next) return next
            return set.map(tr.mapping, tr.doc)
          },
        },
        props: {
          decorations(state) {
            return rowFillKey.getState(state)
          },
        },
      }),
    ]
  },
})

/** The fill renders an integer CSS height, so the stored extra is measured
 *  against that rounded height: the next pass recovers the natural height
 *  exactly instead of drifting by the rounding remainder every remeasure. */
export function rowFillAttrs(
  targetPx: number,
  extraPx = 0,
): { style: string; 'data-split-extra'?: string } {
  const h = Math.round(targetPx)
  const extra = extraPx > 0 ? h - (targetPx - extraPx) : 0
  return { style: `height:${h}px`, ...(extra > 0 ? { 'data-split-extra': extra.toFixed(1) } : {}) }
}

/** Apply/replace the split-row height patches (an empty list clears them). */
export function setRowFills(
  view: EditorView,
  fills: Array<{ el: Element; targetPx: number; extraPx?: number }>,
): void {
  const decos: Decoration[] = []
  for (const [i, fill] of fills.entries()) {
    const attrs = rowFillAttrs(fill.targetPx, fill.extraPx)
    try {
      const $inside = view.state.doc.resolve(view.posAtDOM(fill.el, 0))
      for (let d = $inside.depth; d > 0; d--) {
        if ($inside.node(d).type.name !== 'docTableRow') continue
        decos.push(
          Decoration.node($inside.before(d), $inside.after(d), attrs, {
            key: `row-fill-${i}-${attrs.style}-${attrs['data-split-extra'] ?? ''}`,
          }),
        )
        break
      }
    } catch {
      // unmapped DOM (nested-table NodeView etc.): skip, the row keeps its natural height
    }
  }
  const next = DecorationSet.create(view.state.doc, decos)
  const prev = rowFillKey.getState(view.state)
  if (!prev || !sameGaps(prev, next))
    view.dispatch(view.state.tr.setMeta(rowFillKey, next).setMeta('addToHistory', false))
}

/**
 * Blocks whose sole line exceeds the column capacity (oversized inline
 * pictures): clip the block to the engine's landing-column capacity (Word
 * overflow-clips such a line at the page bottom instead of painting into later
 * pages). The targets are protected-block NodeViews, which don't apply node
 * decorations — patch their DOM directly with the observer paused (the
 * column-layout technique); re-applied by every remeasure pass, and the
 * data-oversize-clip marker lets fillLineBoxes re-derive the flag from the
 * unclipped ink so the layout can't oscillate. Real layout, like the row
 * fills: preview clones and print inherit it. An empty list clears all clips.
 */
export function setOversizeClips(
  view: EditorView,
  clips: Array<{ el: HTMLElement; clipPx: number }>,
): void {
  const obs = (view as unknown as { domObserver?: { stop(): void; start(): void } }).domObserver
  obs?.stop()
  try {
    const want = new Map(clips.map((c) => [c.el, c.clipPx]))
    for (const el of Array.from(view.dom.querySelectorAll<HTMLElement>('[data-oversize-clip]'))) {
      if (want.has(el)) continue
      el.removeAttribute('data-oversize-clip')
      el.classList.remove('doc-oversize-clip')
      el.style.removeProperty('max-height')
    }
    for (const [el, clipPx] of want) {
      const px = clipPx.toFixed(1)
      if (el.dataset.oversizeClip === px) continue
      el.dataset.oversizeClip = px
      el.classList.add('doc-oversize-clip')
      el.style.maxHeight = `${px}px`
    }
  } finally {
    obs?.start()
  }
}

const floatVKey = new PluginKey<DecorationSet>('paginationFloatVShifts')

/**
 * Page/margin-anchored floated tables (w:tblpPr vertAnchor page|margin): node
 * decorations carrying the engine's downward shift to the tblpY target as the
 * --tblp-dy margin. Like the row fills, these are real layout (not page-gap
 * decoration) and must persist through the preview's gap clearing.
 */
export const FloatVShiftsExtension = Extension.create({
  name: 'paginationFloatVShifts',
  addProseMirrorPlugins() {
    return [
      new Plugin<DecorationSet>({
        key: floatVKey,
        state: {
          init: () => DecorationSet.empty,
          apply(tr, set) {
            const next = tr.getMeta(floatVKey) as DecorationSet | undefined
            if (next) return next
            return set.map(tr.mapping, tr.doc)
          },
        },
        props: {
          decorations(state) {
            return floatVKey.getState(state)
          },
        },
      }),
    ]
  },
})

/** transaction meta set when the set of flowed floating tables changed: the
 *  reflow moves real layout, so the pagination pass must run again */
export const floatFlowChangedMeta = 'paginationFloatFlowChanged'

/** element inside a CSS-floated (not flowed) w:tblpPr table */
export function insideFloatTable(el: Element): boolean {
  const t = el.parentElement?.closest('.doc-table-float-left, .doc-table-float-right')
  return !!t && !t.classList.contains('doc-table-float-flow')
}

/** Apply/replace the anchored-table shifts, the over-page float flows and the
 *  float-carry spacers (an empty list clears them). A `flow` entry renders the
 *  floating table in normal flow (doc-table-float-flow) instead of shifting it.
 *  A `carryPx` entry puts a flow spacer of that height before `el` (the anchor
 *  paragraph of a split floating table starts beside the table's last portion;
 *  the spacer is real flow space, unlike a page gap). */
export function setFloatVShifts(
  view: EditorView,
  shifts: Array<{ el: Element; dyPx: number; flow?: boolean; carryPx?: number }>,
): void {
  const decos: Decoration[] = []
  const flowKeys: string[] = []
  for (const [i, shift] of shifts.entries()) {
    if (shift.carryPx !== undefined) {
      const carry = Math.round(shift.carryPx * 10) / 10
      if (carry < 0.5) continue
      try {
        const pos = view.state.doc.resolve(view.posAtDOM(shift.el, 0)).before(1)
        decos.push(
          Decoration.widget(
            pos,
            () => {
              const div = document.createElement('div')
              div.className = 'page-float-carry'
              div.contentEditable = 'false'
              div.style.height = `${carry}px`
              return div
            },
            { side: -1, key: `tblp-carry-${i}-${carry}` },
          ),
        )
      } catch {
        /* stale element: skip this round */
      }
      continue
    }
    const dy = Math.round(shift.dyPx * 10) / 10
    if (!shift.flow && Math.abs(dy) < 0.5) continue
    try {
      const $inside = view.state.doc.resolve(view.posAtDOM(shift.el, 0))
      for (let d = $inside.depth; d > 0; d--) {
        const name = $inside.node(d).type.name
        if (name !== 'docTable' && name !== 'docProtected') continue
        const from = $inside.before(d)
        if (shift.flow) {
          flowKeys.push(String(from))
          decos.push(
            Decoration.node(
              from,
              $inside.after(d),
              { class: 'doc-table-float-flow' },
              { key: `tblp-flow-${i}`, flow: true },
            ),
          )
        } else {
          decos.push(
            Decoration.node(
              from,
              $inside.after(d),
              { style: `--tblp-dy:${dy}px`, 'data-tblp-dy': String(dy) },
              { key: `tblp-dy-${i}-${dy}` },
            ),
          )
        }
        break
      }
    } catch {
      // unmapped DOM: skip, the table keeps its flow position
    }
  }
  const next = DecorationSet.create(view.state.doc, decos)
  const prev = floatVKey.getState(view.state)
  if (!prev || !sameGaps(prev, next)) {
    const prevFlow = (prev?.find(undefined, undefined, (spec) => spec.flow === true) ?? [])
      .map((d) => String(d.from))
      .sort()
    const flowChanged = prevFlow.join(',') !== flowKeys.sort().join(',')
    view.dispatch(
      view.state.tr
        .setMeta(floatVKey, next)
        .setMeta(floatFlowChangedMeta, flowChanged)
        .setMeta('addToHistory', false),
    )
  }
}

export interface GapMetrics {
  marginTop: number
  marginBottom: number
  /** bleed the gap needs to reach the paper edges: the next page's side margins for
   *  block gaps, the host block's / cell's offset from the paper for inline and
   *  in-cell gaps (that offset includes paragraph indents) */
  marginLeft: number
  marginRight: number
  /** the next page's section side margins regardless of gap kind: where its
   *  header/footer strips belong (alignGapHfStrips); defaults to marginLeft/Right */
  sectionMarginLeft?: number
  sectionMarginRight?: number
  sectionMarginTop?: number
  /** the next page's left edge and width on the shared paper (differing-width
   *  documents center narrower pages); absent = the page spans the paper */
  pageLeft?: number
  pageWidth?: number
}

/** height of the gray inter-page band inside a page gap */
export const GAP_BAND = 28

export type GapKind = 'block' | 'inline' | 'table' | 'cut' | 'cell'

export function makeGapEl(m: GapMetrics, kind: GapKind, cols?: number): HTMLElement {
  const gap = document.createElement(kind === 'table' ? 'tr' : 'div')
  gap.contentEditable = 'false'
  if (m.pageLeft !== undefined && m.pageWidth !== undefined) {
    gap.style.setProperty('--gap-page-x', `${m.pageLeft}px`)
    gap.style.setProperty('--gap-page-w', `${m.pageWidth}px`)
  }
  if (kind === 'cut') {
    // in-row table break cut point: zero-height dashed marker (takes no layout height, coordinate bookkeeping unchanged)
    gap.className = 'page-gap-cut'
    return gap
  }
  gap.style.height = `${m.marginBottom + GAP_BAND + m.marginTop}px`
  gap.style.setProperty('--gap-mb', `${m.marginBottom}px`)
  gap.style.setProperty('--gap-mt', `${m.marginTop}px`)
  // written even when zero: a page without a header push must clear the previous one
  if (m.sectionMarginTop !== undefined)
    gap.dataset.topPush = Math.max(0, m.marginTop - m.sectionMarginTop).toFixed(1)
  // the next page's own section side margins: a section whose margins differ from
  // the canvas' first section gets its header/footer strips placed on ITS text
  // column (alignGapHfStrips), not the cover's margin-less one. Every gap kind
  // carries them — inline/cell gaps' marginLeft is the host block's paper offset
  // (indent included), which is the wrong place for a strip
  gap.style.setProperty('--gap-ml', `${m.sectionMarginLeft ?? m.marginLeft}px`)
  gap.style.setProperty('--gap-mr', `${m.sectionMarginRight ?? m.marginRight}px`)
  if (kind === 'table') {
    // A real spanning cell is required here. Chromium's collapsed-border table
    // painting can leak the neighboring row's border/fill through a cell-less
    // display:table-row, leaving a colored remnant in the gray page gutter.
    gap.className = 'page-gap page-gap-inline page-gap-table'
    const cell = document.createElement('td')
    // colSpan must equal the table's real column count: a larger span widens the
    // column grid, and in a fixed-layout table WITHOUT a <colgroup> (AI-inserted
    // tables carry no colWidthsPct) Chromium then splits width:100% across all
    // phantom columns, collapsing every real cell to ~1px — which changes the
    // measured heights and sets off an endless remeasure/re-gap flicker loop
    cell.colSpan = Math.max(1, Math.round(cols ?? 1))
    cell.contentEditable = 'false'
    const fill = document.createElement('div')
    fill.className = 'page-gap-table-fill'
    cell.appendChild(fill)
    gap.appendChild(cell)
    return gap
  }
  gap.className =
    kind === 'cell'
      ? // in-cell gap (single-column in-row table cut): inline gap + opaque bands covering the cell's fill/borders
        'page-gap page-gap-inline page-gap-cell'
      : kind === 'inline'
        ? 'page-gap page-gap-inline'
        : 'page-gap'
  gap.style.marginLeft = `-${m.marginLeft}px`
  gap.style.marginRight = `-${m.marginRight}px`
  // inline-block (not block-level): avoids block-in-inline anonymous-box splitting,
  // which would re-apply text-indent/alignment on continuation lines and drift line
  // breaks; negative margins don't widen an inline-block, so width explicitly adds the bleed
  if (kind !== 'block') gap.style.width = `calc(100% + ${m.marginLeft + m.marginRight}px)`
  return gap
}

/** Page gap anchor: el = top-level page-leading block (inserted before it); pos = document position at a mid-paragraph/mid-table break */
export type PageGapSpec = {
  metrics: GapMetrics
  /** Previous page's bottom footnote area (a ready-made positioned/sized element) and its content signature (change triggers rebuild) */
  notes?: HTMLElement
  notesKey?: string
  /** Previous page's footer / next page's header (ready-made positioned .page-gap-hf elements) and their content signature */
  hfEls?: HTMLElement[]
  hfKey?: string
  /** w:tblHeader repetition: cloned header rows rendered right after a table gap
   *  (the slicing engine already reserved their height on the new page) */
  repeatHeaderEls?: HTMLElement[]
  repeatHeaderKey?: string
  /** page forced by an explicit break (w:br page / pageBreakBefore): Word drops the
   *  lead block's space-before, so a node decoration zeroes its margin-top */
  suppressLeadMt?: boolean
  /** mixed-column page above: pull the gap (and everything below) up over the
   *  vacated stacked-column space (negative margin-top, neutralized while measuring) */
  pullUp?: number
  /** slice boundary (gapless flow px) this gap opens; syncFloatShifts prefers it
   *  over the widget's DOM position (they differ at trailing float-spill pages) */
  boundaryY?: number
  /** table gaps: the host table's real column count (the spanning cell's colSpan
   *  must not widen the column grid — see makeGapEl) */
  cols?: number
} & (
  | {
      el: HTMLElement
      /** invisible spacer of this height instead of a page gap: the anchor
       *  paragraph of a split floating table skips the table's in-table gaps */
      carryPx?: number
    }
  | { pos: number; kind?: Exclude<GapKind, 'block'> }
)

/** Rebuild all page gaps (an empty list clears them); each gap carries its own margins (sections differ) */
export function setPageGaps(
  view: EditorView,
  gaps: PageGapSpec[],
  /** first page's floating header images (later pages get theirs via their gap):
   *  a zero-height page-float-host widget that measurement/clones ignore
   *  (not page-gap: page-boundary consumers must not count it as a page) */
  firstPageEls?: { els: HTMLElement[]; key: string },
): void {
  const decos: Decoration[] = []
  if (firstPageEls?.els.length) {
    decos.push(
      Decoration.widget(
        0,
        () => {
          const wrap = document.createElement('div')
          wrap.className = 'page-float-host page-first-floats'
          wrap.contentEditable = 'false'
          for (const el of firstPageEls.els) wrap.appendChild(el)
          return wrap
        },
        { side: -1, key: `page-first-floats-${firstPageEls.key}` },
      ),
    )
  }
  let ordinal = -1
  for (const gap of gaps) {
    ordinal++
    const { metrics, notes, hfEls } = gap
    let pos: number
    let kind: GapKind
    if ('el' in gap) {
      kind = 'block'
      try {
        const $inside = view.state.doc.resolve(view.posAtDOM(gap.el, 0))
        pos = $inside.before(1)
        if (gap.carryPx !== undefined) {
          const carry = Math.round(gap.carryPx)
          decos.push(
            Decoration.widget(
              pos,
              () => {
                const div = document.createElement('div')
                div.className = 'page-gap page-gap-carry'
                div.contentEditable = 'false'
                div.style.height = `${carry}px`
                return div
              },
              { side: -1, key: `page-gap-carry-${ordinal}-${carry}` },
            ),
          )
          continue
        }
        if (gap.suppressLeadMt)
          decos.push(
            Decoration.node(
              pos,
              $inside.after(1),
              { class: 'page-break-lead' },
              { key: `page-lead-${ordinal}` },
            ),
          )
      } catch {
        continue
      }
    } else {
      pos = gap.pos
      kind = gap.kind ?? 'inline'
    }
    // boundaryY in the key: a reused widget must not keep a stale boundary
    // (cols too: a table-structure edit must rebuild the spanning cell)
    const mKey = `${metrics.marginTop},${metrics.marginBottom},${metrics.marginLeft},${metrics.marginRight},${Math.round(gap.pullUp ?? 0)},${Math.round(gap.boundaryY ?? -1)},${gap.cols ?? 0},${Math.round(metrics.pageLeft ?? -1)}`
    decos.push(
      Decoration.widget(
        pos,
        () => {
          const el = makeGapEl(metrics, kind, gap.cols)
          if (gap.boundaryY != null) el.dataset.boundaryY = gap.boundaryY.toFixed(1)
          // margins don't apply to table-rows and cuts are zero-height markers;
          // tables inside mixed-column regions are out of scope anyway (v1)
          if (gap.pullUp && kind !== 'cut' && kind !== 'table')
            el.style.marginTop = `-${gap.pullUp}px`
          if (notes) el.appendChild(notes)
          if (hfEls && (kind === 'block' || kind === 'inline' || kind === 'cell')) {
            for (const hf of hfEls) el.appendChild(hf)
          } else if (hfEls && kind === 'table') {
            // table-row gaps position their strips inside the absolutely-filled cell;
            // only zero-height cut markers still can't carry them
            const fill = el.querySelector('.page-gap-table-fill')
            if (fill) for (const hf of hfEls) fill.appendChild(hf)
          }
          return el
        },
        {
          side: -1,
          // keyed by page ordinal, NOT pos: edits above a gap shift its mapped
          // position without changing the page, so an ordinal key lets sameGaps
          // skip the dispatch entirely and lets PM reuse the widget DOM when a
          // dispatch does happen
          // full kind, not kind[0]: 'cut' and 'cell' would collide and skip the rebuild
          key: `page-gap-${kind}-${ordinal}-${mKey}${gap.notesKey ? `-${gap.notesKey}` : ''}${gap.hfKey ? `-${gap.hfKey}` : ''}`,
        },
      ),
    )
    // repeated header rows (w:tblHeader) directly after the table gap: one widget per
    // cloned tr, side 0 so they land between the gap (side -1) and the split row
    if (kind === 'table' && gap.repeatHeaderEls?.length) {
      gap.repeatHeaderEls.forEach((rowEl, i) => {
        decos.push(
          Decoration.widget(pos, () => rowEl, {
            side: 0,
            key: `page-gap-rh-${pos}-${i}-${gap.repeatHeaderKey ?? ''}`,
          }),
        )
      })
    }
  }
  const next = DecorationSet.create(view.state.doc, decos)
  const prev = key.getState(view.state)
  if (!prev || !sameGaps(prev, next))
    view.dispatch(view.state.tr.setMeta(key, next).setMeta('addToHistory', false))
  // DOM-only rowspan bridging; observer paused so PM never re-parses the mutated
  // cells (a reparse would wipe cell attrs that don't round-trip through DOM)
  const obs = (view as unknown as { domObserver?: { stop(): void; start(): void } }).domObserver
  obs?.stop()
  try {
    syncPhantomRowspans(view.dom as HTMLElement)
  } finally {
    obs?.start()
  }
}

/** Line top of a cut anchor (screen px); falls back to the parent element's top. */
function anchorTop(a: LineAnchor): number | null {
  if (a.node instanceof Element) return a.node.getBoundingClientRect().top
  if (a.node.length > 0) {
    const range = anchorRange()
    range.setStart(a.node, Math.min(a.charOffset, a.node.length - 1))
    range.setEnd(a.node, Math.min(a.charOffset + 1, a.node.length))
    for (const r of range.getClientRects()) if (r.height > 0) return r.top
  }
  return a.node.parentElement?.getBoundingClientRect().top ?? null
}

/**
 * Cut markers whose anchor lives in a non-PM-addressable subtree (the read-only
 * nested-table NodeView renders arbitrarily deep tables as one PM node): a widget
 * decoration would collapse to the node's start position, stacking every page gap
 * at one spot. Draw them as zero-height absolute overlays on the page wrap instead
 * (no layout height, matching the multi-cell in-row cut markers).
 */
export function syncCutOverlays(
  wrap: HTMLElement,
  anchors: LineAnchor[],
  zoomFactor: number,
): void {
  let layer = wrap.querySelector(':scope > .page-cut-overlays') as HTMLElement | null
  if (anchors.length === 0) {
    layer?.remove()
    return
  }
  if (!layer) {
    layer = document.createElement('div')
    layer.className = 'page-cut-overlays'
    wrap.appendChild(layer)
  }
  layer.textContent = ''
  const wrapTop = wrap.getBoundingClientRect().top
  for (const a of anchors) {
    const top = anchorTop(a)
    if (top == null) continue
    const el = document.createElement('div')
    el.className = 'page-gap-cut page-cut-overlay'
    el.style.top = `${(top - wrapTop) / zoomFactor}px`
    layer.appendChild(el)
  }
}

export type PageBorderSideName = 'top' | 'right' | 'bottom' | 'left'

export interface PageBorderSide {
  /** CSS border shorthand for this side (absent for tiled art patterns) */
  css?: string
  /** tiled art-border approximation: background of a strip along this side */
  art?: PageBorderArtStrip
  /** border inset from the paper edge (CSS px, unzoomed) */
  insetPx: number
}

export interface PageBorderArtStrip {
  /** pattern height (CSS px, unzoomed) = the strip's thickness */
  sizePx: number
  backgroundImage: string
  backgroundSize: string
  backgroundRepeat: string
}

export interface PageBorderStyle {
  /** pages the border applies to (w:pgBorders w:display); undefined = all pages */
  display?: 'firstPage' | 'notFirstPage'
  /** w:zOrder="back": paint behind the page text */
  zOrder?: 'back'
  sides: Partial<Record<PageBorderSideName, PageBorderSide>>
}

/** Two/three-line OOXML border styles: CSS `double` is the closest match. */
const DOUBLE_BORDER_VALS = new Set([
  'double',
  'triple',
  'doubleWave',
  'thinThickSmallGap',
  'thickThinSmallGap',
  'thinThickThinSmallGap',
  'thinThickMediumGap',
  'thickThinMediumGap',
  'thinThickThinMediumGap',
  'thinThickLargeGap',
  'thickThinLargeGap',
  'thinThickThinLargeGap',
])

const DASHED_BORDER_VALS = new Set(['dashed', 'dashSmallGap', 'dotDash', 'dotDotDash'])

/** OOXML page-border line → CSS border shorthand. */
function pageBorderLineCss(val: string, widthPt: number, color: string): string {
  const px = Math.max(1, Math.round((widthPt * 96) / 72))
  if (DASHED_BORDER_VALS.has(val)) return `${px}px dashed ${color}`
  if (val === 'dotted' || val === 'dotDashSlanted') return `${px}px dotted ${color}`
  // w:sz is the individual line width; CSS double splits the total, so widen it
  if (DOUBLE_BORDER_VALS.has(val)) return `${Math.max(3, px * 2)}px double ${color}`
  return `${px}px solid ${color}`
}

/**
 * Word's art borders are bitmaps shipped with Office, tiled along the side at
 * w:sz points of height. Approximation by pattern family: stitches/teeth → a
 * zigzag line, pictogram rows (flowers, gems, people, food, …) → a dot row in
 * the bitmap's main colour, dash/block rows → CSS dashed, line-work → CSS
 * double, and single-rule bitmaps → a solid rule.
 */
type ArtFamily = 'zigzag' | 'dots' | 'dashes' | 'double' | 'solid'

const ART_ZIGZAG = new Set([
  'zigZag',
  'zigZagStitch',
  'crossStitch',
  'sawtooth',
  'sawtoothGray',
  'sharksTeeth',
  'waveline',
  'classicalWave',
  'lightning1',
  'lightning2',
  'triangles',
  'triangle1',
  'triangle2',
  'triangleParty',
  'zanyTriangles',
  'pyramids',
  'pyramidsAbove',
  'heebieJeebies',
  'marqueeToothed',
])

const ART_DASHES = new Set([
  'basicBlackDashes',
  'basicWhiteDashes',
  'basicBlackSquares',
  'basicWhiteSquares',
  'couponCutoutDashes',
  'checkedBarBlack',
  'checkedBarColor',
  'checkered',
  'film',
  'decoBlocks',
  'mosaic',
  'quadrants',
  'shadowedSquares',
  'eclipsingSquares1',
  'eclipsingSquares2',
  'circlesRectangles',
  'postageStamp',
  'marquee',
])

const ART_DOUBLE = new Set([
  'handmade1',
  'handmade2',
  'twistedLines1',
  'twistedLines2',
  'basicThinLines',
  'basicWideInline',
  'basicWideMidline',
  'basicWideOutline',
  'weavingAngles',
  'weavingBraid',
  'weavingRibbon',
  'weavingStrips',
  'chainLink',
  'doubleD',
  'papyrus',
  'tornPaper',
  'tornPaperBlack',
  'woodwork',
  'xIllusions',
  'hypnotic',
  'crazyMaze',
  'celticKnotwork',
  'vine',
  'gradient',
  'decoArch',
  'decoArchColor',
  'fans',
  'northwest',
  'southwest',
  'seattle',
  'safari',
  'swirligig',
  'cornerTriangles',
  'certificateBanner',
])

const ART_SOLID = new Set(['whiteFlowers'])

/** Main colour (and outline ink) of the coloured bitmaps; monochrome ones take w:color / black. */
const ART_COLORS: Record<string, { fill: string; ink?: string }> = {
  zigZagStitch: { fill: '#E2E2E2' },
  gems: { fill: '#BFBFBF', ink: '#000000' },
  cakeSlice: { fill: '#FEEDC9', ink: '#000000' },
  peopleHats: { fill: '#FEFE7F', ink: '#BFBFBF' },
}

function artFamilyOf(val: string): ArtFamily {
  if (ART_ZIGZAG.has(val)) return 'zigzag'
  if (ART_DASHES.has(val)) return 'dashes'
  if (ART_DOUBLE.has(val)) return 'double'
  if (ART_SOLID.has(val)) return 'solid'
  return 'dots'
}

function svgTile(size: number, body: string): string {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 ${size} ${size}">${body}</svg>`
  return `url("data:image/svg+xml,${encodeURIComponent(svg)}")`
}

/** Art side → CSS: a strip pattern for zigzag/dot rows, a border shorthand otherwise. */
function pageBorderArtCss(
  side: PageBorderSideName,
  val: string,
  heightPt: number,
  color: string | undefined,
): Pick<PageBorderSide, 'css' | 'art'> {
  const size = Math.max(2, Math.round((heightPt * 96) / 72))
  const known = ART_COLORS[val]
  const fill = color ?? known?.fill ?? (/Gray/.test(val) ? '#BFBFBF' : '#000000')
  const ink = color ? undefined : known?.ink
  const family = artFamilyOf(val)
  const horizontal = side === 'top' || side === 'bottom'
  if (family === 'dashes') return { css: `${size}px dashed ${fill}` }
  if (family === 'double') return { css: `${Math.max(3, Math.round(size / 2))}px double ${fill}` }
  if (family === 'solid') return { css: `${Math.max(1, Math.round(size / 8))}px solid ${fill}` }
  let body: string
  if (family === 'zigzag') {
    const lo = (size * 0.15).toFixed(2)
    const hi = (size * 0.85).toFixed(2)
    const mid = (size / 2).toFixed(2)
    const points = horizontal
      ? `0,${hi} ${mid},${lo} ${size},${hi}`
      : `${hi},0 ${lo},${mid} ${hi},${size}`
    body = `<polyline points="${points}" fill="none" stroke="${fill}" stroke-width="1.5"/>`
  } else {
    const r = (size * 0.32).toFixed(2)
    const stroke = ink ? ` stroke="${ink}" stroke-width="1"` : ''
    body = `<circle cx="${size / 2}" cy="${size / 2}" r="${r}" fill="${fill}"${stroke}/>`
  }
  return {
    art: {
      sizePx: size,
      backgroundImage: svgTile(size, body),
      backgroundSize: `${size}px ${size}px`,
      backgroundRepeat: horizontal ? 'repeat-x' : 'repeat-y',
    },
  }
}

/** Inline style of the art strip hugging `side` inside the page-border box. */
export function pageBorderArtStripStyle(
  side: PageBorderSideName,
  art: PageBorderArtStrip,
): Record<string, string> {
  const thickness = `${art.sizePx}px`
  const place: Record<string, string> =
    side === 'top' || side === 'bottom'
      ? { left: '0', right: '0', height: thickness, [side]: '0' }
      : { top: '0', bottom: '0', width: thickness, [side]: '0' }
  return {
    ...place,
    backgroundImage: art.backgroundImage,
    backgroundSize: art.backgroundSize,
    backgroundRepeat: art.backgroundRepeat,
  }
}

interface PageBorderSection {
  pageBorder: boolean
  pageBorderProps?: {
    display?: 'firstPage' | 'notFirstPage'
    offsetFrom?: 'page' | 'text'
    zOrder?: 'back'
    spacePt: number
    widthPt: number
    color?: string
    sides?: Partial<
      Record<
        PageBorderSideName,
        { val: string; widthPt: number; spacePt: number; color?: string; art?: true }
      >
    >
  }
  marginTop: number
  marginRight: number
  marginBottom: number
  marginLeft: number
}

/**
 * Per-side page border box (w:pgBorders) for one section, in unzoomed CSS px.
 * offsetFrom='page': w:space measures from the paper edge; 'text': from the
 * text area inward margin edge.
 */
export function pageBorderStyleOf(section: PageBorderSection): PageBorderStyle | null {
  if (!section.pageBorder) return null
  const p = section.pageBorderProps
  const inset = (marginTwips: number, spacePt: number) => {
    const spacePx = (spacePt * 96) / 72
    return !p || p.offsetFrom === 'page'
      ? spacePx
      : Math.max(0, (marginTwips / 1440) * 96 - spacePx)
  }
  const margins = {
    top: section.marginTop,
    right: section.marginRight,
    bottom: section.marginBottom,
    left: section.marginLeft,
  }
  const sides: PageBorderStyle['sides'] = {}
  // legacy single-box fallback: pageBorder set without parse-side details
  const fallback: { val: string; widthPt: number; spacePt: number; color?: string; art?: true } = {
    val: 'single',
    widthPt: p?.widthPt ?? 0.75,
    spacePt: p?.spacePt ?? 24,
    ...(p?.color ? { color: p.color } : {}),
  }
  const declared = p?.sides && Object.keys(p.sides).length > 0 ? p.sides : null
  for (const name of ['top', 'right', 'bottom', 'left'] as const) {
    const line = declared ? declared[name] : fallback
    if (!line) continue
    const insetPx = inset(margins[name], line.spacePt)
    if (line.art) {
      const color = line.color ? `#${line.color}` : undefined
      sides[name] = { ...pageBorderArtCss(name, line.val, line.widthPt, color), insetPx }
      continue
    }
    // w:color absent/auto on a parsed side = automatic (black), never a sibling side's color
    const color = `#${line.color ?? '000000'}`
    sides[name] = { css: pageBorderLineCss(line.val, line.widthPt, color), insetPx }
  }
  return {
    ...(p?.display ? { display: p.display } : {}),
    ...(p?.zOrder ? { zOrder: p.zOrder } : {}),
    sides,
  }
}

export interface PageFrame {
  top: number
  bottom: number
  /** horizontal extent on the page wrap (differing-width documents center narrower pages) */
  left: number
  width: number
}

/**
 * Page frames on the canvas: vertical bounds span between gap decorations (gap =
 * prev bottom margin + band + next top margin; zero-height cut markers are
 * boundaries too). The first page takes `first`, every later page the geometry
 * its gap carries (--gap-page-x/w) or, for markers without one, its predecessor's.
 */
export function pageFramesFromGaps(
  wrap: HTMLElement,
  zoomFactor: number,
  first?: { left: number; width: number },
): PageFrame[] {
  const wr = wrap.getBoundingClientRect()
  const paper = wrap.querySelector<HTMLElement>('.doc-page')
  const paperRect = paper?.getBoundingClientRect()
  const paperX = paperRect ? (paperRect.left - wr.left) / zoomFactor : 0
  first ??= { left: 0, width: (paperRect?.width ?? wr.width) / zoomFactor }
  // the carry spacer of a split floating table displaces the flow like a gap
  // but is no page turn (the table's own in-table gap at the same Y is)
  const gaps = Array.from(
    wrap.querySelectorAll<HTMLElement>('.page-gap:not(.page-gap-carry), .page-gap-cut'),
  )
    .map((g) => {
      const r = g.getBoundingClientRect()
      const cs = getComputedStyle(g)
      const mb = parseFloat(cs.getPropertyValue('--gap-mb')) || 0
      const mt = parseFloat(cs.getPropertyValue('--gap-mt')) || 0
      const x = parseFloat(g.style.getPropertyValue('--gap-page-x'))
      const w = parseFloat(g.style.getPropertyValue('--gap-page-w'))
      return {
        top: (r.top - wr.top) / zoomFactor,
        height: r.height / zoomFactor,
        mb,
        mt,
        page: Number.isFinite(x) && Number.isFinite(w) ? { left: x, width: w } : null,
      }
    })
    .sort((a, b) => a.top - b.top)
  const frames: PageFrame[] = []
  let geom = first
  let top = 0
  for (const g of gaps) {
    frames.push({ top, bottom: g.top + g.mb, left: paperX + geom.left, width: geom.width })
    geom = g.page ?? geom
    top = g.top + g.height - g.mt
  }
  frames.push({ top, bottom: wr.height / zoomFactor, left: paperX + geom.left, width: geom.width })
  return frames
}

/**
 * Page borders (w:pgBorders) as absolute per-page overlays on the page wrap.
 * The continuous canvas can't carry a real border per page, and w:display
 * needs pages skipped; page rects come from the gap widgets, like the
 * per-page screenshot slicing does.
 */
export function syncPageBorders(
  wrap: HTMLElement,
  style: PageBorderStyle | null,
  zoomFactor: number,
  first?: { left: number; width: number },
): void {
  let layer = wrap.querySelector(':scope > .page-border-overlays') as HTMLElement | null
  if (!style) {
    layer?.remove()
    return
  }
  if (!layer) {
    layer = document.createElement('div')
    layer.className = 'page-border-overlays'
    wrap.appendChild(layer)
  }
  // zOrder=back: the paper (.doc-page) is the editor root and owns its DOM, so
  // the layer stays a sibling and blends under the text instead (CSS)
  layer.classList.toggle('page-border-overlays-back', style.zOrder === 'back')
  layer.textContent = ''
  const { sides } = style
  let page = 0
  for (const f of pageFramesFromGaps(wrap, zoomFactor, first)) {
    if (f.bottom - f.top <= 10) continue
    const pageIdx = page++
    if (style.display === 'firstPage' && pageIdx > 0) continue
    if (style.display === 'notFirstPage' && pageIdx === 0) continue
    const el = document.createElement('div')
    el.className = 'page-border-overlay'
    const insetTop = sides.top?.insetPx ?? 0
    const insetBottom = sides.bottom?.insetPx ?? 0
    const insetLeft = sides.left?.insetPx ?? 0
    const insetRight = sides.right?.insetPx ?? 0
    el.style.top = `${f.top + insetTop}px`
    el.style.left = `${f.left + insetLeft}px`
    el.style.width = `${f.width - insetLeft - insetRight}px`
    el.style.height = `${f.bottom - f.top - insetTop - insetBottom}px`
    for (const name of ['top', 'right', 'bottom', 'left'] as const) {
      const side = sides[name]
      if (!side) continue
      if (side.css)
        el.style[`border${name[0].toUpperCase()}${name.slice(1)}` as 'borderTop'] = side.css
      if (side.art) {
        const strip = document.createElement('div')
        strip.className = 'page-border-art'
        Object.assign(strip.style, pageBorderArtStripStyle(name, side.art))
        el.appendChild(strip)
      }
    }
    layer.appendChild(el)
  }
}

/**
 * Differing-width documents: the shared .doc-page paper goes transparent and each
 * page paints its own sheet (paper color + shadow) at its centered position, so a
 * portrait page next to a landscape one keeps its real width. null clears the layer.
 */
export function syncPageSheets(
  wrap: HTMLElement,
  zoomFactor: number,
  first: { left: number; width: number } | null,
): void {
  let layer = wrap.querySelector(':scope > .page-sheets') as HTMLElement | null
  if (!first) {
    layer?.remove()
    return
  }
  if (!layer) {
    layer = document.createElement('div')
    layer.className = 'page-sheets'
    wrap.appendChild(layer)
  }
  layer.textContent = ''
  for (const f of pageFramesFromGaps(wrap, zoomFactor, first)) {
    if (f.bottom - f.top <= 10) continue
    const el = document.createElement('div')
    el.className = 'page-sheet'
    el.style.top = `${f.top}px`
    el.style.height = `${f.bottom - f.top}px`
    el.style.left = `${f.left}px`
    el.style.width = `${f.width}px`
    layer.appendChild(el)
  }
}

/**
 * Phantom table rows (page-gap rows, repeated-header clones) occupy grid row
 * slots, so a vMerge cell spanning across them exhausts its rowspan early and
 * every later cell in the row shifts one column left. Bridge: grow each
 * crossing cell's rowspan by the phantom rows inside its span, keeping the
 * source value in data-base-rowspan (clone/export paths drop phantom rows and
 * restore from it). Idempotent; with no phantom rows left it restores all cells.
 */
export function syncPhantomRowspans(root: HTMLElement): void {
  const isPhantom = (tr: HTMLTableRowElement) =>
    tr.classList.contains('page-gap') || tr.classList.contains('page-repeat-header')
  const grown = new Set<HTMLTableCellElement>()
  for (const table of Array.from(root.querySelectorAll('table'))) {
    const real: HTMLTableRowElement[] = []
    /** phantom rows between real[i-1] and real[i] */
    const phantomsBefore: number[] = []
    let pending = 0
    let hasPhantom = false
    for (const tr of Array.from(table.rows)) {
      if (isPhantom(tr)) {
        pending++
        hasPhantom = true
        continue
      }
      phantomsBefore.push(pending)
      pending = 0
      real.push(tr)
    }
    if (!hasPhantom) continue
    real.forEach((tr, start) => {
      for (const td of Array.from(tr.cells)) {
        const base = Number(td.getAttribute('data-base-rowspan')) || td.rowSpan
        if (base <= 1) continue
        // phantoms strictly inside the span [start, start+base): each boundary counts once
        let extra = 0
        for (let r = start + 1; r < Math.min(start + base, real.length); r++)
          extra += phantomsBefore[r]
        if (extra === 0) continue
        if (td.getAttribute('data-base-rowspan') !== String(base))
          td.setAttribute('data-base-rowspan', String(base))
        if (td.rowSpan !== base + extra) td.rowSpan = base + extra
        grown.add(td)
      }
    })
  }
  for (const td of Array.from(
    root.querySelectorAll<HTMLTableCellElement>('td[data-base-rowspan], th[data-base-rowspan]'),
  )) {
    if (grown.has(td)) continue
    td.rowSpan = Number(td.getAttribute('data-base-rowspan')) || 1
    td.removeAttribute('data-base-rowspan')
  }
}

function sameGaps(a: DecorationSet, b: DecorationSet): boolean {
  const keyOf = (d: Decoration) => (d.spec as { key?: string }).key ?? ''
  const af = a.find()
  const bf = b.find()
  return (
    af.length === bf.length &&
    af.every((d, i) => d.from === bf[i].from && keyOf(d) === keyOf(bf[i]))
  )
}

/**
 * Canvas display correction for floating boxes: a box is absolutely positioned
 * from its anchor's flow position, so page-gap bands inserted between the anchor
 * and the box's virtual Y are not reflected in its offset — the box would overlap
 * the gray gap / next page header area. Translate each box by the gap height
 * above its virtual position (idempotent; runs after setPageGaps).
 */
export function syncFloatShifts(
  pm: HTMLElement,
  floats: Array<{
    el: HTMLElement
    top: number
    anchorTop?: number
    pinned?: boolean
    pageRelV?: boolean
    pageRelFromPage?: boolean
  }>,
  origin: number,
  factor: number,
  firstPagePush = 0,
): void {
  if (floats.length === 0) return
  const gaps: Array<{ v: number; h: number; push?: number }> = []
  let acc = 0
  // in-table gap rows / repeated-header clones displace the DOM below them just
  // like top-level gap widgets: a float anchored after a multi-page table would
  // otherwise resolve one page too high (its virtual top already excludes them)
  for (const el of Array.from(
    pm.querySelectorAll<HTMLElement>('.page-gap, .page-float-host, .page-repeat-header'),
  )) {
    // gap rows inside a floating table grow only the float's own box
    if (insideFloatTable(el)) continue
    const r = el.getBoundingClientRect()
    // true slice boundary when known: the widget can sit at the flow end
    // while its boundary lies inside trailing float-spill space, which would
    // otherwise pull every below-flow box of the same page onto the next one
    const b = parseFloat(el.dataset.boundaryY ?? '')
    gaps.push({
      v: Number.isFinite(b) ? b : (r.top - origin - acc) / factor,
      h: r.height,
      // only page gaps know their page's header push; float hosts and repeated
      // header rows sit at the same virtual Y and must not clear it
      push: el.dataset.topPush === undefined ? undefined : parseFloat(el.dataset.topPush) || 0,
    })
    acc += r.height
  }
  for (const f of floats) {
    let above = 0
    let pageStart = 0
    let pagePush = firstPagePush
    // page-absolute V boxes render on their ANCHOR's page at the page-relative
    // Y (Word): pinned tops already are page coords, pageRelV tops carry the
    // anchor position. Flow-positioned boxes keep their virtual Y.
    const abs = f.pinned || f.pageRelV
    const ref = abs ? (f.anchorTop ?? f.top) + 0.5 : f.top
    for (const g of gaps) {
      if (g.v <= ref) {
        above += g.h
        if (abs && g.v >= pageStart) {
          pageStart = g.v
          if (g.push !== undefined) pagePush = g.push
        }
      }
    }
    // page-edge V offsets count from the pgMar top, not the header-pushed flow start
    const rel = f.pageRelV ? f.top - (f.anchorTop ?? 0) - (f.pageRelFromPage ? pagePush : 0) : f.top
    const desired = origin + (pageStart + rel) * factor + above
    const applied = parseFloat(f.el.dataset.pageFloatDy ?? '0') || 0
    const cur = f.el.getBoundingClientRect().top
    const next = applied + (desired - cur) / factor
    if (Math.abs(next) < 0.5) {
      f.el.style.removeProperty('--page-float-dy')
      delete f.el.dataset.pageFloatDy
    } else if (Math.abs(next - applied) > 0.5 || !f.el.dataset.pageFloatDy) {
      f.el.style.setProperty('--page-float-dy', `${next.toFixed(1)}px`)
      f.el.dataset.pageFloatDy = String(next)
    }
  }
}

/**
 * Word resumes body text below the union of wrapTopAndBottom bands: consecutive
 * anchor paragraphs (photo walls) stack their own lines, not their bands, so a
 * later anchor's origin sits one line below the previous anchor — not below its
 * whole band. Collapse each non-last wrapper to its anchor line and extend the
 * run's last wrapper so following text still resumes below the band union.
 * Layout-affecting, so it runs before measurement; idempotent (inputs are the
 * static data-band values and the anchor-line heights).
 */
export function syncAnchorBands(pm: HTMLElement, factor: number): void {
  let run: HTMLElement[] = []
  const apply = (el: HTMLElement, minHeight: number): void => {
    const own = Math.round(parseFloat(el.dataset.band ?? '0') || 0)
    if (minHeight === own) {
      if (el.dataset.bandAdj === undefined) return
      delete el.dataset.bandAdj
      el.style.minHeight = own > 0 ? `${own}px` : ''
      return
    }
    if (el.dataset.bandAdj === String(minHeight)) return
    el.dataset.bandAdj = String(minHeight)
    el.style.minHeight = minHeight > 0 ? `${minHeight}px` : ''
  }
  const bandsOf = (el: HTMLElement): Array<[number, number]> =>
    (el.dataset.bands ?? '')
      .split(' ')
      .map((s) => s.split(':').map(Number) as [number, number])
      .filter(([a, b]) => Number.isFinite(a) && Number.isFinite(b) && b > a)
  // union of the accumulated band intervals (photos in one row overlap; their
  // coverage must not double-count)
  const mergedOf = (intervals: Array<[number, number]>): Array<[number, number]> => {
    const sorted = [...intervals].sort((a, b) => a[0] - b[0])
    const out: Array<[number, number]> = []
    for (const [a, b] of sorted) {
      const last = out[out.length - 1]
      if (last && a <= last[1]) last[1] = Math.max(last[1], b)
      else out.push([a, b])
    }
    return out
  }
  // a table-pushed band (data-band-beside) leaves side room: the empty
  // paragraphs between the anchor and the table lay their lines beside the
  // boxes in Word, so their heights come off the band instead of adding below
  let besideBand: HTMLElement | null = null
  let besideEmpties = 0
  const isEmptyParagraph = (el: HTMLElement): boolean =>
    el.tagName === 'P' && !(el.textContent ?? '').trim() && !el.querySelector('img, table')
  const settleBeside = (next: HTMLElement | null): void => {
    if (
      besideBand &&
      besideEmpties > 0 &&
      next &&
      (next.tagName === 'TABLE' || next.querySelector('table'))
    ) {
      const own = Math.round(parseFloat(besideBand.dataset.band ?? '0') || 0)
      const line =
        (besideBand
          .querySelector(':scope > .doc-anchor-strut, :scope > .doc-textbox-stray')
          ?.getBoundingClientRect().height ?? 0) / factor
      apply(besideBand, Math.max(Math.round(line), Math.round(own - besideEmpties)))
    }
    besideBand = null
    besideEmpties = 0
  }
  const flush = (): void => {
    if (run.length > 1) {
      const intervals: Array<[number, number]> = []
      const tops: number[] = []
      let t = 0
      let bottom = 0
      for (const el of run) {
        const line =
          (el
            .querySelector(':scope > .doc-anchor-strut, :scope > .doc-textbox-stray')
            ?.getBoundingClientRect().height ?? 0) / factor
        // the anchor's own line lands on the first slot not substantially
        // covered by earlier bands (Word excludes text lines from wrap bands;
        // the half-line tolerance absorbs our taller-than-Word line boxes)
        const merged = mergedOf(intervals)
        let cand = t
        for (let guard = 0; guard < 64 && line > 0; guard++) {
          const hit = merged.filter(([a, b]) => a < cand + line && b > cand)
          const covered = hit.reduce(
            (s, [a, b]) => s + Math.min(b, cand + line) - Math.max(a, cand),
            0,
          )
          if (covered <= line / 2) break
          cand = Math.min(...hit.map(([, b]) => b))
        }
        tops.push(cand)
        for (const [a, b] of bandsOf(el)) {
          intervals.push([cand + a, cand + b])
          bottom = Math.max(bottom, cand + b)
        }
        t = cand + line
        bottom = Math.max(bottom, t)
      }
      run.forEach((el, i) => {
        const h = i === run.length - 1 ? bottom - tops[i] : tops[i + 1] - tops[i]
        apply(el, Math.max(0, Math.round(h)))
      })
    } else if (run.length === 1) {
      apply(run[0], Math.round(parseFloat(run[0].dataset.band ?? '0') || 0))
    }
    const last = run[run.length - 1]
    if (last?.dataset.bandBeside === '1') besideBand = last
    run = []
  }
  for (const el of Array.from(pm.children) as HTMLElement[]) {
    if (
      el.classList.contains('page-gap') ||
      el.classList.contains('page-float-host') ||
      el.classList.contains('page-float-carry') ||
      el.classList.contains('page-repeat-header')
    ) {
      continue
    }
    if (
      el.classList.contains('doc-protected-floating') &&
      !el.classList.contains('doc-protected-pagepinned')
    ) {
      settleBeside(null)
      run.push(el)
      continue
    }
    flush()
    if (besideBand && isEmptyParagraph(el)) {
      const cs = getComputedStyle(el)
      besideEmpties +=
        el.getBoundingClientRect().height / factor +
        (parseFloat(cs.marginTop) || 0) +
        (parseFloat(cs.marginBottom) || 0)
      continue
    }
    settleBeside(el)
  }
  flush()
  settleBeside(null)
}

/**
 * Word keeps anchored objects on the page: a cell-anchored box whose negative
 * offset lifts it above the paper top (0219: a glossary title box -82pt above
 * a first-row cell paragraph rendered clipped off the paper) is pushed back
 * down to the paper edge. Cell boxes are overlay-only (zero flow footprint),
 * so the shift has no layout feedback. Idempotent via the same
 * --page-float-dy channel as syncFloatShifts.
 */
export function clampCellBoxTops(pm: HTMLElement, paperTop: number, factor: number): void {
  for (const box of Array.from(
    pm.querySelectorAll<HTMLElement>('.doc-cell-boxes > .doc-textbox, .doc-cell-boxes > div'),
  )) {
    const r = box.getBoundingClientRect()
    if (r.height <= 0) continue
    const applied = parseFloat(box.dataset.pageFloatDy ?? '0') || 0
    const naturalTop = r.top - applied * factor
    const next = Math.max(0, (paperTop - naturalTop) / factor)
    if (next < 0.5) {
      if (box.dataset.pageFloatDy) {
        box.style.removeProperty('--page-float-dy')
        delete box.dataset.pageFloatDy
      }
    } else if (Math.abs(next - applied) > 0.5 || !box.dataset.pageFloatDy) {
      box.style.setProperty('--page-float-dy', `${next.toFixed(1)}px`)
      box.dataset.pageFloatDy = String(next)
    }
  }
}

/**
 * In-table page gaps paint their bands inside the spanning cell's
 * .page-gap-table-fill, whose containing block is the cell — origin at the
 * table's left edge, width the table's. An indented, margin-spilling, or
 * narrower-than-paper table therefore shifted the whole footer + gray band +
 * header strip off the paper (public issue #174: a full-width w:tblInd table
 * pushed the band 32px right of the page). Re-anchor each fill to the paper
 * box by measurement (idempotent — the absolute fill has no layout feedback;
 * runs after setPageGaps/setColumnLayout while the widgets' rects are final).
 * Strips and floating images inside the fill then live in paper coordinates,
 * same as block/inline gap boxes.
 */
export function alignTableGapFills(pm: HTMLElement, factor: number): void {
  const fills = pm.querySelectorAll<HTMLElement>('.page-gap-table-fill')
  if (fills.length === 0) return
  // 0.1px-rounded without forced decimals: serialized style values round-trip
  // ('-32.0px' would read back as '-32px' and defeat the dirty checks)
  const px = (v: number) => `${Math.round(v * 10) / 10}px`
  const pmRect = pm.getBoundingClientRect()
  for (const fill of Array.from(fills)) {
    const cell = fill.parentElement
    if (!cell) continue
    // differing-width documents: the fill covers the table's own page, not the paper
    const gap = fill.closest<HTMLElement>('.page-gap')
    const pageX = parseFloat(gap?.style.getPropertyValue('--gap-page-x') ?? '')
    const pageW = parseFloat(gap?.style.getPropertyValue('--gap-page-w') ?? '')
    const onPage = Number.isFinite(pageX) && Number.isFinite(pageW)
    const left = px(
      (pmRect.left - cell.getBoundingClientRect().left) / factor + (onPage ? pageX : 0),
    )
    const width = px(onPage ? pageW : pmRect.width / factor)
    if (fill.style.left !== left) fill.style.left = left
    if (fill.style.width !== width) fill.style.width = width
    // inset:0 from the stylesheet would over-constrain against the explicit width
    if (fill.style.right !== 'auto') fill.style.right = 'auto'
  }
}

/**
 * Differing-width documents: gap header/footer strips live inside gap boxes whose
 * origin shifts with the next section's margins (and, for in-table gaps, with the
 * spanning cell's grid position), so no static left fits every gap kind. Align
 * each strip to the body blocks' left edge by measurement (idempotent; runs after
 * setPageGaps while the widgets' rects are final).
 */
export function alignGapHfStrips(pm: HTMLElement, bodyLeftPx: number, factor: number): void {
  const pmLeft = pm.getBoundingClientRect().left
  const canvasTarget = pmLeft + bodyLeftPx * factor
  for (const el of Array.from(pm.querySelectorAll<HTMLElement>('.page-gap-hf'))) {
    // prefer the strip's own section inset (--hf-ml: the footer above a section
    // break belongs to the PREVIOUS section, the header below it to the next one),
    // then the gap's next-section inset (--gap-ml, makeGapEl): mixed-margin
    // documents must not pin every strip to the first section's column; strips
    // with neither (legacy widget DOM) keep the canvas target
    const own =
      el.style.getPropertyValue('--hf-ml') ||
      el.closest<HTMLElement>('.page-gap')?.style.getPropertyValue('--gap-ml')
    const ownPx = own ? parseFloat(own) : NaN
    const target = Number.isFinite(ownPx) ? pmLeft + ownPx * factor : canvasTarget
    // widget DOM reused from an equal-width era still carries the stylesheet
    // centering (left:50% + translateX(-50%)): pin it before measuring, or the
    // increment is applied against the wrong base
    if (el.style.transform !== 'none') el.style.transform = 'none'
    if (!el.style.left) el.style.left = '0px'
    const delta = (target - el.getBoundingClientRect().left) / factor
    if (Math.abs(delta) < 0.5) continue
    el.style.left = `${((parseFloat(el.style.left) || 0) + delta).toFixed(1)}px`
  }
  // floating header images and footnote areas carry their paper x; an image's own
  // rect includes the anchor translate, so re-anchor from the positioned host's origin
  for (const el of Array.from(pm.querySelectorAll<HTMLElement>('.page-gap [data-paper-x]'))) {
    const host = el.offsetParent
    if (!host) continue
    const left = (
      parseFloat(el.dataset.paperX!) -
      (host.getBoundingClientRect().left - pmLeft) / factor
    ).toFixed(1)
    if (el.style.left !== `${left}px`) el.style.left = `${left}px`
  }
}

/** remove all float display shifts (leaving print view) */
export function clearFloatShifts(pm: HTMLElement): void {
  for (const el of Array.from(pm.querySelectorAll<HTMLElement>('[data-page-float-dy]'))) {
    el.style.removeProperty('--page-float-dy')
    delete el.dataset.pageFloatDy
  }
}
