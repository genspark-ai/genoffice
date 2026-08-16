import { Extension } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import type { EditorView } from '@tiptap/pm/view'
import type { LineAnchor } from '../pagination'

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

export interface GapMetrics {
  marginTop: number
  marginBottom: number
  marginLeft: number
  marginRight: number
}

/** height of the gray inter-page band inside a page gap */
export const GAP_BAND = 28

export type GapKind = 'block' | 'inline' | 'table' | 'cut' | 'cell'

export function makeGapEl(m: GapMetrics, kind: GapKind): HTMLElement {
  const gap = document.createElement(kind === 'table' ? 'tr' : 'div')
  gap.contentEditable = 'false'
  if (kind === 'cut') {
    // in-row table break cut point: zero-height dashed marker (takes no layout height, coordinate bookkeeping unchanged)
    gap.className = 'page-gap-cut'
    return gap
  }
  gap.style.height = `${m.marginBottom + GAP_BAND + m.marginTop}px`
  gap.style.setProperty('--gap-mb', `${m.marginBottom}px`)
  gap.style.setProperty('--gap-mt', `${m.marginTop}px`)
  if (kind === 'table') {
    // A real spanning cell is required here. Chromium's collapsed-border table
    // painting can leak the neighboring row's border/fill through a cell-less
    // display:table-row, leaving a colored remnant in the gray page gutter.
    gap.className = 'page-gap page-gap-inline page-gap-table'
    const cell = document.createElement('td')
    cell.colSpan = 1000
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
} & ({ el: HTMLElement } | { pos: number; kind?: Exclude<GapKind, 'block'> })

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
    const mKey = `${metrics.marginTop},${metrics.marginBottom},${metrics.marginLeft},${metrics.marginRight}`
    decos.push(
      Decoration.widget(
        pos,
        () => {
          const el = makeGapEl(metrics, kind)
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
  if (a.node.length > 0) {
    const range = document.createRange()
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
  floats: Array<{ el: HTMLElement; top: number }>,
  origin: number,
  factor: number,
): void {
  if (floats.length === 0) return
  const gaps: Array<{ v: number; h: number }> = []
  let acc = 0
  for (const el of Array.from(pm.children) as HTMLElement[]) {
    if (
      el.classList.contains('page-gap') ||
      el.classList.contains('page-float-host') ||
      el.classList.contains('page-repeat-header')
    ) {
      const r = el.getBoundingClientRect()
      gaps.push({ v: (r.top - origin - acc) / factor, h: r.height })
      acc += r.height
    }
  }
  for (const f of floats) {
    let above = 0
    for (const g of gaps) if (g.v <= f.top) above += g.h
    const desired = origin + f.top * factor + above
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

/** remove all float display shifts (leaving print view) */
export function clearFloatShifts(pm: HTMLElement): void {
  for (const el of Array.from(pm.querySelectorAll<HTMLElement>('[data-page-float-dy]'))) {
    el.style.removeProperty('--page-float-dy')
    delete el.dataset.pageFloatDy
  }
}
