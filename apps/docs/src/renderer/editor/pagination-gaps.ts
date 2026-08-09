import { Extension } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import type { EditorView } from '@tiptap/pm/view'

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
} & ({ el: HTMLElement } | { pos: number; kind?: Exclude<GapKind, 'block'> })

/** Rebuild all page gaps (an empty list clears them); each gap carries its own margins (sections differ) */
export function setPageGaps(view: EditorView, gaps: PageGapSpec[]): void {
  const decos: Decoration[] = []
  let ordinal = -1
  for (const gap of gaps) {
    ordinal++
    const { metrics, notes, hfEls } = gap
    let pos: number
    let kind: GapKind
    if ('el' in gap) {
      kind = 'block'
      try {
        const inside = view.posAtDOM(gap.el, 0)
        pos = view.state.doc.resolve(inside).before(1)
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
  if (prev && sameGaps(prev, next)) return
  view.dispatch(view.state.tr.setMeta(key, next).setMeta('addToHistory', false))
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
