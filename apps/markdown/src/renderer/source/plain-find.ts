import { SearchQuery } from '@codemirror/search'
import { StateEffect, StateField, type Text } from '@codemirror/state'
import { Decoration, EditorView, type DecorationSet } from '@codemirror/view'
import type { FindOptions, FindTarget } from '@genoffice/ui'

export interface FindRange {
  from: number
  to: number
}

export const setFindHits = StateEffect.define<{ ranges: FindRange[]; active: number }>()

const hit = Decoration.mark({ class: 'search-hit' })
const activeHit = Decoration.mark({ class: 'search-hit search-hit-active' })

/** find hits painted by the app's own panel (CodeMirror's search panel is not used) */
export const findHighlight = StateField.define<DecorationSet>({
  create: () => Decoration.none,
  update(deco, tr) {
    deco = deco.map(tr.changes)
    for (const e of tr.effects) {
      if (e.is(setFindHits)) {
        deco = Decoration.set(
          e.value.ranges.map((r, i) =>
            (i === e.value.active ? activeHit : hit).range(r.from, r.to),
          ),
          true,
        )
      }
    }
    return deco
  },
  provide: (f) => EditorView.decorations.from(f),
})

export function collectMatches(doc: Text, query: string, opts: FindOptions): FindRange[] {
  if (!query) return []
  const q = new SearchQuery({
    search: query,
    caseSensitive: opts.matchCase,
    wholeWord: opts.wholeWord,
    literal: true,
  })
  const out: FindRange[] = []
  const cursor = q.getCursor(doc)
  for (let m = cursor.next(); !m.done; m = cursor.next()) {
    out.push({ from: m.value.from, to: m.value.to })
  }
  return out
}

/**
 * Find panel adapter over the source editor. Replacements are plain
 * transactions, so they reach the buffer like typed edits.
 */
export function buildFindTarget(
  view: EditorView,
  onDocChanged: (listener: () => void) => () => void,
): FindTarget {
  let ranges: FindRange[] = []
  let query = ''
  let options: FindOptions = { matchCase: false, wholeWord: false }
  let active = 0
  const paint = () => view.dispatch({ effects: setFindHits.of({ ranges, active }) })
  const scan = () => {
    ranges = collectMatches(view.state.doc, query, options)
    active = ranges.length === 0 ? 0 : Math.min(active, ranges.length - 1)
    paint()
    return ranges.length
  }
  return {
    get editable() {
      return !view.state.readOnly
    },
    search(nextQuery, nextOptions, activeIndex) {
      query = nextQuery
      options = nextOptions
      active = activeIndex
      return scan()
    },
    activate(index) {
      const r = ranges[index]
      if (!r) return
      active = index
      view.dispatch({
        selection: { anchor: r.from, head: r.to },
        effects: [
          setFindHits.of({ ranges, active }),
          EditorView.scrollIntoView(r.from, { y: 'center' }),
        ],
      })
      view.focus()
    },
    replaceOne(index, replacement) {
      if (ranges.length === 0) return
      // positions move once a replacement lands, so re-scan before indexing
      scan()
      const r = ranges[index]
      if (!r) return
      view.dispatch({
        changes: { from: r.from, to: r.to, insert: replacement },
        userEvent: 'input.replace',
      })
    },
    replaceAll(replacement) {
      if (ranges.length === 0) return
      if (scan() === 0) return
      view.dispatch({
        changes: ranges.map((r) => ({ from: r.from, to: r.to, insert: replacement })),
        userEvent: 'input.replace.all',
      })
    },
    clear() {
      ranges = []
      active = 0
      paint()
    },
    onDocChanged,
  }
}
