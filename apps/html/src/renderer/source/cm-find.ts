import { SearchQuery } from '@codemirror/search'
import { StateEffect, StateField, type Text } from '@codemirror/state'
import { Decoration, EditorView, type DecorationSet } from '@codemirror/view'
import type { FindOptions } from '@genoffice/ui'

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
