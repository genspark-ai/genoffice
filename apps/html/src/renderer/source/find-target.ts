import { EditorView } from '@codemirror/view'
import type { FindTarget } from '@genoffice/ui'
import { collectMatches, setFindHits, type FindRange } from './cm-find'
import { External } from './cm-setup'

/**
 * Find panel adapter over the source editor. Replacements are plain
 * transactions, so they reach the document buffer like typed edits.
 */
export function cmFindTarget(
  view: EditorView,
  onDocChanged: (listener: () => void) => () => void,
): FindTarget {
  let ranges: FindRange[] = []
  const paint = (active: number) => view.dispatch({ effects: setFindHits.of({ ranges, active }) })
  return {
    get editable() {
      return !view.state.readOnly
    },
    search(query, opts, activeIndex) {
      ranges = collectMatches(view.state.doc, query, opts)
      paint(ranges.length === 0 ? 0 : Math.min(activeIndex, ranges.length - 1))
      return ranges.length
    },
    activate(index) {
      const r = ranges[index]
      if (!r) return
      view.dispatch({
        selection: { anchor: r.from, head: r.to },
        effects: [
          setFindHits.of({ ranges, active: index }),
          EditorView.scrollIntoView(r.from, { y: 'center' }),
        ],
        annotations: [External.of(true)],
      })
    },
    replaceOne(index, replacement) {
      const r = ranges[index]
      if (!r) return
      view.dispatch({
        changes: { from: r.from, to: r.to, insert: replacement },
        userEvent: 'input.replace',
      })
    },
    replaceAll(replacement) {
      if (ranges.length === 0) return
      view.dispatch({
        changes: ranges.map((r) => ({ from: r.from, to: r.to, insert: replacement })),
        userEvent: 'input.replace.all',
      })
    },
    clear() {
      ranges = []
      paint(0)
    },
    onDocChanged,
  }
}
