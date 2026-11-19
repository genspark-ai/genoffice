import { EditorView } from '@codemirror/view'
import type { FindOptions, FindTarget } from '@genoffice/ui'
import { collectMatches, setFindHits, type FindRange } from './cm-find'
import { External } from './cm-setup'

/**
 * Find panel adapter over the source editor. Replacements are plain
 * transactions, so they reach the document buffer like typed edits.
 */
export function cmFindTarget(
  view: EditorView,
  onDocChanged: (listener: () => void) => () => void,
  onBeforeReplace?: () => void,
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
        annotations: [External.of(true)],
      })
    },
    replaceOne(index, replacement) {
      if (ranges.length === 0) return
      onBeforeReplace?.()
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
      onBeforeReplace?.()
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
