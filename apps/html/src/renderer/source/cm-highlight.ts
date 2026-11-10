import { StateEffect, StateField, type Extension } from '@codemirror/state'
import { Decoration, EditorView, type DecorationSet } from '@codemirror/view'

/** transient wash over ranges the AI just wrote; mapped through later edits, cleared on demand */
export const addAiRanges = StateEffect.define<Array<[number, number]>>({
  map: (ranges, change) => ranges.map(([a, b]) => [change.mapPos(a), change.mapPos(b, 1)]),
})
export const clearAiRanges = StateEffect.define<null>()

const mark = Decoration.mark({ class: 'cm-ai-changed' })

export const aiHighlightField = StateField.define<DecorationSet>({
  create: () => Decoration.none,
  update(set, tr) {
    let next = set.map(tr.changes)
    for (const effect of tr.effects) {
      if (effect.is(clearAiRanges)) next = Decoration.none
      else if (effect.is(addAiRanges)) {
        const ranges = effect.value
          .filter(([a, b]) => b > a)
          .sort((x, y) => x[0] - y[0])
          .map(([a, b]) => mark.range(a, b))
        if (ranges.length) next = next.update({ add: ranges, sort: true })
      }
    }
    return next
  },
  provide: (f) => EditorView.decorations.from(f),
})

export function aiHighlight(): Extension {
  return [aiHighlightField]
}
