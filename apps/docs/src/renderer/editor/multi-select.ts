import { Extension } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import type { Mapping } from '@tiptap/pm/transform'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import type { EditorView } from '@tiptap/pm/view'

/**
 * Ctrl/⌘+click adds a non-contiguous selection range, the way OnlyOffice and
 * browsers do. Asked for in #2013 ("requested in OnlyOffice for many years").
 *
 * ProseMirror's Selection is a single `{from, to}` pair, so the added ranges
 * live in plugin state and are painted as decorations. The live selection stays
 * the one ProseMirror owns: the caret, typing and every existing op keep
 * working on it unchanged. This deliberately does NOT implement multi-caret
 * editing — that means rewriting the core's insertion path, not a plugin.
 */

/** one added range, in document coordinates */
export interface MultiSelectRange {
  readonly from: number
  readonly to: number
}

interface MultiSelectState {
  /** ranges added by ctrl/cmd+click, in the order they were added */
  readonly ranges: readonly MultiSelectRange[]
}

const multiSelectKey = new PluginKey<MultiSelectState>('docsMultiSelect')

const EMPTY: MultiSelectState = { ranges: [] }

type MultiSelectMeta = { add: readonly MultiSelectRange[] } | { clear: true }

export interface MultiSelectOptions {
  /** class painted on an added range; matches the app's selection styling */
  readonly className: string
}

/** the ⌘/ctrl modifier, without the alt/shift combos that mean something else */
export function isMultiSelectModifier(e: MouseEvent): boolean {
  return (e.metaKey || e.ctrlKey) && !e.altKey && !e.shiftKey
}

/** the ranges currently added, [] when there are none */
export function multiSelectRanges(view: EditorView): readonly MultiSelectRange[] {
  return multiSelectKey.getState(view.state)?.ranges ?? []
}

/** drop every added range */
export function clearMultiSelect(view: EditorView): void {
  if (multiSelectRanges(view).length === 0) return
  view.dispatch(view.state.tr.setMeta(multiSelectKey, { clear: true }))
}

/**
 * Ranges survive editing: text inserted before one shifts it, and a range the
 * edit swallowed is dropped rather than highlighted as a zero-width sliver.
 */
function mapRanges(ranges: readonly MultiSelectRange[], mapping: Mapping): MultiSelectRange[] {
  return ranges
    .map((r) => ({ from: mapping.map(r.from, 1), to: mapping.map(r.to, -1) }))
    .filter((r) => r.from < r.to)
}

export const MultiSelect = Extension.create<MultiSelectOptions>({
  name: 'multiSelect',

  addOptions() {
    return { className: 'doc-multi-range' }
  },

  addProseMirrorPlugins() {
    const { className } = this.options
    return [
      new Plugin<MultiSelectState>({
        key: multiSelectKey,
        state: {
          init: () => EMPTY,
          apply(tr, value) {
            const meta = tr.getMeta(multiSelectKey) as MultiSelectMeta | undefined
            if (meta && 'clear' in meta) return EMPTY
            const mapped = mapRanges(value.ranges, tr.mapping)
            if (meta) return { ranges: [...mapped, ...meta.add] }
            // nothing changed and the text did not move: keep the identical
            // state object so decorations are not rebuilt on every keystroke
            return mapped.length === value.ranges.length &&
              mapped.every((r, i) => r.from === value.ranges[i].from && r.to === value.ranges[i].to)
              ? value
              : { ranges: mapped }
          },
        },
        props: {
          decorations(state) {
            const ranges = multiSelectKey.getState(state)?.ranges ?? []
            if (ranges.length === 0) return DecorationSet.empty
            return DecorationSet.create(
              state.doc,
              ranges.map((r) => Decoration.inline(r.from, r.to, { class: className })),
            )
          },
          handleDOMEvents: {
            mousedown: (view, event) => {
              const e = event as MouseEvent
              if (!isMultiSelectModifier(e)) return false
              // Return false so ProseMirror places its own caret first; this
              // runs after that and takes whatever range the click produced. A
              // collapsed click means nothing was selected — skip it instead of
              // storing a zero-width range.
              window.setTimeout(() => {
                if (view.isDestroyed) return
                const { from, to } = view.state.selection
                if (from === to) return
                view.dispatch(view.state.tr.setMeta(multiSelectKey, { add: [{ from, to }] }))
              }, 0)
              return false
            },
          },
        },
      }),
    ]
  },
})
