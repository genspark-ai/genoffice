import { Extension } from '@tiptap/core'
import { Plugin, PluginKey, TextSelection, type Selection } from '@tiptap/pm/state'
import type { Mapping } from '@tiptap/pm/transform'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import type { EditorView } from '@tiptap/pm/view'
import { wordSegmentAt } from './word-range'

/**
 * Ctrl/⌘+click adds a non-contiguous selection range, the way OnlyOffice and
 * browsers do. Asked for in genoffice#2013 ("requested in OnlyOffice for many
 * years").
 *
 * ProseMirror's Selection is a single `{from, to}` pair, so the added ranges
 * live in plugin state and are painted as decorations. The live selection stays
 * the one ProseMirror owns: the caret, typing and every existing op keep
 * working on it unchanged. This deliberately does NOT implement multi-caret
 * editing — that means rewriting the core's insertion path, not a plugin.
 *
 * The click is captured on the mouseup side (handleClick), not from a task
 * queued by the mousedown handler: a modifier click leaves the browser
 * selection untouched, so at mouseup the state still holds the range the user
 * had before the click — that one joins the added ranges, and the click's own
 * range is the word under the cursor, which becomes the live selection.
 * Returning true from handleClick suppresses ProseMirror's own
 * selectNodeModifier path, which would otherwise turn the whole paragraph into
 * a NodeSelection. A plain mousedown clears the added ranges, and a modifier
 * drag adds both the pre-drag range and the dragged one.
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

/** append the ranges that are not added yet; a re-click must not stack copies */
function addRanges(view: EditorView, ranges: readonly MultiSelectRange[]): void {
  const add: MultiSelectRange[] = []
  for (const r of ranges) {
    if (r.from >= r.to) continue
    const seen = (list: readonly MultiSelectRange[]) =>
      list.some((e) => e.from === r.from && e.to === r.to)
    if (seen(multiSelectRanges(view)) || seen(add)) continue
    add.push(r)
  }
  if (add.length === 0) return
  view.dispatch(view.state.tr.setMeta(multiSelectKey, { add }))
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

/**
 * The word the click landed on, by the app's own word segmentation — the
 * shared `Intl.Segmenter`-based `wordSegmentAt`, the base unit the format
 * painter's word range is built from. Null when the click is outside a text
 * block or on no word — then the click adds nothing of its own.
 */
function wordRangeAt(view: EditorView, pos: number): MultiSelectRange | null {
  const { doc } = view.state
  if (pos < 0 || pos > doc.content.size) return null
  const $pos = doc.resolve(pos)
  if (!$pos.parent.isTextblock) return null
  // NUL stands in for leaf nodes: never word-like, so offsets stay aligned
  const text = $pos.parent.textBetween(0, $pos.parent.content.size, '\0', '\0')
  const seg = wordSegmentAt(text, $pos.parentOffset)
  if (!seg) return null
  const base = $pos.start()
  return { from: base + seg.start, to: base + seg.end }
}

export const MultiSelect = Extension.create<MultiSelectOptions>({
  name: 'multiSelect',

  addOptions() {
    return { className: 'doc-multi-range' }
  },

  addProseMirrorPlugins() {
    const { className } = this.options
    // the selection a modifier gesture started from: a modifier click leaves
    // the browser selection alone and a modifier drag replaces it mid-drag, so
    // the pre-gesture range can only be kept from the mousedown side
    let beforeGesture: Selection | null = null
    let gestureStart: { x: number; y: number } | null = null

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
              if (e.button !== 0) return false
              if (!isMultiSelectModifier(e)) {
                // a plain click starts a fresh selection: the added ranges go
                beforeGesture = null
                gestureStart = null
                clearMultiSelect(view)
                return false
              }
              beforeGesture = view.state.selection
              gestureStart = { x: e.clientX, y: e.clientY }
              return false
            },
            mouseup: (view, event) => {
              const e = event as MouseEvent
              if (e.button !== 0 || !isMultiSelectModifier(e)) return false
              const before = beforeGesture
              if (!before || !gestureStart) return false
              // ProseMirror's own drag threshold: past it this is a drag, and
              // the browser has already replaced the selection with the range
              // it dragged, which handleClick never sees
              const dragged =
                Math.abs(e.clientX - gestureStart.x) > 4 || Math.abs(e.clientY - gestureStart.y) > 4
              if (!dragged) return false // a click: handleClick does the work
              const { from, to } = view.state.selection
              addRanges(view, [
                { from: before.from, to: before.to },
                { from, to },
              ])
              beforeGesture = null
              gestureStart = null
              return false
            },
          },
          handleClick: (view, pos, event) => {
            const e = event as MouseEvent
            if (e.button !== 0 || !isMultiSelectModifier(e)) return false
            const before = beforeGesture
            beforeGesture = null
            gestureStart = null
            // mouseup side: the range the user had before this click joins the
            // added ranges (a plain-text selection only — a node selection is
            // not a text range)...
            const ranges: MultiSelectRange[] = []
            if (before instanceof TextSelection && before.from < before.to) {
              ranges.push({ from: before.from, to: before.to })
            }
            // ...and the click's own range is the word under the cursor, which
            // becomes the live selection
            const word = wordRangeAt(view, pos)
            if (word) ranges.push(word)
            addRanges(view, ranges)
            if (word) {
              view.dispatch(
                view.state.tr
                  .setSelection(TextSelection.create(view.state.doc, word.from, word.to))
                  .setMeta('pointer', true),
              )
            }
            // never let PM's selectClickedNode turn the paragraph into a NodeSelection
            return true
          },
        },
      }),
    ]
  },
})
