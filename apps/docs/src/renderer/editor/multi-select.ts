import { Extension } from '@tiptap/core'
import { Fragment, Slice, type Mark, type MarkType, type Node } from '@tiptap/pm/model'
import {
  Plugin,
  PluginKey,
  TextSelection,
  type EditorState,
  type Selection,
  type Transaction,
} from '@tiptap/pm/state'
import { AddMarkStep, RemoveMarkStep, type Mapping } from '@tiptap/pm/transform'
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
 *
 * The added ranges take part in three operations: a mark change applied to the
 * live selection (bold, font, color, … from the ribbon, a shortcut or a dialog)
 * is replayed on every added range by appendTransaction; copy/cut put all
 * ranges on the clipboard; Backspace/Delete and cut remove them all.
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

type MultiSelectMeta = { add: readonly MultiSelectRange[] } | { clear: true } | { replay: true }

export interface MultiSelectOptions {
  /** class painted on an added range; matches the app's selection styling */
  readonly className: string
}

const mac = typeof navigator !== 'undefined' && /mac/i.test(navigator.platform)

/** the platform's ⌘/ctrl modifier (ctrl+click is the right click on a Mac), without alt/shift */
export function isMultiSelectModifier(e: MouseEvent): boolean {
  return (mac ? e.metaKey : e.ctrlKey) && !e.altKey && !e.shiftKey
}

/** ⌘/ctrl+click on a link or TOC entry already means "follow it" in Docs */
function isNavigationTarget(e: MouseEvent): boolean {
  return e.target instanceof Element && e.target.closest('a[href], .doc-toc-line') !== null
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

/**
 * The click's own range is the word under the cursor, which becomes the live
 * selection; the range the user had before joins the added ranges (a
 * plain-text selection only — a node selection is not a text range).
 */
function addClickedWord(view: EditorView, pos: number, before: Selection | null): void {
  const ranges: MultiSelectRange[] = []
  if (before instanceof TextSelection && before.from < before.to) {
    ranges.push({ from: before.from, to: before.to })
  }
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
}

/** every selected range — the added ones plus the live selection — sorted and merged */
export function allSelectedRanges(state: EditorState): MultiSelectRange[] {
  const { selection } = state
  const extra = multiSelectKey.getState(state)?.ranges ?? []
  if (extra.length === 0) return []
  const list = [...extra]
  if (selection instanceof TextSelection && !selection.empty) {
    list.push({ from: selection.from, to: selection.to })
  }
  list.sort((a, b) => a.from - b.from)
  const merged: MultiSelectRange[] = []
  for (const r of list) {
    const last = merged[merged.length - 1]
    if (last && r.from <= last.to) {
      if (r.to > last.to) merged[merged.length - 1] = { from: last.from, to: r.to }
    } else merged.push(r)
  }
  return merged
}

/**
 * The attrs the command set: identical across every AddMarkStep of the type
 * and different from the mark that sat there before. Attrs that vary between
 * steps are the live range's own formatting carried along by setMark's
 * per-run merge, and replaying those would overwrite the other ranges.
 */
function markPatch(oldDoc: Node, steps: readonly AddMarkStep[]): Record<string, unknown> {
  const patch: Record<string, unknown> = {}
  const first = steps[0].mark
  for (const [k, v] of Object.entries(first.attrs)) {
    if (!steps.every((st) => st.mark.attrs[k] === v)) continue
    const changed = steps.some((st) => {
      const prev =
        oldDoc.nodeAt(st.from)?.marks.find((m) => m.type === first.type) ?? first.type.create()
      return prev.attrs[k] !== v
    })
    if (changed) patch[k] = v
  }
  return patch
}

function addMarkPatched(
  tr: Transaction,
  range: MultiSelectRange,
  type: MarkType,
  patch: Record<string, unknown>,
): void {
  if (Object.keys(type.spec.attrs ?? {}).length === 0) {
    tr.addMark(range.from, range.to, type.create())
    return
  }
  const runs: { from: number; to: number; mark: Mark }[] = []
  tr.doc.nodesBetween(range.from, range.to, (node, pos) => {
    if (!node.isInline) return
    const prev = node.marks.find((m) => m.type === type)
    runs.push({
      from: Math.max(pos, range.from),
      to: Math.min(pos + node.nodeSize, range.to),
      mark: type.create({ ...prev?.attrs, ...patch }),
    })
  })
  for (const r of runs) tr.addMark(r.from, r.to, r.mark)
}

/** the mark steps of `trs`, or null when they are not a pure mark change of the live selection */
function markStepsOnSelection(
  trs: readonly Transaction[],
  oldState: EditorState,
): (AddMarkStep | RemoveMarkStep)[] | null {
  const { selection } = oldState
  if (selection.empty) return null
  const steps: (AddMarkStep | RemoveMarkStep)[] = []
  for (const tr of trs) {
    for (const step of tr.steps) {
      if (!(step instanceof AddMarkStep || step instanceof RemoveMarkStep)) return null
      if (step.from < selection.from || step.to > selection.to) return null
      steps.push(step)
    }
  }
  return steps.length > 0 ? steps : null
}

// formatting only: a comment, link or revision anchors one place in the document
const REPLAYED_MARKS = new Set(['bold', 'italic', 'underline', 'strike', 'docTextStyle'])

/** replay the mark changes of `transactions` on the added ranges, or null when there is nothing to do */
function replayMarks(
  transactions: readonly Transaction[],
  oldState: EditorState,
  newState: EditorState,
): Transaction | null {
  const extra = multiSelectKey.getState(newState)?.ranges ?? []
  if (extra.length === 0) return null
  if (transactions.some((t) => t.getMeta(multiSelectKey))) return null
  const steps = markStepsOnSelection(transactions, oldState)
  if (!steps) return null
  const live = oldState.selection
  const targets = extra.filter((r) => r.to <= live.from || r.from >= live.to)
  if (targets.length === 0) return null
  const tr = newState.tr.setMeta(multiSelectKey, { replay: true })
  const adds = new Map<MarkType, AddMarkStep[]>()
  const removes = new Set<MarkType>()
  for (const step of steps) {
    if (!REPLAYED_MARKS.has(step.mark.type.name)) continue
    if (step instanceof AddMarkStep)
      adds.set(step.mark.type, [...(adds.get(step.mark.type) ?? []), step])
    else removes.add(step.mark.type)
  }
  // setMark rewrites a run as remove+add; only a lone remove is an unset
  for (const type of removes) {
    if (adds.has(type)) continue
    for (const r of targets) tr.removeMark(r.from, r.to, type)
  }
  for (const [type, typeSteps] of adds) {
    const patch = markPatch(oldState.doc, typeSteps)
    for (const r of targets) addMarkPatched(tr, r, type, patch)
  }
  return tr.steps.length > 0 ? tr : null
}

/** one block per range so the clipboard HTML keeps the runs' formatting */
function clipboardSlice(state: EditorState, ranges: readonly MultiSelectRange[]): Slice {
  const blocks: Node[] = []
  for (const r of ranges) {
    const $from = state.doc.resolve(r.from)
    const slice = state.doc.slice(r.from, r.to)
    if ($from.parent.isTextblock && $from.sameParent(state.doc.resolve(r.to))) {
      blocks.push($from.parent.copy(slice.content))
    } else slice.content.forEach((n) => blocks.push(n))
  }
  return new Slice(Fragment.from(blocks), 0, 0)
}

function deleteRanges(view: EditorView, ranges: readonly MultiSelectRange[]): void {
  const tr = view.state.tr
  for (let i = ranges.length - 1; i >= 0; i--) tr.delete(ranges[i].from, ranges[i].to)
  view.dispatch(tr.setMeta(multiSelectKey, { clear: true }).scrollIntoView())
}

function copyRanges(view: EditorView, event: ClipboardEvent, cut: boolean): boolean {
  const ranges = allSelectedRanges(view.state)
  if (ranges.length === 0 || !event.clipboardData) return false
  const slice = clipboardSlice(view.state, ranges)
  const { dom } = view.serializeForClipboard(slice)
  event.preventDefault()
  event.clipboardData.clearData()
  event.clipboardData.setData('text/html', dom.innerHTML)
  // tiptap's text serializer reads the live selection, not the slice
  event.clipboardData.setData(
    'text/plain',
    slice.content.textBetween(0, slice.content.size, '\n\n'),
  )
  if (cut && view.editable) deleteRanges(view, ranges)
  return true
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

    // runs after ProseMirror's own mouseup: a click has already gone through
    // handleClick (which resets the gesture), so only a drag is left here
    const onGestureEnd = (view: EditorView, e: MouseEvent) => {
      const before = beforeGesture
      const start = gestureStart
      beforeGesture = null
      gestureStart = null
      if (!before || !start || e.button !== 0 || !isMultiSelectModifier(e)) return
      const dragged = Math.abs(e.clientX - start.x) > 4 || Math.abs(e.clientY - start.y) > 4
      if (!dragged) return
      const { from, to } = view.state.selection
      if (from < to) {
        addRanges(view, [
          { from: before.from, to: before.to },
          { from, to },
        ])
        return
      }
      // moved past the threshold but dragged no text (pointer jitter): a click
      const hit = view.posAtCoords({ left: e.clientX, top: e.clientY })
      if (hit) addClickedWord(view, hit.pos, before)
    }

    return [
      new Plugin<MultiSelectState>({
        key: multiSelectKey,
        state: {
          init: () => EMPTY,
          apply(tr, value) {
            const meta = tr.getMeta(multiSelectKey) as MultiSelectMeta | undefined
            if (meta && 'clear' in meta) return EMPTY
            const mapped = mapRanges(value.ranges, tr.mapping)
            if (meta && 'add' in meta) return { ranges: [...mapped, ...meta.add] }
            // nothing changed and the text did not move: keep the identical
            // state object so decorations are not rebuilt on every keystroke
            return mapped.length === value.ranges.length &&
              mapped.every((r, i) => r.from === value.ranges[i].from && r.to === value.ranges[i].to)
              ? value
              : { ranges: mapped }
          },
        },
        appendTransaction: replayMarks,
        props: {
          decorations(state) {
            const ranges = multiSelectKey.getState(state)?.ranges ?? []
            if (ranges.length === 0) return DecorationSet.empty
            return DecorationSet.create(
              state.doc,
              ranges.map((r) => Decoration.inline(r.from, r.to, { class: className })),
            )
          },
          handleKeyDown: (view, e) => {
            if ((e.key !== 'Backspace' && e.key !== 'Delete') || e.metaKey || e.ctrlKey || e.altKey)
              return false
            const ranges = allSelectedRanges(view.state)
            if (ranges.length === 0 || !view.editable) return false
            deleteRanges(view, ranges)
            return true
          },
          handleDOMEvents: {
            copy: (view, event) => copyRanges(view, event as ClipboardEvent, false),
            cut: (view, event) => copyRanges(view, event as ClipboardEvent, true),
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
              if (isNavigationTarget(e)) return false
              beforeGesture = view.state.selection
              gestureStart = { x: e.clientX, y: e.clientY }
              // the release may land outside the editor (sidebar, window
              // chrome), where the editor's own mouseup never fires
              window.addEventListener('mouseup', (up) => onGestureEnd(view, up), { once: true })
              return false
            },
          },
          handleClick: (view, pos, event) => {
            const e = event as MouseEvent
            if (e.button !== 0 || !isMultiSelectModifier(e) || isNavigationTarget(e)) return false
            const before = beforeGesture
            beforeGesture = null
            gestureStart = null
            addClickedWord(view, pos, before)
            // never let PM's selectClickedNode turn the paragraph into a NodeSelection
            return true
          },
        },
      }),
    ]
  },
})
