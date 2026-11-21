import { Fragment, Slice, type Node as PmNode } from '@tiptap/pm/model'
import type { EditorState, Transaction } from '@tiptap/pm/state'
import { NodeSelection, Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view'
import type { TextboxDisplay } from '@genoffice/docx-engine'
import { shapeWrapOf } from './floating-z-order'

/** what a moved shape's offsets measure from, per axis */
export interface OffsetFrame {
  h: 'page' | 'margin' | 'column'
  v: 'page' | 'margin' | 'paragraph'
}

/** a group member's box inside its group, px */
export interface GroupBox {
  x: number
  y: number
  w: number
  h: number
}

/** what a child of a group taken apart is saved from, in place of a block of the file */
export interface GroupMemberSource {
  xml: string
  /** the boxes as `xml` holds them: what an edit is told from */
  boxes: TextboxDisplay[]
  /** the anchor's offsets, EMU */
  xEmu: number
  yEmu: number
}

/**
 * Several floating objects selected at once (Word: Shift+click). The node
 * selection stays the primary object, so everything that edits "the selected
 * object" keeps working; the others ride along here.
 */
interface ObjectSelectState {
  primary: number
  extras: number[]
  /**
   * The other objects of a drag in progress. A drop can move the dragged
   * object's node, which ends the selection; these keep following the
   * document so the drop can take them along.
   */
  held: number[]
}

const key = new PluginKey<ObjectSelectState | null>('docObjectMultiSelect')

/** the boxes of a group read from the file: one anchor places them all */
export const isFileGroup = (boxes: TextboxDisplay[]): boolean =>
  boxes.length > 0 && boxes.every((box) => box.grouped)

/** a picture, shape, text box or group that floats: the only objects Word selects together */
export function isFloatingObject(node: PmNode | null | undefined): boolean {
  if (!node) return false
  if (node.type.name === 'docInlineImage') return !!node.attrs.wrap && !node.attrs.rule
  if (node.type.name !== 'docProtected') return false
  const boxes = node.attrs.textboxes as TextboxDisplay[] | null
  // drawings sharing a paragraph each keep their own anchor: the node is not one object
  if (Array.isArray(boxes) && boxes.length > 0)
    return (boxes.length === 1 || isFileGroup(boxes)) && shapeWrapOf(boxes[0]) !== null
  return node.attrs.blockType === 'image' && !!node.attrs.imageWrap
}

function primaryOf(state: EditorState): number | null {
  const sel = state.selection
  return sel instanceof NodeSelection && isFloatingObject(sel.node) ? sel.from : null
}

/** the selected floating objects, the primary one first; empty when the selection is not on one */
export function selectedObjectPositions(state: EditorState): number[] {
  const primary = primaryOf(state)
  if (primary === null) return []
  const s = key.getState(state)
  return s && s.primary === primary ? [primary, ...s.extras] : [primary]
}

/** Selects `positions` together, the first as the primary object. */
export function setObjectSelection(tr: Transaction, positions: number[]): Transaction {
  const [primary, ...extras] = positions
  if (primary === undefined) return tr.setMeta(key, null)
  return tr
    .setSelection(NodeSelection.create(tr.doc, primary))
    .setMeta(key, extras.length > 0 ? { primary, extras, held: [] } : null)
}

/**
 * Shift+click on an object: adds it to the selection, or takes it out when it
 * is already there. False when the object stands alone and stays selected.
 */
export function toggleObjectSelection(view: EditorView, pos: number): boolean {
  const { state } = view
  if (!isFloatingObject(state.doc.nodeAt(pos))) return false
  const current = selectedObjectPositions(state)
  // a group goes in or comes out whole
  const unit = [pos, ...groupMates(state.doc, [pos]).filter((p) => p !== pos)]
  const next = current.includes(pos)
    ? current.filter((p) => !unit.includes(p))
    : [...current, ...unit.filter((p) => !current.includes(p))]
  if (next.length === 0) return false
  view.dispatch(setObjectSelection(state.tr, next))
  view.focus()
  return true
}

/** Keeps the objects selected with the primary one through the drag that is starting. */
export function holdObjectSelection(view: EditorView): void {
  const s = key.getState(view.state)
  if (s && s.extras.length > 0) view.dispatch(view.state.tr.setMeta(key, { ...s, held: s.extras }))
}

/** The objects held since the drag began, where they are now; the hold ends. */
export function releaseHeldObjects(view: EditorView): number[] {
  const s = key.getState(view.state)
  if (!s || s.held.length === 0) return []
  view.dispatch(
    view.state.tr
      .setMeta(key, s.extras.length > 0 ? { ...s, held: [] } : null)
      .setMeta('addToHistory', false),
  )
  return s.held
}

/** the positions of every member of the groups `positions` belong to, in document order */
export function groupMates(doc: PmNode, positions: number[]): number[] {
  const ids = new Set<string>()
  for (const pos of positions) {
    const id = doc.nodeAt(pos)?.attrs.groupId as string | null | undefined
    if (id) ids.add(id)
  }
  if (ids.size === 0) return []
  const out: number[] = []
  doc.descendants((node, pos) => {
    if (ids.has(node.attrs.groupId as string) && isFloatingObject(node)) out.push(pos)
  })
  return out
}

/**
 * Copies of group members are not members of the group they were copied
 * from: several that arrive together form a group of their own, one alone is
 * an object of its own. Clipboard HTML does not carry the group, so objects
 * pasted from it arrive ungrouped anyway; this is for content that reaches
 * the document as nodes. (Not a transformPasted prop: having one is how the
 * paste lanes tell the document editor from a text box's.)
 */
function regroupPasted(trs: readonly Transaction[], state: EditorState): Transaction | null {
  const pasted = new Map<string, number[]>()
  trs.forEach((paste, index) => {
    if (!paste.docChanged || paste.getMeta('uiEvent') !== 'paste') return
    const later = trs.slice(index + 1)
    paste.mapping.maps.forEach((map, at) => {
      map.forEach((_from, _to, start, end) => {
        // where the pasted range is now, after the steps and transactions that followed
        const rest = paste.mapping.slice(at + 1)
        const through = (pos: number, side: number) =>
          later.reduce((p, t) => t.mapping.map(p, side), rest.map(pos, side))
        const from = through(start, -1)
        const to = Math.min(through(end, 1), state.doc.content.size)
        if (to <= from) return
        state.doc.nodesBetween(from, to, (node, pos) => {
          const id = node.attrs.groupId as string | null | undefined
          if (!id || pos < from) return
          const list = pasted.get(id) ?? []
          if (!list.includes(pos)) list.push(pos)
          pasted.set(id, list)
        })
      })
    })
  })
  if (pasted.size === 0) return null
  const tr = state.tr
  for (const positions of pasted.values()) {
    const groupId = positions.length > 1 ? newGroupId() : null
    for (const pos of positions) {
      const node = state.doc.nodeAt(pos)!
      tr.setNodeMarkup(pos, undefined, {
        ...node.attrs,
        groupId,
        groupBox: groupId ? node.attrs.groupBox : null,
      })
    }
  }
  return tr
}

/** A carried drawing whose paragraph is gone becomes a paragraph itself, standing where it stood. */
function strandedCarried(state: EditorState): Transaction | null {
  const tr = state.tr
  let anchored = false
  state.doc.forEach((node, pos) => {
    const member = node.type.name === 'docProtected' && !!node.attrs.member
    if (node.attrs.carried && !anchored) {
      const down = Math.round(Number(node.attrs.carriedDy ?? 0) * 9525)
      tr.setNodeMarkup(pos, undefined, {
        ...node.attrs,
        carried: false,
        carriedDy: null,
        textboxes: (node.attrs.textboxes as TextboxDisplay[]).map((box) =>
          box.pageRelV || box.pagePinned
            ? box
            : { ...box, offsetYEmu: (box.offsetYEmu ?? 0) + down },
        ),
      })
    }
    anchored = member
  })
  if (!tr.docChanged) return null
  const sel = state.selection
  return sel instanceof NodeSelection ? tr.setSelection(NodeSelection.create(tr.doc, sel.from)) : tr
}

let carriedSettler: ((view: EditorView) => void) | null = null
export function setCarriedSettler(settler: ((view: EditorView) => void) | null): void {
  carriedSettler = settler
}

const carriedStops = new WeakMap<PmNode, Decoration[]>()

/**
 * A caret stop ahead of every carried drawing: the browser drops a selection
 * whose two ends both lie between non-editable blocks, so a drawing selected
 * between two others would lose it at the next key or pointer move.
 */
function carriedStopsOf(doc: PmNode): Decoration[] {
  let stops = carriedStops.get(doc)
  if (!stops) {
    stops = []
    doc.forEach((node, pos) => {
      if (!node.attrs.carried) return
      stops!.push(
        Decoration.widget(
          pos,
          () => {
            const stop = document.createElement('span')
            stop.className = 'doc-carried-stop'
            stop.appendChild(document.createElement('br'))
            return stop
          },
          // raw keeps it editable, which is what makes it a caret stop
          { side: -1, raw: true, key: 'carried-stop', ignoreSelection: true },
        ),
      )
    })
    carriedStops.set(doc, stops)
  }
  return stops
}

export const newGroupId = (): string =>
  `g${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`

/** The selected objects as clipboard content: each a block of its own, in document order. */
function objectsSlice(state: EditorState, positions: number[]): Slice {
  const paragraph = state.schema.nodes.docParagraph
  const blocks = [...positions]
    .sort((a, b) => a - b)
    .flatMap((pos) => {
      const node = state.doc.nodeAt(pos)
      if (!node) return []
      return [node.isInline ? paragraph.create(null, node) : node]
    })
  return new Slice(Fragment.fromArray(blocks), 0, 0)
}

/**
 * Copy and Cut take every selected object: the node selection alone would put
 * only the primary one on the clipboard and cut only that one out of a group.
 */
function copyObjects(view: EditorView, event: ClipboardEvent, cut: boolean): boolean {
  // text being edited inside one of the objects is what the user is copying, not the objects
  const nested = (event.target as Element | null)?.closest?.(NESTED_EDITORS)
  if (nested && nested !== view.dom) return false
  const positions = selectedObjectPositions(view.state)
  if (positions.length < 2 || !event.clipboardData) return false
  const slice = objectsSlice(view.state, positions)
  const { dom } = view.serializeForClipboard(slice)
  event.preventDefault()
  event.clipboardData.clearData()
  event.clipboardData.setData('text/html', dom.innerHTML)
  event.clipboardData.setData('text/plain', slice.content.textBetween(0, slice.content.size, '\n'))
  if (cut && view.editable) deleteObjects(view, positions)
  return true
}

function deleteObjects(view: EditorView, positions: number[]): void {
  const tr = view.state.tr
  for (const pos of [...positions].sort((a, b) => b - a)) {
    const node = tr.doc.nodeAt(pos)
    if (node) tr.delete(pos, pos + node.nodeSize)
  }
  view.dispatch(tr.setMeta(key, null).scrollIntoView())
}

/**
 * @param objectAt the floating object a press landed on (the hit test knows
 *   about behind-text objects the DOM target never reaches)
 */
export function objectMultiSelectPlugin(
  objectAt: (view: EditorView, event: MouseEvent) => number | null,
): Plugin<ObjectSelectState | null> {
  return new Plugin<ObjectSelectState | null>({
    key,
    state: {
      init: () => null,
      apply(tr, prev, _old, state) {
        const meta = tr.getMeta(key) as ObjectSelectState | null | undefined
        if (meta !== undefined) return meta
        if (!prev) return null
        const follow = (positions: number[]) =>
          positions
            .map((pos) => tr.mapping.map(pos))
            .filter((pos) => isFloatingObject(state.doc.nodeAt(pos)))
        const primary = tr.mapping.map(prev.primary)
        const held = follow(prev.held)
        // the selection moved on: only a hold outlives it, until the drag lets go
        if (primaryOf(state) !== primary)
          return held.length > 0 ? { primary: primaryOf(state) ?? -1, extras: [], held } : null
        const extras = follow(prev.extras).filter((pos) => pos !== primary)
        return extras.length > 0 || held.length > 0 ? { primary, extras, held } : null
      },
    },
    appendTransaction(trs, _old, state) {
      const pasted = regroupPasted(trs, state)
      if (pasted) return pasted
      const stranded = trs.some((tr) => tr.docChanged) ? strandedCarried(state) : null
      if (stranded) return stranded
      // a group selects whole: picking one member brings the others along
      const selected = selectedObjectPositions(state)
      if (selected.length === 0) return null
      const missing = groupMates(state.doc, selected).filter((pos) => !selected.includes(pos))
      if (missing.length === 0) return null
      const held = key.getState(state)?.held ?? []
      const tr = setObjectSelection(state.tr, [...selected, ...missing])
      return held.length > 0
        ? tr.setMeta(key, { ...(tr.getMeta(key) as ObjectSelectState), held })
        : tr
    },
    props: {
      decorations(state) {
        const stops = carriedStopsOf(state.doc)
        const s = key.getState(state)
        if (!s && stops.length === 0) return null
        const extras = (s?.extras ?? []).flatMap((pos) => {
          const node = state.doc.nodeAt(pos)
          return node
            ? [Decoration.node(pos, pos + node.nodeSize, { class: 'doc-multi-selected' })]
            : []
        })
        return DecorationSet.create(state.doc, [...stops, ...extras])
      },
      // Delete takes every selected object, not only the primary one
      handleKeyDown(view, event) {
        if (event.key !== 'Backspace' && event.key !== 'Delete') return false
        const positions = selectedObjectPositions(view.state)
        if (positions.length < 2 || !view.editable) return false
        deleteObjects(view, positions)
        return true
      },
    },
    view(view) {
      // capture phase: ahead of the drag handlers and of the modifier-click
      // that adds a text range
      const onDown = (event: MouseEvent) => {
        // the secondary button opens the menu for the objects selected: it never adds one
        const menu = event.button === 2 || (event.ctrlKey && isMac())
        if ((event.button !== 0 && !menu) || !view.editable) return
        // a grip of the selected object resizes or turns it, whatever keys are held
        if ((event.target as Element | null)?.closest?.(OBJECT_GRIPS)) return
        const selected = selectedObjectPositions(view.state)
        if (selected.length === 0) return
        const adding = !menu && (event.shiftKey || event.metaKey || event.ctrlKey)
        if (!adding && selected.length < 2) return
        const pos = objectAt(view, event)
        if (pos === null || !isFloatingObject(view.state.doc.nodeAt(pos))) return
        if (adding) {
          event.preventDefault()
          event.stopPropagation()
          toggleObjectSelection(view, pos)
        } else if (selected.includes(pos) && pos !== selected[0]) {
          // a press on one of the selected objects grabs them all: it becomes
          // the primary one and the press goes on to the drag handlers
          view.dispatch(
            setObjectSelection(view.state.tr, [pos, ...selected.filter((p) => p !== pos)]),
          )
        }
      }
      // ahead of the handlers that copy the primary object alone
      const onClipboard = (event: ClipboardEvent) => {
        if (copyObjects(view, event, event.type === 'cut')) event.stopImmediatePropagation()
      }
      view.dom.addEventListener('mousedown', onDown, true)
      view.dom.addEventListener('copy', onClipboard, true)
      view.dom.addEventListener('cut', onClipboard, true)
      return {
        update: (current, prev) => {
          if (current.state.doc === prev.doc || carriedStopsOf(current.state.doc).length === 0)
            return
          // a dispatch from inside the update would re-enter the view
          void Promise.resolve().then(() => {
            if (!current.isDestroyed) carriedSettler?.(current)
          })
        },
        destroy: () => {
          view.dom.removeEventListener('mousedown', onDown, true)
          view.dom.removeEventListener('copy', onClipboard, true)
          view.dom.removeEventListener('cut', onClipboard, true)
        },
      }
    },
  })
}

/** the editors of text inside an object: a text box's, a table cell's */
const NESTED_EDITORS = '.doc-textbox-editor, [contenteditable="true"]'

const OBJECT_GRIPS = '.img-resize-handle, .box-resize-handle, .box-rotate-handle'

const isMac = (): boolean => typeof navigator !== 'undefined' && /Mac/.test(navigator.platform)
