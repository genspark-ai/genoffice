/**
 * Editing commands Word for Mac has keys for and the menus do not: field
 * locking and unlinking, moving between fields, outline levels, the ends of a
 * table row or column, the Spike and Go Back. Each edit is an undo step of
 * its own, however fast the key is pressed again.
 */
import { Extension, type Editor } from '@tiptap/core'
import { closeHistory } from '@tiptap/pm/history'
import { Fragment, Slice, type Node as PmNode } from '@tiptap/pm/model'
import { NodeSelection, Plugin, PluginKey, TextSelection, type EditorState } from '@tiptap/pm/state'
import { CellSelection, TableMap, isInTable, selectionCell } from '@tiptap/pm/tables'
import { headingStyleId, type StyleMap } from '../style-gallery'
import { collectFields, type FieldRange } from './field-codes'
import { selectionWithHiddenBody } from './heading-collapse'
import { collectHeadings } from './headings'
import { hfFieldStops, setHfPageFieldsLocked, unlinkHfPageFields } from './hf-field-lock'
import { MAX_HEADING_LEVEL } from './nav-outline'

const FIELD_MARKS = ['instrField', 'refField'] as const

/** the fields that keep a result in the text, in document order */
export function resultFields(doc: PmNode): FieldRange[] {
  const { marks } = doc.type.schema
  return FIELD_MARKS.flatMap((name) => (marks[name] ? collectFields(doc, marks[name]) : [])).sort(
    (a, b) => a.from - b.from,
  )
}

// a form field or one part of a field that runs over several paragraphs is written back as it came
const whole = (field: FieldRange): boolean =>
  !field.mark.attrs.beginXml && !field.mark.attrs.fieldPart

/** the fields the selection touches, a caret at either edge of one included */
function touchedFields(state: EditorState): FieldRange[] {
  const { from, to } = state.selection
  return resultFields(state.doc).filter((f) => whole(f) && f.from <= to && f.to >= from)
}

/** Word's Lock Fields / Unlock Fields on the selection; returns how many changed */
export function setFieldsLocked(editor: Editor, locked: boolean): number {
  // a header's live page number is a field without a result of its own: locking gives it one
  const live = setHfPageFieldsLocked(editor, locked)
  const { state, view } = editor
  const fields = touchedFields(state).filter((f) => (f.mark.attrs.locked === true) !== locked)
  if (fields.length === 0 || !editor.isEditable) return live
  const tr = closeHistory(state.tr)
  for (const f of fields) tr.addMark(f.from, f.to, f.mark.type.create({ ...f.mark.attrs, locked }))
  view.dispatch(tr)
  return live + fields.length
}

/** Word's Unlink Fields: each field the selection touches becomes its result as plain text */
export function unlinkFields(editor: Editor): number {
  const live = unlinkHfPageFields(editor)
  const { state, view } = editor
  const fields = touchedFields(state)
  if (fields.length === 0 || !editor.isEditable) return live
  const tr = closeHistory(state.tr)
  for (const f of fields) tr.removeMark(f.from, f.to, f.mark.type)
  view.dispatch(tr)
  return live + fields.length
}

/**
 * Word's Next Field / Previous Field: selects it; false when there is none that
 * way. `blocks` are where the fields kept as whole blocks start (a table of
 * contents, an index): such a field is selected by its first block.
 */
export function goToField(editor: Editor, dir: 1 | -1, blocks: readonly number[] = []): boolean {
  const { state, view } = editor
  const { from } = state.selection
  const stops = [
    ...resultFields(state.doc).map((f) => ({ from: f.from, to: f.to, node: false })),
    ...[...blocks, ...hfFieldStops(state.doc)].map((pos) => ({ from: pos, to: pos, node: true })),
  ].sort((a, b) => a.from - b.from)
  const next =
    dir > 0 ? stops.find((f) => f.from > from) : stops.reverse().find((f) => f.from < from)
  if (!next) return false
  const selection = next.node
    ? NodeSelection.create(state.doc, next.from)
    : TextSelection.create(state.doc, next.from, next.to)
  view.dispatch(state.tr.setSelection(selection).scrollIntoView())
  return true
}

const OUTLINE_BLOCKS: Record<string, true> = {
  docParagraph: true,
  docHeading: true,
  docListItem: true,
}

/**
 * Word's Promote (-1) / Demote (+1) on the paragraphs of the selection. A heading moves one
 * level; body text becomes a heading of the level of the heading before it when promoted, one
 * below that when demoted.
 */
export function shiftOutlineLevel(editor: Editor, dir: 1 | -1, styles?: StyleMap): boolean {
  const { state, view } = editor
  if (!editor.isEditable) return false
  const heading = state.schema.nodes.docHeading
  const { from, to } = state.selection
  const levelOf = (node: PmNode): number | null => {
    if (node.type === heading) return Number(node.attrs.level) || 1
    const style = styles?.get(String(node.attrs.styleId ?? ''))
    return style?.headingLevel && !style.headingOutlineOff ? style.headingLevel : null
  }
  const $from = state.doc.resolve(from)
  const top = $from.depth ? $from.before(1) : from
  // the heading a body paragraph answers to: the last one before it, inside the selection too
  let prior =
    collectHeadings(state.doc, styles, true)
      .filter((h) => h.pos < top)
      .pop()?.level ?? 1
  const tr = closeHistory(state.tr)
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isTextblock) return true
    if (!OUTLINE_BLOCKS[node.type.name]) return false
    const own = levelOf(node)
    const level = Math.min(
      Math.max(own === null ? prior + (dir > 0 ? 1 : 0) : own + dir, 1),
      MAX_HEADING_LEVEL,
    )
    if (own !== null) prior = level
    if (level === own) return false
    const styleId = headingStyleId(styles, level) ?? null
    if (node.type === heading) {
      // a style that says nothing of the level stays when the document has none for the new one
      const kept = styles?.get(String(node.attrs.styleId ?? ''))?.headingLevel
        ? styleId
        : (styleId ?? node.attrs.styleId)
      // a paragraph that only carries an outline level keeps its style
      const attrs = node.attrs.outlineOnly ? { level } : { level, styleId: kept }
      tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...attrs })
    } else if (own !== null) {
      // a heading numbered through its style stays the list item it was read as
      if (styleId) tr.setNodeMarkup(pos, undefined, { ...node.attrs, styleId })
    } else {
      const $pos = tr.doc.resolve(pos)
      if ($pos.parent.canReplaceWith($pos.index(), $pos.index() + 1, heading))
        tr.setNodeMarkup(pos, heading, { ...node.attrs, styleId, level })
    }
    return false
  })
  if (!tr.docChanged) return false
  view.dispatch(tr.scrollIntoView())
  return true
}

export const CELL_TARGETS = ['row-start', 'row-end', 'column-start', 'column-end'] as const
export type CellTarget = (typeof CELL_TARGETS)[number]

/** `row-end` or `row-end-extend`, as a key binding names the move */
export function parseCellMove(payload: string): { to: CellTarget; extend: boolean } | null {
  const extend = payload.endsWith('-extend')
  const to = CELL_TARGETS.find(
    (t) => t === (extend ? payload.slice(0, -'-extend'.length) : payload),
  )
  return to ? { to, extend } : null
}

/** the caret into the first or last cell of the row or column, or the cells up to it selected */
export function goToCell(editor: Editor, to: CellTarget, extend: boolean): boolean {
  const { state, view } = editor
  if (!isInTable(state)) return false
  const $cell = selectionCell(state)
  const table = $cell.node(-1)
  const tableStart = $cell.start(-1)
  const map = TableMap.get(table)
  const here = map.findCell($cell.pos - tableStart)
  const row = to === 'column-start' ? 0 : to === 'column-end' ? map.height - 1 : here.top
  const col = to === 'row-start' ? 0 : to === 'row-end' ? map.width - 1 : here.left
  const target = tableStart + map.map[row * map.width + col]
  const { selection } = state
  const next = extend
    ? CellSelection.create(
        state.doc,
        selection instanceof CellSelection ? selection.$anchorCell.pos : $cell.pos,
        target,
      )
    : TextSelection.near(state.doc.resolve(target + 1))
  view.dispatch(state.tr.setSelection(next).scrollIntoView())
  return true
}

interface SpikeItem {
  slice: Slice
  text: string
}

let spike: SpikeItem[] = []

export const spikeSize = (): number => spike.length

/** Word's Spike: the selection is cut and kept behind what was cut before */
export function cutToSpike(editor: Editor): boolean {
  const { state, view } = editor
  const { from, to, empty } = state.selection
  if (empty || !editor.isEditable) return false
  // as a cut does: a collapsed heading goes with the body it hides
  const folded = selectionWithHiddenBody(state)
  spike.push({
    slice: folded?.slice ?? state.selection.content(),
    text: state.doc.textBetween(from, folded?.to ?? to, '\n'),
  })
  view.dispatch(closeHistory(state.tr).deleteSelection().scrollIntoView())
  return true
}

/** Word's Insert Spike: everything cut to it goes in at the caret, each cut a paragraph, and it is emptied */
export function insertSpike(editor: Editor): boolean {
  if (spike.length === 0 || !editor.isEditable) return false
  const { state, view } = editor
  const items = spike
  const tr = closeHistory(state.tr)
  items.forEach((item, i) => {
    try {
      if (i > 0) tr.split(tr.selection.from)
      // a cut from another story may not fit this one: its text does
      const content = Fragment.fromJSON(state.schema, item.slice.content.toJSON())
      tr.replaceSelection(new Slice(content, item.slice.openStart, item.slice.openEnd))
    } catch {
      tr.insertText(item.text)
    }
  })
  spike = []
  view.dispatch(tr.scrollIntoView())
  return true
}

interface EditPlaces {
  /** where the last edits were made, the latest first */
  places: number[]
  /** how far back Go Back has gone, and where it set out from */
  step: number
  start: number | null
}

const MAX_PLACES = 3

export const goBackKey = new PluginKey<EditPlaces>('docsGoBack')

/** remembers where the document was last edited */
export const GoBackExtension = Extension.create({
  name: 'docsGoBack',
  addProseMirrorPlugins() {
    return [
      new Plugin<EditPlaces>({
        key: goBackKey,
        state: {
          init: () => ({ places: [], step: 0, start: null }),
          apply(tr, value) {
            const meta = tr.getMeta(goBackKey) as Partial<EditPlaces> | undefined
            if (meta) return { ...value, ...meta }
            // a caret placed by hand starts the round again
            if (!tr.docChanged)
              return tr.selectionSet && value.start !== null
                ? { ...value, step: 0, start: null }
                : value
            const places = value.places.map((pos) => tr.mapping.map(pos))
            // what the editor writes by itself (a settled field, a layout note) is not an edit
            if (tr.getMeta('addToHistory') === false) return { ...value, places }
            const here = tr.selection.head
            const $here = tr.doc.resolve(here)
            // typing on in the same paragraph is one place
            const rest = places.filter((pos) => !tr.doc.resolve(pos).sameParent($here))
            return { places: [here, ...rest].slice(0, MAX_PLACES), step: 0, start: null }
          },
        },
      }),
    ]
  },
})

/** a document just loaded has not been edited anywhere */
export function forgetEditPlaces(editor: Editor): void {
  const fresh: EditPlaces = { places: [], step: 0, start: null }
  editor.view.dispatch(editor.state.tr.setMeta(goBackKey, fresh).setMeta('addToHistory', false))
}

/**
 * Word's Go Back: the caret goes to where the document was last edited, then to the edits
 * before that one, and after the last of them back to where it set out.
 */
export function goBack(editor: Editor): boolean {
  const { state, view } = editor
  const value = goBackKey.getState(state)
  if (!value || value.places.length === 0) return false
  const head = state.selection.head
  const start = value.start ?? head
  const stops = [...value.places.filter((pos) => pos !== start), start]
  if (stops.length < 2 && stops[0] === head) return false
  const step = value.start === null ? 0 : value.step % stops.length
  const pos = Math.min(stops[step], state.doc.content.size)
  view.dispatch(
    state.tr
      .setSelection(TextSelection.near(state.doc.resolve(pos)))
      .setMeta(goBackKey, { step: step + 1, start })
      .scrollIntoView(),
  )
  return true
}
