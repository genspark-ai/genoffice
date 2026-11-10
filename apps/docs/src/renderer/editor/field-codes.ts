/**
 * Word's Toggle Field Codes: a field shows `{ INSTR }` instead of its cached
 * result. The result text stays in the document; the code view is a
 * display-time decoration (hidden result + widget), so toggling never dirties
 * the file or the undo stack. Shift/⌥F9 semantics: `all` flips every field,
 * per-field toggles are exceptions to it.
 */
import { Extension } from '@tiptap/core'
import type { Editor } from '@tiptap/core'
import type { Mark, Node as PmNode } from '@tiptap/pm/model'
import { Plugin, PluginKey, type EditorState, type Transaction } from '@tiptap/pm/state'
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view'

export interface FieldRange {
  from: number
  to: number
  instr: string
  mark: Mark
}

interface Range {
  from: number
  to: number
}

interface FieldCodesState {
  all: boolean
  toggled: Range[]
  decos: DecorationSet
}

type FieldCodesMeta = { toggleAll: true } | { toggle: Range }

export const fieldCodesPluginKey = new PluginKey<FieldCodesState>('fieldCodes')

/** the field (contiguous run of one instrField mark instance) around `pos` */
export function fieldRangeAt(state: EditorState, pos: number): FieldRange | null {
  const markType = state.schema.marks.instrField
  if (!markType || pos < 0 || pos > state.doc.content.size) return null
  const $pos = state.doc.resolve(pos)
  const parent = $pos.parent
  if (!parent.isTextblock) return null
  const markOf = (node: PmNode | null | undefined) =>
    node?.isText ? node.marks.find((m) => m.type === markType) : undefined
  const after = parent.childAfter($pos.parentOffset)
  let mark = markOf(after.node)
  let index = after.index
  if (!mark && $pos.parentOffset > 0) {
    const before = parent.childBefore($pos.parentOffset)
    mark = markOf(before.node)
    index = before.index
  }
  if (!mark) return null
  return fieldAround(parent, $pos.start(), index, mark)
}

function fieldAround(parent: PmNode, base: number, index: number, mark: Mark): FieldRange {
  let start = index
  let end = index + 1
  while (start > 0 && mark.isInSet(parent.child(start - 1).marks)) start--
  while (end < parent.childCount && mark.isInSet(parent.child(end).marks)) end++
  let from = base
  for (let i = 0; i < start; i++) from += parent.child(i).nodeSize
  let to = from
  for (let i = start; i < end; i++) to += parent.child(i).nodeSize
  return { from, to, instr: String(mark.attrs.instr ?? ''), mark }
}

/** every field in the document, in order */
export function collectFields(
  doc: PmNode,
  markType = doc.type.schema.marks.instrField,
): FieldRange[] {
  const out: FieldRange[] = []
  if (!markType) return out
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return
    let i = 0
    let offset = pos + 1
    while (i < node.childCount) {
      const child = node.child(i)
      const mark = child.isText ? child.marks.find((m) => m.type === markType) : undefined
      if (!mark) {
        offset += child.nodeSize
        i++
        continue
      }
      const field = fieldAround(node, pos + 1, i, mark)
      out.push(field)
      while (offset < field.to) {
        offset += node.child(i).nodeSize
        i++
      }
    }
    return false
  })
  return out
}

/** Word's code view text: braces with one space padding around the instruction */
export function fieldCodeText(instr: string): string {
  return `{ ${instr.trim()} }`
}

const overlaps = (a: Range, b: Range) => a.from < b.to && b.from < a.to

export function isFieldCodeShown(state: EditorState, range: Range): boolean {
  const s = fieldCodesPluginKey.getState(state)
  if (!s) return false
  return s.all !== s.toggled.some((r) => overlaps(r, range))
}

export function toggleFieldCodes(view: EditorView, range: Range): void {
  const meta: FieldCodesMeta = { toggle: { from: range.from, to: range.to } }
  view.dispatch(view.state.tr.setMeta(fieldCodesPluginKey, meta))
}

export function toggleAllFieldCodes(view: EditorView): void {
  const meta: FieldCodesMeta = { toggleAll: true }
  view.dispatch(view.state.tr.setMeta(fieldCodesPluginKey, meta))
}

function buildDecorations(doc: PmNode, all: boolean, toggled: Range[]): DecorationSet {
  if (!all && toggled.length === 0) return DecorationSet.empty
  const decos: Decoration[] = []
  for (const field of collectFields(doc)) {
    if (all === toggled.some((r) => overlaps(r, field))) continue
    const code = fieldCodeText(field.instr)
    decos.push(Decoration.inline(field.from, field.to, { class: 'doc-field-result-hidden' }))
    decos.push(
      Decoration.widget(
        field.from,
        () => {
          const el = document.createElement('span')
          el.className = 'doc-field-code'
          el.textContent = code
          return el
        },
        { side: -1, key: `fc:${field.from}:${code}` },
      ),
    )
  }
  return decos.length ? DecorationSet.create(doc, decos) : DecorationSet.empty
}

function applyState(tr: Transaction, old: FieldCodesState): FieldCodesState {
  const meta = tr.getMeta(fieldCodesPluginKey) as FieldCodesMeta | undefined
  let { all, toggled } = old
  if (tr.docChanged && toggled.length) {
    toggled = toggled
      .map((r) => ({ from: tr.mapping.map(r.from, 1), to: tr.mapping.map(r.to, -1) }))
      .filter((r) => r.from < r.to)
  }
  if (meta && 'toggleAll' in meta) {
    all = !all
    toggled = []
  } else if (meta) {
    const hit = toggled.findIndex((r) => overlaps(r, meta.toggle))
    toggled = hit === -1 ? [...toggled, meta.toggle] : toggled.filter((_, i) => i !== hit)
  }
  // a change that deletes the last toggled field must also drop its decorations
  const active = all || toggled.length > 0 || old.toggled.length > 0
  if (!meta && !(active && tr.docChanged)) return old
  return { all, toggled, decos: buildDecorations(tr.doc, all, toggled) }
}

export const FieldCodesExtension = Extension.create({
  name: 'fieldCodes',
  addProseMirrorPlugins() {
    return [
      new Plugin<FieldCodesState>({
        key: fieldCodesPluginKey,
        state: {
          init: () => ({ all: false, toggled: [], decos: DecorationSet.empty }),
          apply: applyState,
        },
        props: {
          decorations(state) {
            return fieldCodesPluginKey.getState(state)?.decos ?? null
          },
        },
      }),
    ]
  },
})

/**
 * Edit Field… OK: rewrite the instruction and refresh the cached result in
 * one step (`value` empty = keep the old result, e.g. REF fields Word
 * recomputes on open). The run's other marks survive.
 */
export function setFieldInstr(
  editor: Editor,
  range: FieldRange,
  instr: string,
  value: string,
): boolean {
  const next = instr.trim()
  if (!next || !editor.isEditable) return false
  const { state } = editor
  const markType = state.schema.marks.instrField
  if (!markType) return false
  const mark = markType.create({ ...range.mark.attrs, instr: next, dirty: true })
  const text = value || state.doc.textBetween(range.from, range.to, '')
  if (!text) return false
  const runMarks = state.doc.nodeAt(range.from)?.marks ?? []
  const marks = [...markType.removeFromSet(runMarks), mark]
  const tr = state.tr.replaceWith(range.from, range.to, state.schema.text(text, marks))
  editor.view.dispatch(tr)
  return true
}
