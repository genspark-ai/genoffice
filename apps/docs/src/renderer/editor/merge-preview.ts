/**
 * Word's Preview Results: the merge fields show one recipient's values in
 * place of their «names». Like the field-code view it is a display-time
 * decoration (hidden result + widget), so previewing never changes the
 * document, its unsaved state or the undo stack.
 */
import { Extension } from '@tiptap/core'
import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { Plugin, PluginKey, type EditorState } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import type { MergeRecord } from '@genoffice/docx-engine'
import {
  isRuleOnlyField,
  previewFieldTexts,
  previewPrompts,
  type MergePrompt,
} from '../mail-merge/template'
import { collectFields, fieldCodesPluginKey, isFieldCodeShown } from './field-codes'

/** the list the preview reads and the record it stands on */
interface Preview {
  records: readonly MergeRecord[]
  index: number
  /** what the user answered the Ask and Fill-in fields while previewing */
  answers?: ReadonlyMap<string, string>
}

const key = new PluginKey<Preview | null>('mergePreview')

/** the fields a merge reads, in document order */
function mergeFields(state: EditorState) {
  // a REF field shows the text a SET or an ASK before it gave its bookmark
  return [
    ...collectFields(state.doc),
    ...collectFields(state.doc, state.schema.marks.refField).map((ref) => ({
      ...ref,
      instr: `REF ${String(ref.mark.attrs.name ?? '')}`,
    })),
  ].sort((a, b) => a.from - b.from)
}

/** what previewing `records[index]` still has to ask */
export function mergePreviewPrompts(
  state: EditorState,
  records: readonly MergeRecord[],
  index: number,
  answers?: ReadonlyMap<string, string>,
): MergePrompt[] {
  return previewPrompts(
    mergeFields(state).map((f) => f.instr),
    records,
    index,
    answers,
  )
}

function buildDecorations(state: EditorState, preview: Preview): DecorationSet {
  const { doc } = state
  const fieldMark = state.schema.marks.instrField
  const refMark = state.schema.marks.refField
  const decos: Decoration[] = []
  const fields = mergeFields(state)
  const values = previewFieldTexts(
    fields.map((f) => f.instr),
    preview.records,
    preview.index,
    preview.answers,
  )
  fields.forEach((field, i) => {
    const value = values[i]
    // a field the merge leaves alone (a PAGE, a DATE) keeps showing what it shows,
    // one that only steers the merge shows nothing either way,
    // and one whose code is showing keeps showing its code
    if (value === null || isRuleOnlyField(field.instr) || isFieldCodeShown(state, field)) return
    // drawn inside the field's own formatting, as the merge writes it
    const marks = (doc.nodeAt(field.from)?.marks ?? []).filter(
      (m) => m.type !== fieldMark && m.type !== refMark,
    )
    decos.push(Decoration.inline(field.from, field.to, { class: 'doc-field-result-hidden' }))
    decos.push(
      Decoration.widget(
        field.from,
        () => {
          const el = document.createElement('span')
          el.className = 'doc-merge-preview'
          el.textContent = value
          return el
        },
        { side: -1, marks, key: `mp:${field.from}:${value}` },
      ),
    )
  })
  return decos.length ? DecorationSet.create(doc, decos) : DecorationSet.empty
}

export const MergePreviewExtension = Extension.create({
  name: 'mergePreview',
  addProseMirrorPlugins() {
    // rebuilt only when what it is drawn from changes: the document, the record,
    // or which fields show their codes
    let drawn: { doc: PmNode; preview: Preview; codes: unknown; decos: DecorationSet } | null = null
    return [
      new Plugin<Preview | null>({
        key,
        state: {
          init: () => null,
          apply(tr, preview) {
            const meta = tr.getMeta(key) as { preview: Preview | null } | undefined
            return meta ? meta.preview : preview
          },
        },
        props: {
          decorations(state) {
            const preview = key.getState(state)
            if (!preview) return null
            const codes = fieldCodesPluginKey.getState(state)
            if (drawn?.doc !== state.doc || drawn.preview !== preview || drawn.codes !== codes)
              drawn = { doc: state.doc, preview, codes, decos: buildDecorations(state, preview) }
            return drawn.decos
          },
        },
      }),
    ]
  },
})

/**
 * Shows `record` in the merge fields; null goes back to the field names. The
 * record is `records[index]`: the list is what NEXT reads on in and what the
 * record numbers count over.
 */
export function setMergePreview(
  editor: Editor,
  record: MergeRecord | null,
  records: readonly MergeRecord[] = record ? [record] : [],
  index = 0,
  answers?: ReadonlyMap<string, string>,
): void {
  const shown = key.getState(editor.state) ?? null
  if (
    !record
      ? !shown
      : shown?.records === records && shown.index === index && shown.answers === answers
  )
    return
  const preview = record ? { records, index, answers } : null
  editor.view.dispatch(editor.state.tr.setMeta(key, { preview }).setMeta('addToHistory', false))
}

/** the record the merge fields are showing, or null when they show their names */
export function mergePreviewRecord(state: EditorState): MergeRecord | null {
  const preview = key.getState(state)
  return preview?.records[preview.index] ?? null
}
