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
import { mergeFieldValue } from '../mail-merge/template'
import { collectFields, fieldCodesPluginKey, isFieldCodeShown } from './field-codes'

const key = new PluginKey<MergeRecord | null>('mergePreview')

function buildDecorations(state: EditorState, record: MergeRecord): DecorationSet {
  const { doc } = state
  const fieldMark = state.schema.marks.instrField
  const decos: Decoration[] = []
  for (const field of collectFields(doc)) {
    const value = mergeFieldValue(field.instr, record)
    // a field the merge leaves alone (a PAGE, a DATE) keeps showing what it shows,
    // and one whose code is showing keeps showing its code
    if (value === null || isFieldCodeShown(state, field)) continue
    // drawn inside the field's own formatting, as the merge writes it
    const marks = (doc.nodeAt(field.from)?.marks ?? []).filter((m) => m.type !== fieldMark)
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
  }
  return decos.length ? DecorationSet.create(doc, decos) : DecorationSet.empty
}

export const MergePreviewExtension = Extension.create({
  name: 'mergePreview',
  addProseMirrorPlugins() {
    // rebuilt only when what it is drawn from changes: the document, the record,
    // or which fields show their codes
    let drawn: { doc: PmNode; record: MergeRecord; codes: unknown; decos: DecorationSet } | null =
      null
    return [
      new Plugin<MergeRecord | null>({
        key,
        state: {
          init: () => null,
          apply(tr, record) {
            const meta = tr.getMeta(key) as { record: MergeRecord | null } | undefined
            return meta ? meta.record : record
          },
        },
        props: {
          decorations(state) {
            const record = key.getState(state)
            if (!record) return null
            const codes = fieldCodesPluginKey.getState(state)
            if (drawn?.doc !== state.doc || drawn.record !== record || drawn.codes !== codes)
              drawn = { doc: state.doc, record, codes, decos: buildDecorations(state, record) }
            return drawn.decos
          },
        },
      }),
    ]
  },
})

/** Shows `record` in the merge fields; null goes back to the field names. */
export function setMergePreview(editor: Editor, record: MergeRecord | null): void {
  if ((key.getState(editor.state) ?? null) === record) return
  editor.view.dispatch(editor.state.tr.setMeta(key, { record }).setMeta('addToHistory', false))
}

/** the record the merge fields are showing, or null when they show their names */
export const mergePreviewRecord = (state: EditorState): MergeRecord | null =>
  key.getState(state) ?? null
