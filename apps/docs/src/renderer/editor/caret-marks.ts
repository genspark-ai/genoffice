/**
 * Word keeps an empty paragraph's pending format on its paragraph mark: type
 * in Calibri, press Enter, arrow up and back — typing still produces Calibri.
 * ProseMirror instead keeps the pending format in transient storedMarks,
 * which any selection move clears; the caret then re-derives marks from
 * adjacent text, and an empty paragraph has none — so navigation silently
 * reset new paragraphs to the document default (alpha ledger r114).
 *
 * This extension gives empty blocks a pilcrow memory via the shared
 * `caretMarks` attr: while the caret sits in an empty block *with* stored
 * marks, the marks are stamped onto the block; when the caret re-enters the
 * block bare (storedMarks === null, i.e. cleared by navigation — an explicit
 * empty array from a user toggle is respected), the stamp is restored as
 * storedMarks. Session-only: the attr is not persisted to the file.
 */
import { Extension } from '@tiptap/core'
import type { Mark } from '@tiptap/pm/model'
import { Plugin, PluginKey } from '@tiptap/pm/state'

/// Only pending FORMATTING survives; structural/annotation marks (comments,
/// links, revisions) must never re-materialize from a caret memory.
const FORMAT_MARKS = new Set(['bold', 'italic', 'underline', 'strike', 'docTextStyle'])

const serializeMarks = (marks: readonly Mark[]): string | null => {
  const kept = marks
    .filter((mark) => FORMAT_MARKS.has(mark.type.name))
    .map((mark) => ({ type: mark.type.name, attrs: mark.attrs }))
  return kept.length > 0 ? JSON.stringify(kept) : null
}

export const CaretMarksMemory = Extension.create({
  name: 'caretMarksMemory',
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: new PluginKey('caretMarksMemory'),
        appendTransaction: (_transactions, _oldState, newState) => {
          const { selection, storedMarks, schema } = newState
          if (!selection.empty) return null
          const $from = selection.$from
          const block = $from.parent
          if (!block.isTextblock || block.content.size > 0) return null
          if (!('caretMarks' in block.attrs)) return null
          const blockPos = $from.before($from.depth)
          if (storedMarks !== null) {
            // Caret holds pending marks (fresh Enter split, or a ribbon
            // format on the empty block): stamp them so they survive
            // navigation. Kept out of history — this is caret bookkeeping,
            // not a user edit to undo.
            const snapshot = serializeMarks(storedMarks)
            if (snapshot === block.attrs.caretMarks) return null
            return newState.tr
              .setNodeMarkup(blockPos, undefined, { ...block.attrs, caretMarks: snapshot })
              .setStoredMarks(storedMarks)
              .setMeta('addToHistory', false)
          }
          // Bare re-entry: navigation cleared the stored marks — restore the
          // block's pilcrow memory.
          const stamped = block.attrs.caretMarks as string | null
          if (!stamped) return null
          try {
            const parsed = JSON.parse(stamped) as { type: string; attrs: Record<string, unknown> }[]
            const marks = parsed.flatMap((entry) => {
              const type = schema.marks[entry.type]
              return type && FORMAT_MARKS.has(entry.type) ? [type.create(entry.attrs)] : []
            })
            if (marks.length === 0) return null
            return newState.tr.setStoredMarks(marks)
          } catch {
            return null
          }
        },
      }),
    ]
  },
})
