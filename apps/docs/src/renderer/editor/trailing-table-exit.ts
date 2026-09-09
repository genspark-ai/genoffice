import { Extension } from '@tiptap/core'
import { GapCursor } from '@tiptap/pm/gapcursor'
import { Plugin, PluginKey, TextSelection } from '@tiptap/pm/state'

const key = new PluginKey('trailingTableExit')

export const TrailingTableExitExtension = Extension.create({
  name: 'trailingTableExit',

  addProseMirrorPlugins() {
    const editor = this.editor
    return [
      new Plugin({
        key,
        appendTransaction(transactions, oldState, newState) {
          if (
            !editor.view.editable ||
            oldState.doc !== newState.doc ||
            !(newState.selection instanceof GapCursor) ||
            newState.selection.from !== newState.doc.content.size ||
            newState.doc.lastChild?.type.name !== 'docTable' ||
            !transactions.some((transaction) => transaction.selectionSet && !transaction.docChanged)
          ) {
            return null
          }

          const paragraph = newState.schema.nodes.docParagraph?.createAndFill()
          if (!paragraph) return null
          const pos = newState.doc.content.size
          const transaction = newState.tr.insert(pos, paragraph)
          return transaction
            .setSelection(TextSelection.create(transaction.doc, pos + 1))
            .scrollIntoView()
        },
      }),
    ]
  },
})
