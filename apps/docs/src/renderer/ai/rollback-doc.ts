import type { Editor } from '@tiptap/core'
import type { PmNode } from '../editor/convert'
import { TABLE_TRAILING_SKIP } from '../editor/extensions'

/**
 * Swap the document for a version snapshot.
 *
 * Kept out of ProseMirror history on purpose: a roll back and its undo are one
 * action pair owned by the AI panel (#543), and letting `⌘Z` also act on the
 * same transaction is exactly how the document and the panel's version list
 * drift apart. The panel's own "undo roll back" is the way back.
 *
 * `TABLE_TRAILING_SKIP` is the same marker the file-open path sets, so the
 * trailing-table geometry pass does not fight the replacement.
 */
export function applyDocument(editor: Editor, doc: PmNode): void {
  editor
    .chain()
    .setMeta(TABLE_TRAILING_SKIP, true)
    .setMeta('addToHistory', false)
    .setContent(doc as never)
    .run()
}
