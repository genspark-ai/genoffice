import { Extension } from '@tiptap/core'
import type { Editor } from '@tiptap/core'
import type { PmNode } from '../editor/convert'

/**
 * The storage slot where the mail-merge preview keeps the original template.
 *
 * Preview swaps the editor's document for merged values; the original (with its
 * merge-field codes) lives here so save/export and the AI readers can see the
 * real document while the preview is on. This extension only declares the
 * storage — it adds no behaviour. genoffice#2097.
 */
export interface MailMergePreviewStorage {
  mailmergePreviewJson: PmNode | null
}

declare module '@tiptap/core' {
  interface Storage {
    // TipTap keys an extension's storage by the extension name, so the slot
    // lives at editor.storage.mailmergePreview — not at the top level.
    mailmergePreview: MailMergePreviewStorage
  }
}

export const MailMergePreviewExtension = Extension.create<
  Record<string, never>,
  MailMergePreviewStorage
>({
  name: 'mailmergePreview',

  addStorage() {
    return { mailmergePreviewJson: null }
  },
})

/**
 * The original template the preview is holding, or null when no preview is on.
 *
 * The extension's own storage is the only place the template lives, so every
 * reader goes through here: reading the top level of `editor.storage` would
 * always be undefined and silently hand back the preview document.
 */
export function previewTemplate(editor: Editor | null | undefined): PmNode | null {
  return editor?.storage.mailmergePreview?.mailmergePreviewJson ?? null
}

/** hand the original template to the preview (null hands the document back) */
export function setPreviewTemplate(editor: Editor, doc: PmNode | null): void {
  const storage = editor.storage.mailmergePreview
  if (storage) storage.mailmergePreviewJson = doc
}

/**
 * Make the editor read-only (or writable again) while the preview is on.
 *
 * `editor.setEditable` only flips the editor's own flag: TipTap never hands
 * `editable` to the ProseMirror view, so the view would keep accepting input
 * and the ribbon would keep dispatching on the preview document. Setting the
 * view prop is what actually stops the typing, and `editor.isEditable` is what
 * the ribbon reads to grey its editing commands out.
 *
 * The update event is suppressed: a preview is a view, not an edit, and the
 * document must not come out of it marked dirty. genoffice#2097.
 */
export function setPreviewReadOnly(editor: Editor, readOnly: boolean): void {
  editor.setEditable(!readOnly, false)
  editor.view.setProps({ editable: () => !readOnly })
}
