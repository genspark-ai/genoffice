import { Extension } from '@tiptap/core'
import type { Editor } from '@tiptap/core'
import { Plugin } from '@tiptap/pm/state'
import type { PmNode } from '../editor/convert'

/**
 * The storage slot where the mail-merge preview keeps the original template.
 *
 * Preview swaps the editor's document for merged values; the original (with its
 * merge-field codes) lives here so save/export and the AI readers can see the
 * real document while the preview is on. genoffice#2097.
 */
export interface MailMergePreviewStorage {
  mailmergePreviewJson: PmNode | null
  /**
   * The App-level preview lock. While it is on, the plugin below refuses every
   * document-changing transaction — the ribbon's greyed-out buttons are only
   * the visual layer, this is the real one.
   */
  locked: boolean
  /**
   * The editable state the document was in when the preview took over. The
   * preview ends by restoring it, so Read Mode / Protect Document is not
   * silently lifted (an unconditional restore used to hard-code true).
   */
  editableBeforePreview: boolean | null
}

declare module '@tiptap/core' {
  interface Storage {
    // TipTap keys an extension's storage by the extension name, so the slot
    // lives at editor.storage.mailmergePreview — not at the top level.
    mailmergePreview: MailMergePreviewStorage
  }
}

/**
 * Set on the transactions that are allowed to change the document even while
 * the preview lock is on: the preview's own swaps in and out of the merged
 * document, and the document-level replacements (open, phased stream, save
 * rebase). They are views of another document, not edits to the preview.
 */
export const MAIL_MERGE_PREVIEW_BYPASS = 'mailMergePreviewBypass'

export const MailMergePreviewExtension = Extension.create<
  Record<string, never>,
  MailMergePreviewStorage
>({
  name: 'mailmergePreview',

  addStorage() {
    return { mailmergePreviewJson: null, locked: false, editableBeforePreview: null }
  },

  addProseMirrorPlugins() {
    const getStorage = (): MailMergePreviewStorage | null =>
      this.editor?.storage?.mailmergePreview ?? null
    return [
      new Plugin({
        // The merged document is a view, not an edit: while the preview lock is
        // on, every transaction that changes the doc is refused wherever it
        // came from — typing, keyboard shortcuts, the context menu, AI tools,
        // programmatic dispatch. Selection-only and meta-only transactions go
        // through (walking the preview, scrolling to it, stays possible), and
        // so do the swaps the preview itself and the document load path make,
        // which carry MAIL_MERGE_PREVIEW_BYPASS.
        filterTransaction: (tr) => {
          const storage = getStorage()
          if (!storage?.locked) return true
          if (!tr.docChanged) return true
          return tr.getMeta(MAIL_MERGE_PREVIEW_BYPASS) === true
        },
      }),
    ]
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

/** turn the App-level preview lock on (locked) or off */
export function setPreviewLocked(editor: Editor, locked: boolean): void {
  const storage = editor.storage.mailmergePreview
  if (storage) storage.locked = locked
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
 * Ending the preview restores the editable state the document had before it
 * started rather than forcing true: Read Mode and Protect Document must
 * survive a preview that walked over them.
 *
 * The update event is suppressed: a preview is a view, not an edit, and the
 * document must not come out of it marked dirty. genoffice#2097.
 */
export function setPreviewReadOnly(editor: Editor, readOnly: boolean): void {
  const storage = editor.storage.mailmergePreview
  if (readOnly) {
    if (storage && storage.editableBeforePreview === null) {
      storage.editableBeforePreview = editor.isEditable
    }
    editor.setEditable(false, false)
    editor.view.setProps({ editable: () => false })
  } else {
    const editable = storage?.editableBeforePreview ?? true
    if (storage) storage.editableBeforePreview = null
    editor.setEditable(editable, false)
    editor.view.setProps({ editable: () => editable })
  }
}
