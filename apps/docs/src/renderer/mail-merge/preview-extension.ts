import { Extension } from '@tiptap/core'
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
    mailmergePreviewJson: PmNode | null
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
