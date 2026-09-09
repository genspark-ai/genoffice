import type { Editor } from '@tiptap/core'
import { runUiOps } from '../ai/ops'

/** Word's Ctrl+T / Ctrl+Shift+T: hanging indent one half-inch stop out or back (see the stepHangingIndent op) */
export function stepHangingIndent(editor: Editor, delta: 1 | -1): boolean {
  return runUiOps(editor, [{ op: 'stepHangingIndent', target: { scope: 'selection' }, delta }])
}

/** Ribbon increase/decrease indent: list items change level, paragraphs and headings snap to the next half-inch stop */
export function stepParagraphIndent(editor: Editor, delta: 1 | -1): boolean {
  return runUiOps(editor, [{ op: 'stepIndent', target: { scope: 'selection' }, delta }])
}
