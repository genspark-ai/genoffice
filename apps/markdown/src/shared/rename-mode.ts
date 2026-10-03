import { isSourceMode, textModeForPath } from './text-mode'

/** What a rename has to do to an open document beyond moving its path. */
export type RenameAction =
  /** same editing surface: the path and the title move, the buffer stays */
  | 'keep'
  /** the extension crossed into another surface: reload from the new path */
  | 'reload'
  /** crossed surfaces with unsaved edits: refuse, the buffer cannot be moved */
  | 'block-dirty'

/**
 * A rename that changes the extension changes what the file *is*, so the open
 * document has to follow: `note.md` renamed to `note.txt` is no longer a block
 * document, and leaving the block editor in place would save the file back
 * through markdown serialization — rewriting the very text the rename was
 * meant to preserve.
 *
 * Reloading throws away unsaved edits, so a dirty document is refused instead.
 * The caller keeps the tab open and the file on disk is untouched; the user
 * saves (or discards) and the rename takes effect on the next attempt.
 */
export function renameAction(
  currentPath: string | null,
  newPath: string,
  dirty: boolean,
): RenameAction {
  const from = textModeForPath(currentPath)
  const to = textModeForPath(newPath)
  if (from === to) return 'keep'
  // an untitled document has nothing on disk to lose, so there is nothing a
  // reload could discard: the mode just follows the new extension
  if (currentPath === null) return 'reload'
  if (isSourceMode(to) !== isSourceMode(from)) return dirty ? 'block-dirty' : 'reload'
  // same surface, different grammar: json and plain both edit source, so only
  // the file path needs to move
  return 'keep'
}
