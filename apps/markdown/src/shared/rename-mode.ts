import { isSourceMode, textModeForPath } from './text-mode'

export type RenameAction = 'keep' | 'reload' | 'block-dirty'

/**
 * A rename across editing surfaces (note.md -> note.txt) must reload, or the
 * block editor would save the file back through markdown serialization.
 * Reloading discards unsaved edits, so a dirty document is refused instead.
 */
export function renameAction(
  currentPath: string | null,
  newPath: string,
  dirty: boolean,
): RenameAction {
  const from = textModeForPath(currentPath)
  const to = textModeForPath(newPath)
  if (from === to) return 'keep'
  if (currentPath === null) return 'reload'
  if (isSourceMode(to) !== isSourceMode(from)) return dirty ? 'block-dirty' : 'reload'
  return 'keep'
}
