import { copyFile, rename, unlink } from 'node:fs/promises'
import { dirname, extname, join } from 'node:path'

/**
 * Where a rename to `name` puts the file: the same folder, the same extension
 * (typed or not). Null when the name is empty, names a folder, or carries
 * another extension.
 */
export function renamedDocPath(from: string, name: string): string | null {
  const typed = name.trim()
  if (!typed || typed.startsWith('.') || /[\\/:*?"<>|\0]/.test(typed)) return null
  const ext = extname(from)
  const hasExt = typed.toLowerCase().endsWith(ext.toLowerCase())
  const stem = hasExt ? typed.slice(0, typed.length - ext.length) : typed
  // a different extension would hand the file to another app
  if (!stem.trim() || (!hasExt && /\.(docx?|dotx?|rtf|odt|txt|pdf|md|html?)$/i.test(typed)))
    return null
  return join(dirname(from), stem + ext)
}

/** a move across volumes cannot be a rename: copy, then remove the original */
export async function moveFile(from: string, to: string): Promise<void> {
  try {
    await rename(from, to)
  } catch (err) {
    if ((err as NodeJS.ErrnoException).code !== 'EXDEV') throw err
    await copyFile(from, to)
    await unlink(from)
  }
}
