import { statSync } from 'node:fs'

/** Name characters Windows forbids (plus controls). Mirrors the PDF
    auto-renamer set (apps/pdf/src/main/pdf-main.ts) — the Home rename gate
    must reject them with a localized error instead of letting renameSync
    throw a raw OS error. */
// eslint-disable-next-line no-control-regex -- the C0 range IS the check: Windows forbids controls in names.
export const RENAME_ILLEGAL_NAME_CHARS = /[\\/:*?"<>|\u0000-\u001f]/

/** Windows device names reserved with or without an extension (CON.pdf is still CON). */
const RENAME_RESERVED_BASE = /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(\..*)?$/i

/** True when the raw requested name is usable: surrounding whitespace is
    rejected instead of silently trimmed (the Home rename gate must reject
    "report " with a localized error rather than renaming to "report"). */
export function isValidRawRenameName(newName: string): boolean {
  if (newName !== newName.trim()) return false
  return isValidRenameName(newName.trim())
}

/** True when the trimmed name is usable as a file name. */
export function isValidRenameName(name: string): boolean {
  if (name.length === 0 || name.length > 255) return false
  if (RENAME_ILLEGAL_NAME_CHARS.test(name)) return false
  // Windows strips trailing dots/spaces: renaming to "file." lands elsewhere,
  // so reject with the localized gate instead of a surprising rename.
  if (name.endsWith('.') || name.endsWith(' ')) return false
  if (RENAME_RESERVED_BASE.test(name)) return false
  return true
}

/** True when both paths resolve to the same on-disk file (same device +
    inode). Case-insensitive filesystems report the source itself for a
    case-only rename target, so callers use this instead of a bare exists
    check; on case-sensitive volumes two case variants are distinct files
    and must still trip the already-exists gate. */
export function isSameFile(a: string, b: string): boolean {
  try {
    const sa = statSync(a)
    const sb = statSync(b)
    return sa.dev === sb.dev && sa.ino === sb.ino
  } catch {
    return false
  }
}

/** Save As target for an open PDF, or null when the dialog was cancelled or
    the pick resolves to the open file itself (case/symlink variants included),
    which would otherwise copy a file onto itself. */
export function pdfSaveAsTarget(
  picked: { canceled: boolean; filePath?: string },
  currentPath: string,
): string | null {
  if (picked.canceled || !picked.filePath) return null
  if (picked.filePath === currentPath || isSameFile(picked.filePath, currentPath)) return null
  return picked.filePath
}
