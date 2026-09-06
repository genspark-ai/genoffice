import { statSync } from 'node:fs'

/** Name characters Windows forbids (plus controls). Mirrors the PDF
    auto-renamer set (apps/pdf/src/main/pdf-main.ts) — the Home rename gate
    must reject them with a localized error instead of letting renameSync
    throw a raw OS error. */
// eslint-disable-next-line no-control-regex -- the C0 range IS the check: Windows forbids controls in names.
export const RENAME_ILLEGAL_NAME_CHARS = /[\\/:*?"<>|\u0000-\u001f]/

/** True when the trimmed name is usable as a file name. */
export function isValidRenameName(name: string): boolean {
  return name.length > 0 && !RENAME_ILLEGAL_NAME_CHARS.test(name)
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
