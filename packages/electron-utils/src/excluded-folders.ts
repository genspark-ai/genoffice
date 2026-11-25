/// Folders the user excluded from GenOffice's memory surfaces (genoffice#2844): the
/// recents list, the search index and the session restore stop recording and
/// listing files under them. An explicit star keeps working — that is the
/// user's own intent, not the app remembering something they did not ask it
/// to remember.
///
/// The list lives in app-settings.json (`excludedFolders`), which every
/// app's main process reads; this module is the one shared definition of the
/// shape and the match rule, so the surfaces cannot drift apart.

import { isAbsolute, resolve, sep } from 'node:path'

export const EXCLUDED_FOLDERS_KEY = 'excludedFolders'

/** a hand-edited settings file must not carry hundreds of entries */
const MAX_EXCLUDED_FOLDERS = 200

/** the stored exclusion list: absolute paths, resolved, once each, capped */
export function readExcludedFolders(settings: Record<string, unknown>): string[] {
  const raw = settings[EXCLUDED_FOLDERS_KEY]
  return normalizeExcludedFolders(Array.isArray(raw) ? raw : [])
}

/** what the set IPC accepts: absolute strings only, resolved, once each, capped */
export function normalizeExcludedFolders(paths: readonly unknown[]): string[] {
  if (!Array.isArray(paths)) return []
  const out: string[] = []
  const seen = new Set<string>()
  for (const raw of paths) {
    if (typeof raw !== 'string' || !isAbsolute(raw)) continue
    const resolved = resolve(raw)
    const key = foldKey(resolved)
    if (seen.has(key)) continue
    seen.add(key)
    out.push(resolved)
    if (out.length >= MAX_EXCLUDED_FOLDERS) break
  }
  return out
}

/** true when path IS an excluded folder or sits inside one. A folder the
 *  user excluded must not shadow a sibling: the boundary is the separator,
 *  never a bare prefix ("/Users/x/notes-old" is not inside "/Users/x/notes").
 *  macOS and Windows compare case-insensitively, like their filesystems. */
export function isUnderExcludedFolder(path: string, folders: readonly string[]): boolean {
  if (folders.length === 0) return false
  const fold = CASE_FOLD
  const candidate = resolve(path)
  const needle = fold ? candidate.toLowerCase() : candidate
  for (const raw of folders) {
    const folder = resolve(raw)
    const base = fold ? folder.toLowerCase() : folder
    if (needle === base) return true
    const bounded = folder.endsWith(sep) ? folder : folder + sep
    const boundedFolded = fold ? bounded.toLowerCase() : bounded
    if (needle.startsWith(boundedFolded)) return true
  }
  return false
}

/** the string both sides of the comparison fold to (platform case rules) */
function foldKey(resolved: string): string {
  return CASE_FOLD ? resolved.toLowerCase() : resolved
}

const CASE_FOLD = process.platform === 'darwin' || process.platform === 'win32'
