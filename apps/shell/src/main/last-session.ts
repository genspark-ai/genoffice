import { statSync } from 'node:fs'
import { isAbsolute } from 'node:path'

/**
 * Reopen the documents of the previous session at launch (genoffice#2162).
 * Off by default; when on, the shell window's file tabs are stored in
 * app-settings.json as it closes and reopened at the next launch.
 */
export const RESTORE_LAST_SESSION_KEY = 'restoreLastSession'
export const LAST_SESSION_DOCUMENTS_KEY = 'lastSessionDocuments'

/** a hand-edited settings file must not open hundreds of tabs at launch */
const MAX_RESTORED_DOCUMENTS = 50

export function restoreLastSessionEnabled(settings: Record<string, unknown>): boolean {
  return settings[RESTORE_LAST_SESSION_KEY] === true
}

/** the saved tab paths, once each, absolute only, in tab order */
export function readLastSessionDocuments(settings: Record<string, unknown>): string[] {
  const raw = settings[LAST_SESSION_DOCUMENTS_KEY]
  if (!Array.isArray(raw)) return []
  const paths = raw.filter((p): p is string => typeof p === 'string' && isAbsolute(p))
  return [...new Set(paths)].slice(0, MAX_RESTORED_DOCUMENTS)
}

function isFile(path: string): boolean {
  try {
    return statSync(path).isFile()
  } catch {
    return false
  }
}

/**
 * What to open at launch: the previous session's documents that still exist,
 * then the files the launch itself names, so a double-clicked file ends up the
 * active tab. A file already restored is not opened twice.
 */
export function launchPathsWithLastSession(
  settings: Record<string, unknown>,
  launchPaths: readonly string[],
  exists: (path: string) => boolean = isFile,
): string[] {
  if (!restoreLastSessionEnabled(settings)) return [...launchPaths]
  const restored = readLastSessionDocuments(settings).filter(
    (p) => !launchPaths.includes(p) && exists(p),
  )
  return [...restored, ...launchPaths]
}
