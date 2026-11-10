/// The default folder where new/untitled files land on their first (silent)
/// save and where AI-generated drafts go. Historically hardcoded to
/// <Documents>/GenOffice; now user-configurable via the `defaultSaveDir` key
/// in userData/app-settings.json (set from the home screen's account menu).
/// Every editor main module resolves through here so they all honor the same
/// setting.
import { accessSync, constants, existsSync, mkdirSync, readFileSync, statSync } from 'node:fs'
import { dirname, isAbsolute, join, resolve } from 'node:path'

/** the subset of Electron's `app` needed here (kept structural: this package has no Electron dependency) */
export interface PathProvider {
  getPath(name: 'userData' | 'documents'): string
}

export const DEFAULT_SAVE_DIR_KEY = 'defaultSaveDir'

/**
 * Create `dir` and any missing parents one level at a time.
 *
 * `mkdirSync(dir, { recursive: true })` cannot be used on the startup path: when
 * the filesystem refuses the leaf with ENOENT while its parent exists, libuv's
 * recursive implementation keeps re-attempting the parent (EEXIST) and the leaf
 * (ENOENT) forever and pins the calling process at 100% CPU. Windows Defender's
 * Controlled Folder Access does exactly that for Documents, which froze the
 * whole app on launch. A plain per-level mkdir lets the refusal surface as a
 * thrown error that the callers below already handle.
 */
export function mkdirDirs(dir: string): void {
  const missing: string[] = []
  let probe = resolve(dir)
  for (;;) {
    if (existsSync(probe)) break
    missing.push(probe)
    const parent = dirname(probe)
    if (parent === probe) break
    probe = parent
  }
  for (const path of missing.reverse()) mkdirSync(path)
}

/** the configured folder from app-settings.json, or null when unset/unreadable */
export function readDefaultSaveDirSetting(settingsPath: string): string | null {
  try {
    const raw: unknown = JSON.parse(readFileSync(settingsPath, 'utf8'))
    if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return null
    const value = (raw as Record<string, unknown>)[DEFAULT_SAVE_DIR_KEY]
    return typeof value === 'string' && isAbsolute(value) ? value : null
  } catch {
    return null
  }
}

/** true when the directory exists (or could be created) and is writable */
export function isUsableSaveDir(dir: string): boolean {
  try {
    mkdirDirs(dir)
    if (!statSync(dir).isDirectory()) return false
    accessSync(dir, constants.W_OK)
    return true
  } catch {
    return false
  }
}

/**
 * Pick the effective default save folder: the configured one when it is
 * usable, otherwise the fallback (created on demand). A configured folder on
 * an unplugged drive or with revoked permissions degrades to the fallback
 * instead of breaking silent saves.
 */
export function resolveDefaultSaveDir(configured: string | null, fallbackDir: string): string {
  if (configured && isUsableSaveDir(configured)) return configured
  if (!isUsableSaveDir(fallbackDir)) throw new Error(`default save dir not usable: ${fallbackDir}`)
  return fallbackDir
}

/** convenience for the Electron mains: settings lookup + fallback in one call */
export function configuredDefaultSaveDir(app: PathProvider): string {
  const settingsPath = join(app.getPath('userData'), 'app-settings.json')
  const fallback = join(app.getPath('documents'), 'GenOffice')
  return resolveDefaultSaveDir(readDefaultSaveDirSetting(settingsPath), fallback)
}
