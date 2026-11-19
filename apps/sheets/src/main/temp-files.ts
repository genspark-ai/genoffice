import { lstat, readdir, rm, unlink } from 'node:fs/promises'
import { basename, dirname, join, resolve } from 'node:path'

export const PASTED_FILE_TTL_MS = 7 * 24 * 60 * 60 * 1000

const UUID = '[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}'
const SESSION_SNAPSHOT_NAME = new RegExp(`^${UUID}\\.xlsx$`, 'i')
const IMPORT_DIRECTORY_NAME = new RegExp(`^${UUID}$`, 'i')
const OWNED_PASTED_FILE = /^pasted-[0-9]{8}-[0-9]{6}-[0-9]+\.(?:png|jpe?g|gif|webp)$/i

export interface SessionTemporaryResources {
  readonly snapshotPath: string
  readonly importTempDir?: string | undefined
}

interface CleanupSessionResourcesOptions extends SessionTemporaryResources {
  readonly tempRoot: string
  readonly closeSidecar: () => Promise<unknown>
}

function isDirectOwnedChild(path: string, parent: string, namePattern: RegExp): boolean {
  const resolvedPath = resolve(path)
  return dirname(resolvedPath) === resolve(parent) && namePattern.test(basename(resolvedPath))
}

/**
 * Close the sidecar first, remove its snapshot second, then remove a converted
 * CSV/XLS directory. Every filesystem deletion is constrained to an app-owned
 * temp root and filename shape.
 */
export async function cleanupSessionResources(
  options: CleanupSessionResourcesOptions,
): Promise<void> {
  try {
    await options.closeSidecar()
  } catch {
    // Continue with owned temp cleanup even when the sidecar is already gone.
  }
  const snapshotRoot = join(options.tempRoot, 'genoffice-sheets-sessions')
  if (isDirectOwnedChild(options.snapshotPath, snapshotRoot, SESSION_SNAPSHOT_NAME)) {
    await rm(options.snapshotPath, { force: true }).catch(() => undefined)
  }
  if (options.importTempDir !== undefined) {
    await cleanupImportTempDirectory(options.tempRoot, options.importTempDir)
  }
}

export async function cleanupImportTempDirectory(
  tempRoot: string,
  importTempDir: string,
): Promise<void> {
  const importRoot = join(tempRoot, 'genoffice-imports')
  if (!isDirectOwnedChild(importTempDir, importRoot, IMPORT_DIRECTORY_NAME)) return
  await rm(importTempDir, { recursive: true, force: true }).catch(() => undefined)
}

/**
 * Remove only old regular files created by savePastedImage. Arbitrary files,
 * directories, symlinks, and even similarly named nested paths are untouched.
 */
export async function cleanupExpiredPastedFiles(
  tempRoot: string,
  now = Date.now(),
): Promise<string[]> {
  const dir = join(tempRoot, 'genoffice-pasted')
  const cutoff = now - PASTED_FILE_TTL_MS
  const removed: string[] = []
  let entries
  try {
    entries = await readdir(dir, { withFileTypes: true })
  } catch {
    return removed
  }
  for (const entry of entries) {
    if (!entry.isFile() || !OWNED_PASTED_FILE.test(entry.name)) continue
    const path = join(dir, entry.name)
    try {
      const info = await lstat(path)
      if (!info.isFile() || info.mtimeMs >= cutoff) continue
      await unlink(path)
      removed.push(path)
    } catch {
      // Best-effort startup cleanup; another tab/process may win the race.
    }
  }
  return removed
}
