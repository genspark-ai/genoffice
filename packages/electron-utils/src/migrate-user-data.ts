/**
 * One-time migration for the product rename.
 *
 * Electron derives the packaged `userData` directory from `productName`, so
 * renaming the app changes the default userData path and would orphan every
 * user's existing settings (ai-settings.json, app-settings.json, recent files,
 * projects, chat history, workspace index, autosave/recovery data, PDF
 * signatures). On the first launch under the new name we copy the old
 * directory into the new one, mirroring the shell's earlier "AI Office →
 * GenOffice" migration (apps/shell/src/main/index.ts).
 *
 * Idempotent: once the new directory has any content it is left alone, so a
 * fresh install never resurrects stale data and a partial copy never
 * duplicates itself. Known Chromium cache dirs are skipped (they are
 * disposable and can be large).
 */
import { cpSync, existsSync, readdirSync } from 'node:fs'
import { join } from 'node:path'

/** Chromium runtime caches inside userData that are safe to skip on copy */
const SKIP_DIRS = new Set(['Cache', 'Code Cache', 'GPUCache', 'DawnCache', 'DawnGraphiteCache', 'DawnWebGPUCache', 'blob_storage'])

/**
 * Copy `oldName` → `newName` inside `appDataDir` when the new directory is
 * missing or empty. No-op when the old directory does not exist or the names
 * are equal.
 */
export function migrateUserDataDir(appDataDir: string, oldName: string, newName: string): void {
  if (!oldName || !newName || oldName === newName) return
  const oldDir = join(appDataDir, oldName)
  const newDir = join(appDataDir, newName)
  if (!existsSync(oldDir)) return
  const newEmpty = !existsSync(newDir) || readdirSync(newDir).length === 0
  if (!newEmpty) return
  cpSync(oldDir, newDir, {
    recursive: true,
    filter: (src) => {
      const top = (src ?? '').slice(oldDir.length + 1).split(/[\\/]/)[0] ?? ''
      return !SKIP_DIRS.has(top)
    },
  })
}
