/// Electron 43 pins file dialogs without an explicit `defaultPath` to the
/// user's Downloads folder and no longer lets the OS restore the last-used
/// directory between invocations (electron/electron#49868). These wrappers
/// restore the pre-43 behavior the way the Electron breaking-changes guide
/// recommends: remember the directory of the last confirmed pick and thread
/// it into the next dialog's `defaultPath`.
import { basename, dirname, join } from 'node:path'

import type {
  BrowserWindow,
  Dialog,
  OpenDialogOptions,
  OpenDialogReturnValue,
  SaveDialogOptions,
  SaveDialogReturnValue,
} from 'electron'

// Keyed by the dialog module so tests with fake dialogs stay isolated. Each
// app bundles its own copy of this module, so the memory is scoped per
// editor — close enough to the per-app directory tracking the OS did before.
const lastUsedDirectoryByDialog = new WeakMap<Dialog, string>()

function withRememberedDirectory<T extends OpenDialogOptions | SaveDialogOptions>(
  dialog: Dialog,
  options: T,
  fallbackDir?: string,
): T {
  // No pick confirmed yet this session: anchor in the caller's fallback (the
  // configurable default save folder) instead of Electron's Downloads pin.
  const lastDir = lastUsedDirectoryByDialog.get(dialog) ?? fallbackDir
  if (!lastDir) return options
  if (options.defaultPath === undefined) return { ...options, defaultPath: lastDir }
  // A bare file name is a Save As suggestion: anchor it in the remembered
  // directory instead of letting Electron relocate it to Downloads.
  if (basename(options.defaultPath) === options.defaultPath) {
    return { ...options, defaultPath: join(lastDir, options.defaultPath) }
  }
  return options
}

export async function showOpenDialogWithMemory(
  dialog: Dialog,
  parent: BrowserWindow | null | undefined,
  options: OpenDialogOptions,
  fallbackDir?: string,
): Promise<OpenDialogReturnValue> {
  const withDir = withRememberedDirectory(dialog, options, fallbackDir)
  const result = parent
    ? await dialog.showOpenDialog(parent, withDir)
    : await dialog.showOpenDialog(withDir)
  const picked = result.filePaths[0]
  if (!result.canceled && picked) {
    lastUsedDirectoryByDialog.set(
      dialog,
      options.properties?.includes('openDirectory') ? picked : dirname(picked),
    )
  }
  return result
}

export async function showSaveDialogWithMemory(
  dialog: Dialog,
  parent: BrowserWindow | null | undefined,
  options: SaveDialogOptions,
  fallbackDir?: string,
): Promise<SaveDialogReturnValue> {
  const withDir = withRememberedDirectory(dialog, options, fallbackDir)
  const result = parent
    ? await dialog.showSaveDialog(parent, withDir)
    : await dialog.showSaveDialog(withDir)
  if (!result.canceled && result.filePath) {
    lastUsedDirectoryByDialog.set(dialog, dirname(result.filePath))
  }
  return result
}
