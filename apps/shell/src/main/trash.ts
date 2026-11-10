import { execFile, spawn } from 'node:child_process'
import { homedir } from 'node:os'
import { join } from 'node:path'
import { ipcMain, shell } from 'electron'
import { HOME_CHANNELS } from '../shared/home-api'

/**
 * Open the operating system's Recycle Bin / Trash window.
 * Resolves true once the system accepted the request, false when it could not be opened
 * (the shell shows a toast in that case).
 */
export async function openRecycleBin(): Promise<boolean> {
  switch (process.platform) {
    case 'win32':
      return openWindowsRecycleBin()
    case 'darwin':
      return openMacTrash()
    default:
      return openLinuxTrash()
  }
}

/**
 * Windows: the Recycle Bin is a shell folder, not a real path, so shell.openPath can't reach it.
 * `explorer.exe shell:RecycleBinFolder` opens it. explorer.exe exits with code 1 even on success,
 * so success is "the process started", never its exit code.
 */
function openWindowsRecycleBin(): Promise<boolean> {
  return new Promise((resolve) => {
    try {
      const child = spawn('explorer.exe', ['shell:RecycleBinFolder'], {
        detached: true,
        stdio: 'ignore',
      })
      child.once('error', () => resolve(false))
      child.once('spawn', () => {
        child.unref()
        resolve(true)
      })
    } catch {
      resolve(false)
    }
  })
}

function openMacTrash(): Promise<boolean> {
  return new Promise((resolve) => {
    // Finder owns the Trash; asking it avoids needing Full Disk Access to read ~/.Trash directly
    execFile(
      'osascript',
      ['-e', 'tell application "Finder" to open trash', '-e', 'tell application "Finder" to activate'],
      (error) => resolve(!error),
    )
  })
}

async function openLinuxTrash(): Promise<boolean> {
  const viaXdg = await new Promise<boolean>((resolve) => {
    execFile('xdg-open', ['trash:///'], (error) => resolve(!error))
  })
  if (viaXdg) return true
  // no trash:/// handler installed: open the freedesktop trash folder directly
  const dataHome = process.env.XDG_DATA_HOME || join(homedir(), '.local', 'share')
  const error = await shell.openPath(join(dataHome, 'Trash', 'files'))
  return error === ''
}

/** (Re)register the IPC handler behind the Trash / Recycle Bin button and its shortcut. */
export function registerTrashHandler(): void {
  ipcMain.removeHandler(HOME_CHANNELS.openTrash)
  ipcMain.handle(HOME_CHANNELS.openTrash, () => openRecycleBin())
}
