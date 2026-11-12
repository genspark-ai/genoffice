import { spawn } from 'node:child_process'
import { join } from 'node:path'
import { app, shell } from 'electron'

/**
 * Open the operating system's Recycle Bin / Trash window.
 *
 * Platform behaviors:
 * - Windows: The Recycle Bin is a virtual shell namespace, not a standard directory.
 *   `explorer.exe shell:RecycleBinFolder` opens it. Explorer often returns non-zero
 *   exit codes even on success, so success is defined by process spawn.
 * - macOS: Hand `~/.Trash` to Launch Services via `shell.openPath`. Finder
 *   shows the Trash window without triggering AppleScript automation entitlements.
 * - Linux: Follows the Freedesktop spec via `xdg-open trash:///`. Spawns
 *   detached with ignored stdio so the IPC promise settles immediately.
 *
 * Resolves true if the system accepted the request, false otherwise.
 */
export async function openTrash(platform: NodeJS.Platform = process.platform): Promise<boolean> {
  switch (platform) {
    case 'win32':
      return openWindowsRecycleBin()
    case 'darwin':
      return openMacTrash()
    default:
      return openLinuxTrash()
  }
}

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

async function openMacTrash(): Promise<boolean> {
  try {
    const err = await shell.openPath(join(app.getPath('home'), '.Trash'))
    return err === ''
  } catch {
    return false
  }
}

function openLinuxTrash(): Promise<boolean> {
  return new Promise((resolve) => {
    try {
      const child = spawn('xdg-open', ['trash:///'], {
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
