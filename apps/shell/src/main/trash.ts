import { execFile, spawn } from 'node:child_process'
import { homedir } from 'node:os'
import { join } from 'node:path'
import { shell } from 'electron'

/**
 * Open the operating system's Recycle Bin / Trash window.
 *
 * Platform behaviors:
 * - Windows: The Recycle Bin is a virtual shell namespace, not a standard directory.
 *   `explorer.exe shell:RecycleBinFolder` opens it. Explorer often returns non-zero
 *   exit codes even on success, so success is defined by process spawn.
 * - macOS: Finder owns the Trash. Asking Finder via AppleScript avoids the Full Disk Access
 *   (TCC) restriction that blocks direct access to `~/.Trash`.
 * - Linux: Follows the Freedesktop spec via `xdg-open trash:///`, falling back to the
 *   raw trash files directory (`$XDG_DATA_HOME/Trash/files` or `~/.local/share/Trash/files`).
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

function openMacTrash(): Promise<boolean> {
  return new Promise((resolve) => {
    execFile(
      'osascript',
      [
        '-e',
        'tell application "Finder" to open trash',
        '-e',
        'tell application "Finder" to activate',
      ],
      (error) => resolve(!error),
    )
  })
}

async function openLinuxTrash(): Promise<boolean> {
  const viaXdg = await new Promise<boolean>((resolve) => {
    execFile('xdg-open', ['trash:///'], (error) => resolve(!error))
  })
  if (viaXdg) return true

  const dataHome = process.env.XDG_DATA_HOME || join(homedir(), '.local', 'share')
  const error = await shell.openPath(join(dataHome, 'Trash', 'files'))
  return error === ''
}
