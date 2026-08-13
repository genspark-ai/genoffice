/**
 * Persisted Google Drive preferences for this app (currently just the default
 * destination folder for "Send to Google Docs"). Stored at
 * userData/google-settings.json — small, best-effort, never holds secrets.
 */
import { existsSync, readFileSync } from 'node:fs'
import { writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { app } from 'electron'

export interface GoogleSettings {
  defaultFolderId: string | null
  defaultFolderName: string | null
}

const DEFAULT_SETTINGS: GoogleSettings = { defaultFolderId: null, defaultFolderName: null }

function settingsPath(): string {
  return join(app.getPath('userData'), 'google-settings.json')
}

export async function readGoogleSettings(): Promise<GoogleSettings> {
  try {
    if (!existsSync(settingsPath())) return { ...DEFAULT_SETTINGS }
    const parsed = JSON.parse(readFileSync(settingsPath(), 'utf-8')) as Partial<GoogleSettings>
    return {
      defaultFolderId: typeof parsed.defaultFolderId === 'string' ? parsed.defaultFolderId : null,
      defaultFolderName:
        typeof parsed.defaultFolderName === 'string' ? parsed.defaultFolderName : null,
    }
  } catch {
    return { ...DEFAULT_SETTINGS }
  }
}

export async function writeGoogleSettings(settings: GoogleSettings): Promise<void> {
  await writeFile(settingsPath(), JSON.stringify(settings), 'utf-8')
}
