/**
 * Round-trips google-settings.json under a real temp userData dir (electron's
 * app.getPath is mocked to point there, same pattern as google-auth-refresh.test.ts).
 */
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

let userDataDir: string

vi.mock('electron', () => ({
  app: { getPath: () => userDataDir },
}))

const { readGoogleSettings, writeGoogleSettings } = await import('../src/main/google-settings')

describe('google-settings read/write round trip', () => {
  beforeEach(() => {
    userDataDir = mkdtempSync(join(tmpdir(), 'gs-'))
  })

  afterEach(() => {
    rmSync(userDataDir, { recursive: true, force: true })
  })

  it('defaults to null/null when no settings file exists yet', async () => {
    const settings = await readGoogleSettings()
    expect(settings).toEqual({ defaultFolderId: null, defaultFolderName: null })
  })

  it('persists and reloads a default folder', async () => {
    await writeGoogleSettings({ defaultFolderId: 'folder-123', defaultFolderName: 'Clients' })
    const settings = await readGoogleSettings()
    expect(settings).toEqual({ defaultFolderId: 'folder-123', defaultFolderName: 'Clients' })
  })

  it('clearing writes back to null/null', async () => {
    await writeGoogleSettings({ defaultFolderId: 'folder-123', defaultFolderName: 'Clients' })
    await writeGoogleSettings({ defaultFolderId: null, defaultFolderName: null })
    const settings = await readGoogleSettings()
    expect(settings).toEqual({ defaultFolderId: null, defaultFolderName: null })
  })

  it('treats a malformed settings file as unset rather than throwing', async () => {
    await writeGoogleSettings({ defaultFolderId: 'x', defaultFolderName: 'y' })
    const { writeFileSync } = await import('node:fs')
    writeFileSync(join(userDataDir, 'google-settings.json'), '{not json')
    const settings = await readGoogleSettings()
    expect(settings).toEqual({ defaultFolderId: null, defaultFolderName: null })
  })
})
