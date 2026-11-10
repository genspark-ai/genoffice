import { readFileSync, statSync } from 'node:fs'
import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it, vi } from 'vitest'

vi.mock('electron', () => ({ app: { isPackaged: false, getVersion: () => '0.0.0' } }))

import { isEphemeralInstall, launcherFilePath, writeLauncherFile } from '../src/main/cli-link'

describe('genoffice launcher file', () => {
  it('lives in the genoffice auth directory, overridable like auth.json', () => {
    expect(launcherFilePath({})).toBe(join(process.env.HOME ?? '', '.genoffice', 'launcher'))
    expect(launcherFilePath({ GENOFFICE_AUTH_DIR: '/tmp/x' })).toBe(join('/tmp/x', 'launcher'))
  })

  it('writes one line, creates the directory, and only rewrites on change', () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-launcher-'))
    const file = join(dir, 'nested', 'launcher')
    expect(writeLauncherFile(file, '/Applications/GenOffice.app/Contents/Resources/cli')).toBe(true)
    expect(readFileSync(file, 'utf-8')).toBe('/Applications/GenOffice.app/Contents/Resources/cli\n')
    const before = statSync(file).mtimeMs
    expect(writeLauncherFile(file, '/Applications/GenOffice.app/Contents/Resources/cli')).toBe(
      false,
    )
    expect(statSync(file).mtimeMs).toBe(before)
    expect(writeLauncherFile(file, 'C:\\Programs\\GenOffice\\resources\\genoffice')).toBe(true)
    expect(readFileSync(file, 'utf-8')).toBe('C:\\Programs\\GenOffice\\resources\\genoffice\n')
  })

  it('treats dmg and AppImage mounts as temporary', () => {
    expect(isEphemeralInstall('/Volumes/GenOffice/GenOffice.app/Contents/Resources', {})).toBe(true)
    expect(isEphemeralInstall('/tmp/.mount_GenOfxyz/resources', {})).toBe(true)
    expect(isEphemeralInstall('/opt/GenOffice/resources', { APPIMAGE: '/home/u/G.AppImage' })).toBe(
      true,
    )
    expect(isEphemeralInstall('/Applications/GenOffice.app/Contents/Resources', {})).toBe(false)
  })
})
