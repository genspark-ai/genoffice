import { createServer, type AddressInfo } from 'node:net'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { zoteroAvailable, zoteroInstallPaths } from '../src/main/zotero-detect'

describe('Zotero detection', () => {
  it('looks for the app and for the profile a past run left behind', () => {
    expect(zoteroInstallPaths('darwin', {}, '/Users/u')).toEqual([
      '/Applications/Zotero.app',
      join('/Users/u', 'Applications', 'Zotero.app'),
      join('/Users/u', 'Library', 'Application Support', 'Zotero'),
    ])
    expect(zoteroInstallPaths('linux', {}, '/home/u')).toContain(
      join('/home/u', '.zotero', 'zotero'),
    )
  })

  it('skips Windows roots the environment does not define', () => {
    const paths = zoteroInstallPaths(
      'win32',
      { ProgramFiles: 'C:\\PF', APPDATA: 'C:\\AD' },
      'C:\\u',
    )
    expect(paths).toEqual([
      join('C:\\PF', 'Zotero', 'zotero.exe'),
      join('C:\\AD', 'Zotero', 'Zotero'),
    ])
  })

  it('lets the environment pin the answer either way', async () => {
    expect(await zoteroAvailable({ env: { GENOFFICE_ZOTERO: '1' }, exists: () => false })).toBe(
      true,
    )
    expect(await zoteroAvailable({ env: { GENOFFICE_ZOTERO: '0' }, exists: () => true })).toBe(
      false,
    )
  })

  it('finds an install on disk without probing the port', async () => {
    expect(await zoteroAvailable({ env: {}, exists: () => true, port: 1 })).toBe(true)
  })

  it('falls back to a Zotero that is running from an unknown location', async () => {
    const server = createServer()
    await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve))
    const { port } = server.address() as AddressInfo
    try {
      expect(await zoteroAvailable({ env: {}, exists: () => false, port })).toBe(true)
    } finally {
      await new Promise((resolve) => server.close(resolve))
    }
    expect(await zoteroAvailable({ env: {}, exists: () => false, port })).toBe(false)
  })
})
