import { mkdirSync, mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

import { controlEndpoint } from '../src/control'
import { genofficeUserDataDir, genofficeUserDataDirs, guiOpenDocuments } from '../src/gui'

/** Where the packaged profile sits for this platform, under a fake home. */
function appData(home: string): string {
  if (process.platform === 'darwin') return join(home, 'Library', 'Application Support')
  if (process.platform === 'win32') return join(home, 'AppData', 'Roaming')
  return join(home, '.config')
}

function fakeHome(): { home: string; env: NodeJS.ProcessEnv } {
  const home = mkdtempSync(join(tmpdir(), 'genoffice-devdir-'))
  return { home, env: { HOME: home } }
}

function publish(dir: string, name: string, body: unknown): void {
  mkdirSync(dir, { recursive: true })
  writeFileSync(join(dir, name), JSON.stringify(body))
}

describe('userData directory list', () => {
  it('scans the dev profile under both spellings it has had', () => {
    const { env, home } = fakeHome()
    const packaged = join(appData(home), 'GenOffice')
    // The shell writes "-Dev" since genoffice#1879; " Dev" is what it wrote
    // before, and a profile nobody migrated is still there.
    expect(genofficeUserDataDirs(env)).toEqual([packaged, `${packaged}-Dev`, `${packaged} Dev`])
  })

  it('prefers the current spelling and keeps the packaged profile first', () => {
    const { env } = fakeHome()
    const dirs = genofficeUserDataDirs(env)
    expect(dirs[0]).toBe(genofficeUserDataDir(env))
    expect(dirs.indexOf(`${dirs[0]}-Dev`)).toBeLessThan(dirs.indexOf(`${dirs[0]} Dev`))
  })

  it('scans nothing else when GENOFFICE_USER_DATA is set', () => {
    expect(genofficeUserDataDirs({ GENOFFICE_USER_DATA: '/u' })).toEqual(['/u'])
  })
})

describe('a dev shell is reachable', () => {
  it('finds the control endpoint published under "-Dev"', () => {
    // The regression genoffice#1879 left behind: the shell moved to "-Dev"
    // and this reader still looked for " Dev", so `open` stopped finding a
    // running dev instance and could only ever find the packaged one.
    const { env, home } = fakeHome()
    const dev = join(appData(home), 'GenOffice-Dev')
    publish(dev, 'control.json', {
      pid: process.pid,
      endpoint: 'http://127.0.0.1:1/dev',
      token: 'dev-token',
    })
    expect(controlEndpoint(env)?.token).toBe('dev-token')
  })

  it('still finds a profile left under the old " Dev" spelling', () => {
    const { env, home } = fakeHome()
    const old = join(appData(home), 'GenOffice Dev')
    publish(old, 'control.json', {
      pid: process.pid,
      endpoint: 'http://127.0.0.1:1/old',
      token: 'old-token',
    })
    expect(controlEndpoint(env)?.token).toBe('old-token')
  })

  it('reads a dev registry of open documents under "-Dev"', () => {
    // genoffice#980 made this check cover every live instance; it could only
    // ever cover the packaged one for as long as the name was wrong.
    const { env, home } = fakeHome()
    const dev = join(appData(home), 'GenOffice-Dev')
    publish(dev, 'open-documents.json', { pid: process.pid, paths: ['/dev.docx'] })
    expect(guiOpenDocuments(env)).toEqual([{ pid: process.pid, paths: ['/dev.docx'] }])
  })

  it('covers both instances at once, which is what the write guard needs', () => {
    const { env, home } = fakeHome()
    const packaged = join(appData(home), 'GenOffice')
    const dev = join(appData(home), 'GenOffice-Dev')
    publish(packaged, 'open-documents.json', { pid: process.pid, paths: ['/a.docx'] })
    publish(dev, 'open-documents.json', { pid: process.pid, paths: ['/b.docx'] })
    const paths = guiOpenDocuments(env).flatMap((instance) => instance.paths)
    expect(paths).toContain('/a.docx')
    expect(paths).toContain('/b.docx')
  })
})
