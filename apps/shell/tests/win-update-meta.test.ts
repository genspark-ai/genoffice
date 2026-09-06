import { execFileSync } from 'node:child_process'
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'

const SCRIPT = join(__dirname, '../../../scripts/win-update-meta.cjs')
// the script refuses installers under 10 MB (a wrong-file guard)
const INSTALLER_BYTES = 10 * 1024 * 1024 + 1

let dir: string

function fakeInstaller(name: string, fill: number): string {
  const file = join(dir, name)
  writeFileSync(file, Buffer.alloc(INSTALLER_BYTES, fill))
  return file
}

function run(args: string[]): { ok: boolean; stderr: string } {
  try {
    execFileSync(process.execPath, [SCRIPT, ...args], { stdio: ['ignore', 'ignore', 'pipe'] })
    return { ok: true, stderr: '' }
  } catch (err) {
    return { ok: false, stderr: String((err as { stderr?: Buffer }).stderr ?? '') }
  }
}

/** Windows shares one feed across arches; electron-updater falls back to the
 * first listed file when none matches process.arch, so x64 must lead. */
describe('win-update-meta', () => {
  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), 'win-update-meta-'))
  })
  afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  it('writes an x64-only feed with the legacy archive name', () => {
    const x64 = fakeInstaller('GenOfficeSetup-v0.9.1.exe', 1)
    expect(run([x64, '0.9.1', 'GenOfficeSetup-v0.9.1.exe', dir, 'beta.yml']).ok).toBe(true)
    const feed = readFileSync(join(dir, 'beta.yml'), 'utf8')
    expect(feed).toMatch(/^version: 0\.9\.1\n/)
    expect(feed.match(/- url: /g)).toHaveLength(1)
    expect(feed).toContain('path: GenOfficeSetup-v0.9.1.exe\n')
    expect(readFileSync(join(dir, 'GenOffice-win-0.9.1.yml'), 'utf8')).toBe(feed)
  })

  it('lists x64 before arm64 and archives under the dual-arch name', () => {
    const x64 = fakeInstaller('GenOfficeSetup-v0.9.1.exe', 1)
    const arm = fakeInstaller('GenOfficeSetup-v0.9.1-arm64.exe', 2)
    const res = run([
      x64,
      '0.9.1',
      'GenOfficeSetup-v0.9.1.exe',
      dir,
      'beta.yml',
      arm,
      'GenOfficeSetup-v0.9.1-arm64.exe',
    ])
    expect(res.ok).toBe(true)
    const feed = readFileSync(join(dir, 'beta.yml'), 'utf8')
    const urls = [...feed.matchAll(/- url: (\S+)/g)].map((m) => m[1])
    expect(urls).toEqual(['GenOfficeSetup-v0.9.1.exe', 'GenOfficeSetup-v0.9.1-arm64.exe'])
    const hashes = [...feed.matchAll(/sha512: (\S+)/g)].map((m) => m[1])
    expect(hashes[0]).not.toBe(hashes[1])
    // legacy top-level fields describe the x64 installer
    expect(hashes[2]).toBe(hashes[0])
    expect(feed).toContain('path: GenOfficeSetup-v0.9.1.exe\n')
    expect(readFileSync(join(dir, 'GenOffice-win-x64-arm64-0.9.1.yml'), 'utf8')).toBe(feed)
  })

  it('rejects an arm64 name without the arch marker or an arm64-looking x64 name', () => {
    const x64 = fakeInstaller('GenOfficeSetup-v0.9.1.exe', 1)
    const arm = fakeInstaller('GenOfficeSetup-v0.9.1-arm64.exe', 2)
    expect(
      run([
        x64,
        '0.9.1',
        'GenOfficeSetup-v0.9.1.exe',
        dir,
        'beta.yml',
        arm,
        'GenOfficeSetup-v0.9.1-arm.exe',
      ]).stderr,
    ).toContain('must contain "arm64"')
    expect(
      run([
        arm,
        '0.9.1',
        'GenOfficeSetup-v0.9.1-arm64.exe',
        dir,
        'beta.yml',
        x64,
        'GenOfficeSetup-v0.9.1-arm64.exe',
      ]).stderr,
    ).toContain('must not contain "arm64"')
    expect(run([x64, '0.9.1', 'GenOfficeSetup-v0.9.1.exe', dir, 'beta.yml', arm]).stderr).toContain(
      'Usage',
    )
  })
})
