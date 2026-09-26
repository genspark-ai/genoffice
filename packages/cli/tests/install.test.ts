import { chmodSync, lstatSync, mkdirSync, readlinkSync, symlinkSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { defaultCandidateDirs, inspectCliLink, installCliLink, isOurLauncher } from '../src/install'
import { tempDir } from './helpers'

describe('installCliLink', () => {
  it('links into the first writable directory and is idempotent', () => {
    const dir = tempDir()
    const bin = join(dir, 'bin')
    mkdirSync(bin)
    const launcher = join(dir, 'app', 'genoffice')
    mkdirSync(join(dir, 'app'))
    writeFileSync(launcher, '#!/bin/sh\n')
    expect(inspectCliLink({ launcher, platform: 'linux', candidateDirs: [bin] }).status).toBe(
      'missing',
    )
    const first = installCliLink({ launcher, platform: 'linux', candidateDirs: [bin] })
    expect(first).toEqual({ status: 'linked', location: join(bin, 'genoffice') })
    expect(readlinkSync(join(bin, 'genoffice'))).toBe(launcher)
    const again = installCliLink({ launcher, platform: 'linux', candidateDirs: [bin] })
    expect(again.status).toBe('present')
    expect(inspectCliLink({ launcher, platform: 'linux', candidateDirs: [bin] })).toEqual({
      status: 'present',
      location: join(bin, 'genoffice'),
    })
  })

  it('replaces our stale symlink but never a real file or a foreign symlink, and reports unwritable dirs with a manual command', () => {
    const dir = tempDir()
    const bin = join(dir, 'bin')
    mkdirSync(bin)
    const launcher = join(dir, 'new-app', 'cli', 'genoffice')
    mkdirSync(join(dir, 'new-app', 'cli'), { recursive: true })
    writeFileSync(launcher, '')
    symlinkSync(join(dir, 'old-app', 'cli', 'genoffice'), join(bin, 'genoffice'))
    expect(inspectCliLink({ launcher, platform: 'linux', candidateDirs: [bin] }).status).toBe(
      'missing',
    )
    expect(installCliLink({ launcher, platform: 'linux', candidateDirs: [bin] }).status).toBe(
      'linked',
    )
    expect(readlinkSync(join(bin, 'genoffice'))).toBe(launcher)

    const npm = join(dir, 'npm-bin')
    mkdirSync(npm)
    const npmTarget = join(dir, 'lib', 'node_modules', 'genoffice', 'bin', 'genoffice.js')
    mkdirSync(join(npmTarget, '..'), { recursive: true })
    writeFileSync(npmTarget, '')
    symlinkSync(npmTarget, join(npm, 'genoffice'))
    expect(installCliLink({ launcher, platform: 'linux', candidateDirs: [npm] }).status).toBe(
      'occupied',
    )
    expect(readlinkSync(join(npm, 'genoffice'))).toBe(npmTarget)
    expect(inspectCliLink({ launcher, platform: 'linux', candidateDirs: [npm] }).status).toBe(
      'occupied',
    )
    const spare = join(dir, 'spare-bin')
    mkdirSync(spare)
    expect(inspectCliLink({ launcher, platform: 'linux', candidateDirs: [npm, spare] })).toEqual({
      status: 'missing',
      location: join(spare, 'genoffice'),
      manual: expect.any(String),
    })
    expect(installCliLink({ launcher, platform: 'linux', candidateDirs: [npm, spare] })).toEqual({
      status: 'linked',
      location: join(spare, 'genoffice'),
    })

    const taken = join(dir, 'taken')
    mkdirSync(taken)
    writeFileSync(join(taken, 'genoffice'), 'someone else')
    const occupied = installCliLink({ launcher, platform: 'linux', candidateDirs: [taken] })
    expect(occupied.status).toBe('occupied')
    expect(occupied.manual).toContain('sudo')
    expect(occupied.manual).toContain('ln -sf')
    expect(lstatSync(join(taken, 'genoffice')).isSymbolicLink()).toBe(false)
    expect(inspectCliLink({ launcher, platform: 'linux', candidateDirs: [taken] }).status).toBe(
      'occupied',
    )

    const locked = join(dir, 'locked')
    mkdirSync(locked)
    chmodSync(locked, 0o555)
    const r = installCliLink({ launcher, platform: 'linux', candidateDirs: [locked] })
    const seen = inspectCliLink({ launcher, platform: 'linux', candidateDirs: [locked] })
    chmodSync(locked, 0o755)
    if (process.getuid?.() !== 0) {
      expect(r.status).toBe('unwritable')
      expect(r.manual).toContain(launcher)
      expect(seen.status).toBe('unwritable')
    }
  })

  it('owns only launchers shipped with the app, not any path ending in /cli/genoffice (genoffice#895)', () => {
    const dir = tempDir()
    const launcher = join(dir, 'GenOffice.app', 'Contents', 'Resources', 'cli', 'genoffice')
    mkdirSync(join(launcher, '..'), { recursive: true })
    writeFileSync(launcher, '#!/bin/sh\n')
    writeFileSync(join(launcher, '..', 'genoffice.cjs'), '')

    const vendor = join(dir, 'opt', 'vendor', 'cli', 'genoffice')
    mkdirSync(join(vendor, '..'), { recursive: true })
    writeFileSync(vendor, '#!/bin/sh\necho vendor\n')
    expect(isOurLauncher(vendor, launcher)).toBe(false)
    const bin = join(dir, 'bin')
    mkdirSync(bin)
    symlinkSync(vendor, join(bin, 'genoffice'))
    expect(installCliLink({ launcher, platform: 'linux', candidateDirs: [bin] })).toEqual({
      status: 'occupied',
      location: join(bin, 'genoffice'),
      manual: expect.any(String),
    })
    expect(readlinkSync(join(bin, 'genoffice'))).toBe(vendor)
    expect(inspectCliLink({ launcher, platform: 'linux', candidateDirs: [bin] }).status).toBe(
      'occupied',
    )

    const dead = join(dir, 'dead-bin')
    mkdirSync(dead)
    symlinkSync(join(dir, 'gone', 'cli', 'genoffice'), join(dead, 'genoffice'))
    expect(installCliLink({ launcher, platform: 'linux', candidateDirs: [dead] }).status).toBe(
      'linked',
    )
    expect(readlinkSync(join(dead, 'genoffice'))).toBe(launcher)

    const older = join(dir, 'opt', 'GenOffice', 'resources', 'cli', 'genoffice')
    mkdirSync(join(older, '..'), { recursive: true })
    writeFileSync(older, '#!/bin/sh\n')
    writeFileSync(join(older, '..', 'genoffice.cjs'), '')
    expect(isOurLauncher(older, launcher)).toBe(true)
    const upgraded = join(dir, 'upgraded-bin')
    mkdirSync(upgraded)
    symlinkSync(older, join(upgraded, 'genoffice'))
    expect(installCliLink({ launcher, platform: 'linux', candidateDirs: [upgraded] }).status).toBe(
      'linked',
    )
    expect(readlinkSync(join(upgraded, 'genoffice'))).toBe(launcher)

    const alias = join(dir, 'Applications')
    symlinkSync(join(dir, 'GenOffice.app'), alias, 'dir')
    const viaAlias = join(alias, 'Contents', 'Resources', 'cli', 'genoffice')
    expect(isOurLauncher(viaAlias, launcher)).toBe(true)
    const relative = join(dir, 'relative-bin')
    mkdirSync(relative)
    symlinkSync(join('..', 'opt', 'vendor', 'cli', 'genoffice'), join(relative, 'genoffice'))
    expect(installCliLink({ launcher, platform: 'linux', candidateDirs: [relative] }).status).toBe(
      'occupied',
    )
  })

  it('reports a missing /usr/local/bin as unwritable instead of skipping it', () => {
    const dir = tempDir()
    const launcher = join(dir, 'genoffice')
    writeFileSync(launcher, '')
    const absent = join(dir, 'no-such-bin')
    const r = installCliLink({ launcher, platform: 'linux', candidateDirs: [absent] })
    expect(r.status).toBe('unwritable')
    expect(r.location).toBe(join(absent, 'genoffice'))
    expect(r.manual).toContain('mkdir -p /usr/local/bin')
    expect(defaultCandidateDirs('darwin')[0]).toBe('/usr/local/bin')

    // inspect agrees with install: a later writable candidate counts
    const brew = join(dir, 'brew-bin')
    mkdirSync(brew)
    expect(inspectCliLink({ launcher, platform: 'darwin', candidateDirs: [absent, brew] })).toEqual(
      {
        status: 'missing',
        location: join(brew, 'genoffice'),
        manual: expect.stringContaining('ln -sf'),
      },
    )
    expect(inspectCliLink({ launcher, platform: 'darwin', candidateDirs: [absent] }).status).toBe(
      'unwritable',
    )
  })

  it('edits the user PATH on Windows without expanding existing entries', () => {
    const scripts: string[] = []
    const run = (script: string) => {
      scripts.push(script)
      return { ok: true, stdout: scripts.length === 1 ? 'linked\n' : 'present\n' }
    }
    const launcher =
      "C:\\Users\\O'Brien\\AppData\\Local\\Programs\\GenOffice\\resources\\genoffice\\genoffice.cmd"
    const dir = "C:\\Users\\O'Brien\\AppData\\Local\\Programs\\GenOffice\\resources\\genoffice"
    const first = installCliLink({ launcher, platform: 'win32', runPowerShell: run })
    expect(first).toEqual({ status: 'linked', location: dir })
    expect(scripts[0]).toContain("$dir = 'C:\\Users\\O''Brien\\AppData")
    expect(scripts[0]).toContain("'DoNotExpandEnvironmentNames'")
    expect(scripts[0]).toContain('RegistryValueKind]::ExpandString')
    expect(scripts[0]).toContain('SendMessageTimeout')
    expect(scripts[0]).not.toContain('SetEnvironmentVariable')
    expect(installCliLink({ launcher, platform: 'win32', runPowerShell: run }).status).toBe(
      'present',
    )
    const failed = installCliLink({
      launcher,
      platform: 'win32',
      runPowerShell: () => ({ ok: false, stdout: '' }),
    })
    expect(failed.status).toBe('unwritable')
    expect(failed.manual).toContain('SetEnvironmentVariable')
  })

  it('inspects the Windows PATH read-only', () => {
    const scripts: string[] = []
    const launcher = 'C:\\GenOffice\\resources\\genoffice\\genoffice.cmd'
    const present = inspectCliLink({
      launcher,
      platform: 'win32',
      runPowerShell: (s) => (scripts.push(s), { ok: true, stdout: 'present\n' }),
    })
    expect(present.status).toBe('present')
    expect(present.location).toBe('C:\\GenOffice\\resources\\genoffice')
    expect(scripts[0]).not.toContain('SetValue')
    const missing = inspectCliLink({
      launcher,
      platform: 'win32',
      runPowerShell: () => ({ ok: true, stdout: 'missing\n' }),
    })
    expect(missing.status).toBe('missing')
    expect(missing.manual).toContain('SetEnvironmentVariable')
  })
})
