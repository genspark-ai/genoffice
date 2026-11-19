import { spawnSync } from 'node:child_process'
import {
  chmodSync,
  existsSync,
  lstatSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readlinkSync,
  rmSync,
  symlinkSync,
  writeFileSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

const BUILD = join(__dirname, '..', 'build')

/** The scripts hardcode the package paths; a copy with the two roots swapped runs against a temp tree. */
function makeRoot() {
  const root = mkdtempSync(join(tmpdir(), 'genoffice-postinst-'))
  const opt = join(root, 'opt', 'GenOffice')
  const bin = join(root, 'usr', 'bin')
  mkdirSync(join(opt, 'resources', 'cli'), { recursive: true })
  mkdirSync(bin, { recursive: true })
  const launcher = join(opt, 'resources', 'cli', 'genoffice')
  writeFileSync(launcher, '#!/bin/sh\n')
  chmodSync(launcher, 0o755)
  const script = (name: string) => {
    const file = join(root, name)
    const body = readFileSync(join(BUILD, name), 'utf-8')
      .replaceAll('/opt/GenOffice', opt)
      .replaceAll('/usr/bin/', `${bin}/`)
    writeFileSync(file, body)
    return (arg = '') => {
      const r = spawnSync('sh', [file, arg], { encoding: 'utf-8' })
      return { code: r.status, stderr: r.stderr }
    }
  }
  return {
    launcher,
    link: join(bin, 'genoffice'),
    install: script('linux-after-install.sh'),
    remove: script('linux-after-remove.sh'),
  }
}

describe.skipIf(process.platform === 'win32')('linux cli link scripts (genoffice#893)', () => {
  it('links a free name, a dead link and a link into the old install', () => {
    const t = makeRoot()
    expect(t.install()).toEqual({ code: 0, stderr: '' })
    expect(readlinkSync(t.link)).toBe(t.launcher)

    expect(t.remove('remove').code).toBe(0)
    expect(existsSync(t.link)).toBe(false)
    symlinkSync('/nowhere/cli/genoffice', t.link)
    expect(t.install()).toEqual({ code: 0, stderr: '' })
    expect(readlinkSync(t.link)).toBe(t.launcher)

    rmSync(t.link)
    symlinkSync(t.launcher.replace('/resources/', '/old-resources/'), t.link)
    expect(t.install()).toEqual({ code: 0, stderr: '' })
    expect(readlinkSync(t.link)).toBe(t.launcher)
  })

  it('leaves a foreign command or symlink alone and says so', () => {
    const t = makeRoot()
    writeFileSync(t.link, '#!/bin/sh\necho other\n')
    const r = t.install()
    expect(r.code).toBe(0)
    expect(r.stderr).toContain('another program')
    expect(lstatSync(t.link).isSymbolicLink()).toBe(false)
    expect(readFileSync(t.link, 'utf-8')).toContain('other')
    expect(t.remove('remove').code).toBe(0)
    expect(existsSync(t.link)).toBe(true)

    rmSync(t.link)
    const vendor = join(t.launcher, '..', '..', '..', '..', 'vendor', 'cli', 'genoffice')
    mkdirSync(join(vendor, '..'), { recursive: true })
    writeFileSync(vendor, '#!/bin/sh\n')
    symlinkSync(vendor, t.link)
    expect(t.install().stderr).toContain('another program')
    expect(readlinkSync(t.link)).toBe(vendor)
    expect(t.remove('remove').code).toBe(0)
    expect(readlinkSync(t.link)).toBe(vendor)
  })

  it('keeps the link on an upgrade and drops only ours on uninstall', () => {
    const t = makeRoot()
    t.install()
    expect(t.remove('1').code).toBe(0)
    expect(readlinkSync(t.link)).toBe(t.launcher)
    expect(t.remove('upgrade').code).toBe(0)
    expect(existsSync(t.link)).toBe(true)
    expect(t.remove('0').code).toBe(0)
    expect(existsSync(t.link)).toBe(false)
  })
})
