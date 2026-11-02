import { spawnSync } from 'node:child_process'
import {
  chmodSync,
  copyFileSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  statSync,
  writeFileSync,
} from 'node:fs'
import { createRequire } from 'node:module'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { runInNewContext } from 'node:vm'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

const SHELL_DIR = resolve(import.meta.dirname, '..')
const BUILD_DIR = join(SHELL_DIR, 'build')
const TEMPLATE_DIR = join(BUILD_DIR, 'shell-new')

// KDE Plasma's Dolphin builds its "Create New" menu from the XDG templates
// directories, which include the system-wide /usr/share/templates. The deb/rpm
// post-install script installs one blank document per format there
// (genoffice#2384); the menu label is the installed file name.
const TEMPLATES = {
  docx: 'GenOffice Document.docx',
  xlsx: 'GenOffice Spreadsheet.xlsx',
  pptx: 'GenOffice Presentation.pptx',
} as const

const INSTALL_SCRIPT = join(BUILD_DIR, 'linux-after-install.sh')
const REMOVE_SCRIPT = join(BUILD_DIR, 'linux-after-remove.sh')

const require = createRequire(import.meta.url)
// Inspect packaging metadata without downloading Electron or compiling OCR
// helpers (same approach as shell-new.test.ts).
const configModule = { exports: {} as ReturnType<typeof require> }
runInNewContext(readFileSync(join(SHELL_DIR, 'electron-builder.cjs'), 'utf8'), {
  module: configModule,
  __dirname: SHELL_DIR,
  process: { platform: 'linux', arch: 'x64', env: {} },
  require: (id: string) =>
    id === 'node:fs' ? { ...require(id), existsSync: () => true } : require(id),
})
const config = configModule.exports

describe('Linux package: KDE Create New templates (genoffice#2384)', () => {
  // electron-builder only logs a missing extraResources source and still exits
  // 0, so a dropped shell-new entry would ship a deb/rpm whose post-install
  // silently finds nothing to install.
  it('ships the blank templates in the Linux payload', () => {
    const entry = (
      config.linux.extraResources as { from: string; to: string; filter?: string[] }[]
    ).find((candidate) => candidate.to === 'shell-new')
    expect(entry, 'linux must ship shell-new like win does').toBeDefined()
    expect(entry?.from).toBe('build/shell-new')
    expect(existsSync(join(SHELL_DIR, entry?.from ?? ''))).toBe(true)
    expect([...(entry?.filter ?? [])].sort()).toEqual(['*.docx', '*.pptx', '*.xlsx'])
  })

  it('runs the template scripts as the deb/rpm install and uninstall hooks', () => {
    for (const target of [config.deb, config.rpm]) {
      expect(target.afterInstall).toBe('build/linux-after-install.sh')
      expect(target.afterRemove).toBe('build/linux-after-remove.sh')
    }
  })

  it('keeps the checked-in templates as valid OOXML ZIP packages', async () => {
    expect(readdirSync(TEMPLATE_DIR).sort()).toEqual(['blank.docx', 'blank.pptx', 'blank.xlsx'])
    for (const ext of Object.keys(TEMPLATES)) {
      const bytes = readFileSync(join(TEMPLATE_DIR, `blank.${ext}`))
      expect(bytes.subarray(0, 2).toString('latin1')).toBe('PK')
      const zip = await JSZip.loadAsync(bytes, { checkCRC32: true })
      expect(zip.file('[Content_Types].xml'), `blank.${ext} is not an OOXML package`).not.toBeNull()
    }
  })

  it('installs the three menu names from the packaged payload', () => {
    const script = readFileSync(INSTALL_SCRIPT, 'utf-8')
    expect(script).toContain('templates_dir="/usr/share/templates"')
    expect(script).toContain('$app/resources/shell-new')
    expect(script).toContain('mkdir -p "$templates_dir"')
    expect(script).toContain('install -m 0644')
    expect(script).toContain('cp -f')
    expect(script).toContain('blank.$ext')
    for (const name of Object.values(TEMPLATES)) {
      expect(script, `post-install must install ${name}`).toContain(name)
    }
  })

  it('removes the same three names on a real uninstall', () => {
    const script = readFileSync(REMOVE_SCRIPT, 'utf-8')
    expect(script).toContain('"/usr/share/templates/$name"')
    for (const name of Object.values(TEMPLATES)) {
      expect(script, `post-remove must remove ${name}`).toContain(name)
    }
  })
})

describe.skipIf(process.platform === 'win32')('Linux post-install scripts in a sandbox', () => {
  // The scripts hardcode the package paths; a copy with the install root and
  // the system templates directory swapped runs against a temp tree, so the
  // test needs no root and touches no real system directory.
  function makeRoot(withPayload: boolean) {
    const root = mkdtempSync(join(tmpdir(), 'genoffice-kde-templates-'))
    const app = join(root, 'opt', 'GenOffice')
    const bin = join(root, 'usr', 'bin')
    const templates = join(root, 'usr', 'share', 'templates')
    const launcher = join(app, 'resources', 'cli', 'genoffice')
    mkdirSync(join(app, 'resources', 'cli'), { recursive: true })
    mkdirSync(bin, { recursive: true })
    writeFileSync(launcher, '#!/bin/sh\n')
    chmodSync(launcher, 0o755)
    if (withPayload) {
      const shellNew = join(app, 'resources', 'shell-new')
      mkdirSync(shellNew, { recursive: true })
      for (const ext of Object.keys(TEMPLATES)) {
        copyFileSync(join(TEMPLATE_DIR, `blank.${ext}`), join(shellNew, `blank.${ext}`))
      }
    }
    const script = (name: string) => {
      const file = join(root, name)
      writeFileSync(
        file,
        readFileSync(join(BUILD_DIR, name), 'utf-8')
          .replaceAll('/opt/GenOffice', app)
          .replaceAll('/usr/share/templates', templates)
          .replaceAll('/usr/bin/', `${bin}/`),
      )
      return (arg = '') => {
        const result = spawnSync('sh', [file, arg], { encoding: 'utf-8' })
        return { code: result.status, stderr: result.stderr }
      }
    }
    return {
      templates,
      install: script('linux-after-install.sh'),
      remove: script('linux-after-remove.sh'),
    }
  }

  it('keeps both scripts valid POSIX sh', () => {
    for (const script of [INSTALL_SCRIPT, REMOVE_SCRIPT]) {
      const result = spawnSync('sh', ['-n', script], { encoding: 'utf-8' })
      expect(result.stderr, `${script} is not valid sh`).toBe('')
      expect(result.status).toBe(0)
    }
  })

  it('installs the three templates with the KDE menu names and 0644 mode', () => {
    const root = makeRoot(true)
    expect(root.install()).toEqual({ code: 0, stderr: '' })
    for (const [ext, name] of Object.entries(TEMPLATES)) {
      const installed = join(root.templates, name)
      expect(existsSync(installed), `${name} was not installed`).toBe(true)
      expect(readFileSync(installed).equals(readFileSync(join(TEMPLATE_DIR, `blank.${ext}`)))).toBe(
        true,
      )
      expect(statSync(installed).mode & 0o777).toBe(0o644)
    }
    // Reinstalling (an upgrade) rewrites the same three files, nothing else.
    expect(root.install()).toEqual({ code: 0, stderr: '' })
    expect(readdirSync(root.templates).sort()).toEqual(Object.values(TEMPLATES).sort())
  })

  it('skips the templates when the payload has no shell-new directory', () => {
    const root = makeRoot(false)
    expect(root.install()).toEqual({ code: 0, stderr: '' })
    expect(existsSync(root.templates)).toBe(false)
  })

  it('removes only our names on uninstall and keeps foreign templates', () => {
    const root = makeRoot(true)
    root.install()
    const foreign = join(root.templates, 'Other Suite Document.docx')
    writeFileSync(foreign, 'not ours')
    expect(root.remove('remove')).toEqual({ code: 0, stderr: '' })
    expect(readdirSync(root.templates)).toEqual(['Other Suite Document.docx'])
    expect(existsSync(foreign)).toBe(true)
  })

  it('keeps the templates through an upgrade and deletes them on a real uninstall', () => {
    const root = makeRoot(true)
    root.install()
    // rpm runs the old package's %postun (arg 1) and deb passes "upgrade"
    // after the new %post already reinstalled the files: neither may delete.
    expect(root.remove('1').code).toBe(0)
    expect(root.remove('upgrade').code).toBe(0)
    for (const name of Object.values(TEMPLATES)) {
      expect(existsSync(join(root.templates, name))).toBe(true)
    }
    expect(root.remove('0').code).toBe(0)
    for (const name of Object.values(TEMPLATES)) {
      expect(existsSync(join(root.templates, name))).toBe(false)
    }
    // The shared system templates directory is not ours to remove.
    expect(existsSync(root.templates)).toBe(true)
  })
})
