import { mkdirSync, mkdtempSync, rmSync, symlinkSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { isInDocDir, isInVaultRoot } from '../src/main/asset-lifecycle'

describe('isInDocDir (md-asset:// serve gate)', () => {
  it('allows sibling images of filesystem-root documents', () => {
    expect(isInDocDir('/a.png', '/', '/')).toBe(true)
    expect(isInDocDir('C:\\a.png', 'C:\\', '\\')).toBe(true)
  })

  it('allows siblings of nested documents and rejects the rest', () => {
    expect(isInDocDir('/docs/a.png', '/docs', '/')).toBe(true)
    expect(isInDocDir('/other/a.png', '/docs', '/')).toBe(false)
    // Prefix attacks: a sibling directory sharing the name prefix is outside
    expect(isInDocDir('/docs-evil/a.png', '/docs', '/')).toBe(false)
    expect(isInDocDir('C:\\docs\\a.png', 'C:\\docs', '\\')).toBe(true)
    expect(isInDocDir('C:\\docs-evil\\a.png', 'C:\\docs', '\\')).toBe(false)
  })

  it('never serves the directory itself', () => {
    expect(isInDocDir('/docs', '/docs', '/')).toBe(false)
    expect(isInDocDir('/', '/', '/')).toBe(false)
  })
})

let base = ''
let vault = ''
let outside = ''

beforeAll(() => {
  base = mkdtempSync(join(tmpdir(), 'md-asset-vault-'))
  vault = join(base, 'vault')
  outside = join(base, 'outside')
  // Obsidian-style layout: note in notes/, picture in vault-root attachments/
  mkdirSync(join(vault, 'notes'), { recursive: true })
  mkdirSync(join(vault, 'attachments'), { recursive: true })
  mkdirSync(join(vault, 'images'), { recursive: true })
  mkdirSync(outside, { recursive: true })
  writeFileSync(join(vault, 'notes', 'Index.md'), '![[img.png]]')
  writeFileSync(join(vault, 'attachments', 'img.png'), 'png')
  writeFileSync(join(vault, 'images', 'sibling.png'), 'png')
  writeFileSync(join(outside, 'secret.png'), 'png')
  if (process.platform !== 'win32') {
    symlinkSync(join(outside, 'secret.png'), join(vault, 'attachments', 'escape.png'))
  }
})
afterAll(() => rmSync(base, { recursive: true, force: true }))

describe('isInVaultRoot (cross-folder md-asset:// embeds)', () => {
  it('serves a picture in a root-level attachments folder from a subfolder note', async () => {
    expect(await isInVaultRoot(join(vault, 'attachments', 'img.png'), vault)).toBe(true)
  })

  it('serves a picture in a sibling subfolder', async () => {
    expect(await isInVaultRoot(join(vault, 'images', 'sibling.png'), vault)).toBe(true)
  })

  it('rejects files outside the vault and prefix-attack siblings', async () => {
    expect(await isInVaultRoot(join(outside, 'secret.png'), vault)).toBe(false)
    expect(await isInVaultRoot(join(`${vault}-evil`, 'img.png'), vault)).toBe(false)
  })

  it.skipIf(process.platform === 'win32')(
    'rejects a symlink inside the vault that escapes to outside',
    async () => {
      expect(await isInVaultRoot(join(vault, 'attachments', 'escape.png'), vault)).toBe(false)
    },
  )
})
