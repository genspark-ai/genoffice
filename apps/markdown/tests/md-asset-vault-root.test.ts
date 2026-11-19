import { mkdtempSync, mkdirSync, rmSync, symlinkSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { resolveSafeVaultImagePath } from '../src/main/asset-lifecycle'

/** symlinkSync needs privilege on Windows (developer mode); a machine without it skips those cases */
const canSymlink = (() => {
  try {
    const dir = mkdtempSync(join(tmpdir(), 'md-asset-link-'))
    symlinkSync(join(dir, 'a'), join(dir, 'a-link'))
    rmSync(dir, { recursive: true, force: true })
    return true
  } catch {
    return false
  }
})()

const cleanup: string[] = []

function vault(name: string): string {
  const dir = mkdtempSync(join(tmpdir(), `md-asset-${name}-`))
  cleanup.push(dir)
  return dir
}

afterEach(() => {
  for (const dir of cleanup.splice(0)) rmSync(dir, { recursive: true, force: true })
})

describe('resolveSafeVaultImagePath (md-asset serve gate widened to the vault root)', () => {
  it('serves an attachment from the vault root for a nested note (attachments/ layout)', async () => {
    const root = vault('attachments')
    mkdirSync(join(root, 'attachments'))
    mkdirSync(join(root, 'notes'))
    writeFileSync(join(root, 'notes', 'note.md'), 'text')
    writeFileSync(join(root, 'attachments', 'pic.png'), 'png')
    const served = await resolveSafeVaultImagePath(root, join(root, 'attachments', 'pic.png'))
    expect(served).toBeTruthy()
  })

  it('serves a picture in any vault subfolder, not only the document directory', async () => {
    const root = vault('media')
    mkdirSync(join(root, 'media'))
    writeFileSync(join(root, 'media', 'diagram.png'), 'png')
    const served = await resolveSafeVaultImagePath(root, join(root, 'media', 'diagram.png'))
    expect(served).toBeTruthy()
  })

  it('refuses a target outside the vault root', async () => {
    const root = vault('outside')
    const outsideSub = vault('elsewhere')
    mkdirSync(join(outsideSub, 'media'))
    writeFileSync(join(outsideSub, 'media', 'pic.png'), 'png')
    const served = await resolveSafeVaultImagePath(root, join(outsideSub, 'media', 'pic.png'))
    expect(served).toBeNull()
  })

  it('refuses .. traversal out of the vault root', async () => {
    const root = vault(' traversal')
    writeFileSync(join(root, 'pic.png'), 'png')
    const served = await resolveSafeVaultImagePath(root, join(root, '..', 'pic.png'))
    expect(served).toBeNull()
  })

  it('refuses the vault root itself and a directory target', async () => {
    const root = vault('dirs')
    expect(await resolveSafeVaultImagePath(root, root)).toBeNull()
    mkdirSync(join(root, 'media'))
    expect(await resolveSafeVaultImagePath(root, join(root, 'media'))).toBeNull()
  })

  it.skipIf(!canSymlink)('refuses a symlink inside the vault that escapes the root', async () => {
    const root = vault('symlink-escape')
    const outside = vault('secret')
    writeFileSync(join(outside, 'pic.png'), 'png')
    mkdirSync(join(root, 'attachments'))
    symlinkSync(join(outside, 'pic.png'), join(root, 'attachments', 'escape.png'))
    const served = await resolveSafeVaultImagePath(root, join(root, 'attachments', 'escape.png'))
    expect(served).toBeNull()
  })

  it.skipIf(!canSymlink)(
    'still serves a symlink whose real target stays inside the vault',
    async () => {
      const root = vault('symlink-inside')
      writeFileSync(join(root, 'pic.png'), 'png')
      mkdirSync(join(root, 'attachments'))
      symlinkSync(join(root, 'pic.png'), join(root, 'attachments', 'in-link.png'))
      const served = await resolveSafeVaultImagePath(root, join(root, 'attachments', 'in-link.png'))
      expect(served).toBeTruthy()
    },
  )
})
