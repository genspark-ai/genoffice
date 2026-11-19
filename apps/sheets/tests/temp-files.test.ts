import { randomUUID } from 'node:crypto'
import { existsSync } from 'node:fs'
import { mkdir, mkdtemp, rm, utimes, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { afterEach, describe, expect, it } from 'vitest'

import {
  cleanupExpiredPastedFiles,
  cleanupImportTempDirectory,
  cleanupSessionResources,
  PASTED_FILE_TTL_MS,
} from '../src/main/temp-files'

const roots: string[] = []

afterEach(async () => {
  await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
})

async function tempRoot(): Promise<string> {
  const root = await mkdtemp(join(tmpdir(), 'genoffice-sheets-temp-test-'))
  roots.push(root)
  return root
}

describe('session temporary resources', () => {
  it('closes the sidecar before removing the snapshot and converted import directory', async () => {
    const root = await tempRoot()
    const snapshotDir = join(root, 'genoffice-sheets-sessions')
    const importDir = join(root, 'genoffice-imports', randomUUID())
    const snapshotPath = join(snapshotDir, `${randomUUID()}.xlsx`)
    await mkdir(importDir, { recursive: true })
    await mkdir(snapshotDir, { recursive: true })
    await writeFile(snapshotPath, 'snapshot')
    await writeFile(join(importDir, 'converted.xlsx'), 'converted')
    let closeObservedOpenFiles = false

    await cleanupSessionResources({
      tempRoot: root,
      snapshotPath,
      importTempDir: importDir,
      closeSidecar: async () => {
        closeObservedOpenFiles = existsSync(snapshotPath) && existsSync(importDir)
        throw new Error('sidecar already stopped')
      },
    })

    expect(closeObservedOpenFiles).toBe(true)
    expect(existsSync(snapshotPath)).toBe(false)
    expect(existsSync(importDir)).toBe(false)
  })

  it('never recursively deletes an unowned import path', async () => {
    const root = await tempRoot()
    const arbitrary = join(root, 'user-data')
    await mkdir(arbitrary, { recursive: true })
    await writeFile(join(arbitrary, 'keep.txt'), 'keep')

    await cleanupImportTempDirectory(root, arbitrary)

    expect(existsSync(join(arbitrary, 'keep.txt'))).toBe(true)
  })
})

describe('pasted-image TTL cleanup', () => {
  it('removes only expired app-owned files from the direct pasted temp directory', async () => {
    const root = await tempRoot()
    const pastedDir = join(root, 'genoffice-pasted')
    const oldOwned = join(pastedDir, 'pasted-20260101-010203-1.png')
    const recentOwned = join(pastedDir, 'pasted-20260101-010203-2.jpeg')
    const arbitrary = join(pastedDir, 'vacation.png')
    const lookalikeDirectory = join(pastedDir, 'pasted-20260101-010203-3.png')
    const nestedOwned = join(lookalikeDirectory, 'pasted-20260101-010203-4.png')
    await mkdir(lookalikeDirectory, { recursive: true })
    await Promise.all([
      writeFile(oldOwned, 'old'),
      writeFile(recentOwned, 'recent'),
      writeFile(arbitrary, 'user'),
      writeFile(nestedOwned, 'nested'),
    ])
    const now = Date.now()
    const expired = new Date(now - PASTED_FILE_TTL_MS - 60_000)
    await Promise.all([
      utimes(oldOwned, expired, expired),
      utimes(arbitrary, expired, expired),
      utimes(nestedOwned, expired, expired),
    ])

    const removed = await cleanupExpiredPastedFiles(root, now)

    expect(removed).toEqual([oldOwned])
    expect(existsSync(oldOwned)).toBe(false)
    expect(existsSync(recentOwned)).toBe(true)
    expect(existsSync(arbitrary)).toBe(true)
    expect(existsSync(nestedOwned)).toBe(true)
  })
})
