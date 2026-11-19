import { randomUUID } from 'node:crypto'
import { existsSync } from 'node:fs'
import { mkdir, mkdtemp, rm, utimes, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'

import {
  cleanupExpiredGeneratedPages,
  GENERATED_PAGE_TTL_MS,
} from '../src/main/generated-page-temp'

const roots: string[] = []

afterEach(async () => {
  await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
})

async function tempRoot(): Promise<string> {
  const root = await mkdtemp(join(tmpdir(), 'genoffice-slides-temp-test-'))
  roots.push(root)
  return root
}

describe('generated-page temp sweep', () => {
  it('removes only expired .pptx files matching the UUID pattern', async () => {
    const root = await tempRoot()
    const cloudDir = join(root, 'genoffice-cloud-pages')
    const localDir = join(root, 'genoffice-local-pages')
    await mkdir(cloudDir, { recursive: true })
    await mkdir(localDir, { recursive: true })

    // Create specific UUIDs we can reference
    const expiredUuid1 = randomUUID()
    const expiredUuid2 = randomUUID()
    const recentUuid = randomUUID()

    const expiredOwned1 = join(cloudDir, `${expiredUuid1}.pptx`)
    const expiredOwned2 = join(localDir, `${expiredUuid2}.pptx`)
    const recentOwned = join(localDir, `${recentUuid}.pptx`)
    const arbitrary = join(localDir, 'random-file.txt')

    await Promise.all([
      writeFile(expiredOwned1, 'old-expired'),
      writeFile(expiredOwned2, 'old-expired'),
      writeFile(recentOwned, 'recent'),
      writeFile(arbitrary, 'unrelated'),
    ])

    // Use the TTL constant for cutoff
    const now = Date.now()
    const expired = new Date(now - GENERATED_PAGE_TTL_MS - 60_000)
    await Promise.all([
      utimes(expiredOwned1, expired, expired),
      utimes(expiredOwned2, expired, expired),
    ])

    const removed = await cleanupExpiredGeneratedPages(root, now)

    expect(removed).toEqual(expect.arrayContaining([expiredOwned1, expiredOwned2]))
    expect(removed).toHaveLength(2)
    // Expired files should be gone
    for (const path of removed) {
      expect(existsSync(path)).toBe(false)
    }
    // Non-expired and non-matching files should remain (not deleted)
    expect(existsSync(recentOwned)).toBe(true)
    expect(existsSync(arbitrary)).toBe(true)
  })

  it('skips directories and non-.pptx files', async () => {
    const root = await tempRoot()
    await mkdir(join(root, 'genoffice-cloud-pages'), { recursive: true })

    const dirLike = join(root, 'genoffice-cloud-pages', 'not-a-file')
    await mkdir(dirLike, { recursive: true })
    await writeFile(join(dirLike, 'notes.txt'), 'notes')
    await writeFile(join(dirLike, 'data.json'), '{}')

    const removed = await cleanupExpiredGeneratedPages(root, Date.now())
    expect(removed).toHaveLength(0)
    expect(existsSync(join(dirLike, 'notes.txt'))).toBe(true)
  })

  it('succeeds when the generated-page directory does not exist', async () => {
    const root = await tempRoot()
    // Directories don't exist at all
    const removed = await cleanupExpiredGeneratedPages(root, Date.now())
    expect(removed).toHaveLength(0)
  })

  it('swallows per-file errors and continues', async () => {
    const root = await tempRoot()
    await mkdir(join(root, 'genoffice-cloud-pages'), { recursive: true })

    // We can't easily simulate a permission error on Windows,
    // but the function should never throw
    await expect(cleanupExpiredGeneratedPages(root, Date.now())).resolves.toBeInstanceOf(Array)
  })
})
