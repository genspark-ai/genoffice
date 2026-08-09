import { chmod, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { afterAll, beforeAll, describe, expect, it } from 'vitest'

import { syncFileBestEffort, writeXlsxAtomically } from '../src/gateway/xlsx-gateway'

describe('syncFileBestEffort', () => {
  let directory: string

  beforeAll(async () => {
    directory = await mkdtemp(join(tmpdir(), 'xlsx-atomic-sync-'))
  })

  afterAll(async () => {
    await rm(directory, { recursive: true, force: true })
  })

  it('flushes a writable file without altering it', async () => {
    const path = join(directory, 'plain.bin')
    await writeFile(path, 'payload')
    await syncFileBestEffort(path)
    expect(await readFile(path, 'utf8')).toBe('payload')
  })

  it('tolerates a file it cannot reopen for writing (cloud-sync/AV lock shape)', async () => {
    const path = join(directory, 'readonly.bin')
    await writeFile(path, 'payload')
    await chmod(path, 0o444)
    await expect(syncFileBestEffort(path)).resolves.toBeUndefined()
  })

  it('still surfaces a missing file', async () => {
    await expect(syncFileBestEffort(join(directory, 'absent.bin'))).rejects.toMatchObject({
      code: 'ENOENT',
    })
  })
})

describe('writeXlsxAtomically', () => {
  it('writes through a flushed temp file and renames it into place', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'xlsx-atomic-write-'))
    try {
      const path = join(directory, 'book.xlsx')
      await writeXlsxAtomically(path, Buffer.from('bytes'))
      expect(await readFile(path, 'utf8')).toBe('bytes')
    } finally {
      await rm(directory, { recursive: true, force: true })
    }
  })
})
