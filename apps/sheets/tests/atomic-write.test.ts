import { mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs'
import { rename, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'

import { atomicWriteFile } from '../src/main/atomic-write'

vi.mock('node:fs/promises', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs/promises')>()
  return { ...actual, rename: vi.fn(actual.rename), writeFile: vi.fn(actual.writeFile) }
})

let dir = ''

afterEach(() => {
  if (dir) rmSync(dir, { recursive: true, force: true })
  dir = ''
  vi.mocked(rename).mockClear()
  vi.mocked(writeFile).mockClear()
})

const epermError = () =>
  Object.assign(new Error('EPERM: operation not permitted'), { code: 'EPERM' })

describe('atomicWriteFile', () => {
  it('replaces the target and leaves no temp file behind', async () => {
    dir = mkdtempSync(join(tmpdir(), 'sheets-aw-'))
    const target = join(dir, 'data.csv')
    writeFileSync(target, 'old')

    await atomicWriteFile(target, Buffer.from('new'))

    expect(readFileSync(target, 'utf8')).toBe('new')
    expect(readdirSync(dir)).toEqual(['data.csv'])
  })

  it('retries transient Windows rename locks', async () => {
    dir = mkdtempSync(join(tmpdir(), 'sheets-aw-'))
    const target = join(dir, 'data.csv')
    writeFileSync(target, 'old')
    vi.mocked(rename).mockRejectedValueOnce(epermError()).mockRejectedValueOnce(epermError())

    await atomicWriteFile(target, Buffer.from('new'))

    expect(vi.mocked(rename)).toHaveBeenCalledTimes(3)
    expect(readFileSync(target, 'utf8')).toBe('new')
    expect(readdirSync(dir)).toEqual(['data.csv'])
  })

  it('falls back to an in-place write when the rename stays locked', async () => {
    dir = mkdtempSync(join(tmpdir(), 'sheets-aw-'))
    const target = join(dir, 'data.csv')
    writeFileSync(target, 'old')
    for (let attempt = 0; attempt < 5; attempt += 1) {
      vi.mocked(rename).mockRejectedValueOnce(epermError())
    }

    await atomicWriteFile(target, Buffer.from('new'))

    expect(readFileSync(target, 'utf8')).toBe('new')
    expect(readdirSync(dir)).toEqual(['data.csv'])
  })

  it('preserves the completed temp file when the fallback write fails', async () => {
    dir = mkdtempSync(join(tmpdir(), 'sheets-aw-'))
    const target = join(dir, 'data.csv')
    writeFileSync(target, 'old')
    for (let attempt = 0; attempt < 5; attempt += 1) {
      vi.mocked(rename).mockRejectedValueOnce(epermError())
    }
    const actual = await vi.importActual<typeof import('node:fs/promises')>('node:fs/promises')
    const fallbackError = Object.assign(new Error('EIO: fallback write failed'), { code: 'EIO' })
    vi.mocked(writeFile)
      .mockImplementationOnce(actual.writeFile)
      .mockRejectedValueOnce(fallbackError)

    await expect(atomicWriteFile(target, Buffer.from('new'))).rejects.toThrow(
      'fallback write failed',
    )

    const files = readdirSync(dir)
    expect(files).toContain('data.csv')
    const temp = files.find((file) => file !== 'data.csv')
    expect(temp).toBeDefined()
    expect(readFileSync(join(dir, temp!), 'utf8')).toBe('new')
  })

  it('uses distinct temp files for concurrent writes to one target', async () => {
    dir = mkdtempSync(join(tmpdir(), 'sheets-aw-'))
    const target = join(dir, 'data.csv')

    await Promise.all([
      atomicWriteFile(target, Buffer.from('first')),
      atomicWriteFile(target, Buffer.from('second')),
    ])

    expect(['first', 'second']).toContain(readFileSync(target, 'utf8'))
    expect(readdirSync(dir)).toEqual(['data.csv'])
  })
})
