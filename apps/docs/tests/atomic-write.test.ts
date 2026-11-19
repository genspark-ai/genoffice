import { mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs'
import { rename } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { atomicWriteFile, looksLikeZip } from '../src/main/atomic-write'

vi.mock('node:fs/promises', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs/promises')>()
  return { ...actual, rename: vi.fn(actual.rename) }
})

let dir: string

afterEach(() => {
  rmSync(dir, { recursive: true, force: true })
  vi.mocked(rename).mockClear()
})

const epermError = () =>
  Object.assign(new Error('EPERM: operation not permitted'), { code: 'EPERM' })

describe('atomicWriteFile', () => {
  it('replaces the target and leaves no temp file behind', async () => {
    dir = mkdtempSync(join(tmpdir(), 'aw-'))
    const target = join(dir, 'a.docx')
    writeFileSync(target, 'old')
    await atomicWriteFile(target, Buffer.from('new'))
    expect(readFileSync(target, 'utf-8')).toBe('new')
    expect(readdirSync(dir)).toEqual(['a.docx'])
  })

  it('keeps the original intact and cleans the temp file when the write fails', async () => {
    dir = mkdtempSync(join(tmpdir(), 'aw-'))
    const target = join(dir, 'a.docx')
    writeFileSync(target, 'old')
    const bad = { byteLength: -1 } as unknown as Buffer
    await expect(atomicWriteFile(target, bad)).rejects.toThrow()
    expect(readFileSync(target, 'utf-8')).toBe('old')
    expect(readdirSync(dir)).toEqual(['a.docx'])
  })

  // Windows: Defender/indexer briefly locks the target and rename throws EPERM
  it('retries a transiently locked rename and still lands atomically', async () => {
    dir = mkdtempSync(join(tmpdir(), 'aw-'))
    const target = join(dir, 'a.docx')
    writeFileSync(target, 'old')
    vi.mocked(rename).mockRejectedValueOnce(epermError()).mockRejectedValueOnce(epermError())
    await atomicWriteFile(target, Buffer.from('new'))
    expect(readFileSync(target, 'utf-8')).toBe('new')
    expect(readdirSync(dir)).toEqual(['a.docx'])
  })

  it('falls back to an in-place write when the rename stays locked', async () => {
    dir = mkdtempSync(join(tmpdir(), 'aw-'))
    const target = join(dir, 'a.docx')
    writeFileSync(target, 'old')
    for (let i = 0; i < 5; i++) vi.mocked(rename).mockRejectedValueOnce(epermError())
    await atomicWriteFile(target, Buffer.from('new'))
    expect(readFileSync(target, 'utf-8')).toBe('new')
    expect(readdirSync(dir)).toEqual(['a.docx'])
  })
})

describe('looksLikeZip', () => {
  it('accepts a zip local-file header and rejects truncated/garbage bytes', () => {
    dir = mkdtempSync(join(tmpdir(), 'aw-'))
    expect(looksLikeZip(Buffer.from([0x50, 0x4b, 0x03, 0x04, 0x00]))).toBe(true)
    expect(looksLikeZip(Buffer.from([0x50, 0x4b]))).toBe(false)
    expect(looksLikeZip(Buffer.from('not a zip'))).toBe(false)
    expect(looksLikeZip(Buffer.alloc(0))).toBe(false)
  })
})
