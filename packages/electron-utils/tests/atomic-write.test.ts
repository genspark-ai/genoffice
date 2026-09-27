import {
  fsyncSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  renameSync,
  rmSync,
  writeFileSync,
} from 'node:fs'
import { open, rename } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { atomicWriteFile, writeJsonAtomic } from '../src/atomic-write'

vi.mock('node:fs/promises', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs/promises')>()
  return { ...actual, rename: vi.fn(actual.rename), open: vi.fn(actual.open) }
})

vi.mock('node:fs', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs')>()
  return {
    ...actual,
    writeFileSync: vi.fn(actual.writeFileSync),
    renameSync: vi.fn(actual.renameSync),
    fsyncSync: vi.fn(actual.fsyncSync),
  }
})

let dir = ''

afterEach(() => {
  if (dir) rmSync(dir, { recursive: true, force: true })
  dir = ''
  vi.clearAllMocks()
})

const errnoError = (code: string) => Object.assign(new Error(`${code}: injected`), { code })

describe('atomicWriteFile', () => {
  it('fsyncs the temp file before the rename publishes it', async () => {
    dir = mkdtempSync(join(tmpdir(), 'eu-aw-'))
    const target = join(dir, 'a.docx')
    writeFileSync(target, 'old')
    const order: string[] = []
    const actual = await vi.importActual<typeof import('node:fs/promises')>('node:fs/promises')
    vi.mocked(open).mockImplementationOnce(async (...args) => {
      const handle = await actual.open(...(args as Parameters<typeof actual.open>))
      const sync = handle.sync.bind(handle)
      handle.sync = async () => {
        order.push('sync')
        await sync()
      }
      return handle
    })
    vi.mocked(rename).mockImplementationOnce(async (from, to) => {
      order.push('rename')
      await actual.rename(from, to)
    })
    await atomicWriteFile(target, Buffer.from('new'))
    expect(order).toEqual(['sync', 'rename'])
    expect(readFileSync(target, 'utf8')).toBe('new')
    expect(readdirSync(dir)).toEqual(['a.docx'])
  })

  // OneDrive/AV-locked folders refuse fsync: the bytes are written, so the save must still land
  it('tolerates a refused flush and still renames atomically', async () => {
    dir = mkdtempSync(join(tmpdir(), 'eu-aw-'))
    const target = join(dir, 'a.docx')
    writeFileSync(target, 'old')
    vi.mocked(open).mockRejectedValueOnce(errnoError('EPERM'))
    await atomicWriteFile(target, Buffer.from('new'))
    expect(readFileSync(target, 'utf8')).toBe('new')
    expect(readdirSync(dir)).toEqual(['a.docx'])
  })

  it('surfaces a non-tolerated flush failure and leaves the target intact', async () => {
    dir = mkdtempSync(join(tmpdir(), 'eu-aw-'))
    const target = join(dir, 'a.docx')
    writeFileSync(target, 'old')
    vi.mocked(open).mockRejectedValueOnce(errnoError('EIO'))
    await expect(atomicWriteFile(target, Buffer.from('new'))).rejects.toThrow('EIO')
    expect(readFileSync(target, 'utf8')).toBe('old')
    expect(readdirSync(dir)).toEqual(['a.docx'])
  })

  it('retries a locked rename, then falls back to an in-place write', async () => {
    dir = mkdtempSync(join(tmpdir(), 'eu-aw-'))
    const target = join(dir, 'a.docx')
    writeFileSync(target, 'old')
    for (let i = 0; i < 5; i += 1) vi.mocked(rename).mockRejectedValueOnce(errnoError('EPERM'))
    await atomicWriteFile(target, Buffer.from('new'))
    expect(readFileSync(target, 'utf8')).toBe('new')
    expect(readdirSync(dir)).toEqual(['a.docx'])
  })
})

describe('writeJsonAtomic', () => {
  it('replaces the file and leaves no temp behind', () => {
    dir = mkdtempSync(join(tmpdir(), 'eu-json-'))
    const target = join(dir, 'state', 'recent.json')
    writeJsonAtomic(target, ['a'])
    writeJsonAtomic(target, ['b', 'a'])
    expect(JSON.parse(readFileSync(target, 'utf8'))).toEqual(['b', 'a'])
    expect(readdirSync(join(dir, 'state'))).toEqual(['recent.json'])
  })

  // Simulates a write interrupted after truncation: the previous state must survive
  it('keeps the previous state when the write dies half-way', () => {
    dir = mkdtempSync(join(tmpdir(), 'eu-json-'))
    const target = join(dir, 'starred.json')
    writeJsonAtomic(target, ['keep'])
    const actual = vi.mocked(writeFileSync).getMockImplementation()!
    vi.mocked(writeFileSync).mockImplementationOnce((path, _data, options) => {
      actual(path, '', options)
      throw errnoError('EIO')
    })
    expect(() => writeJsonAtomic(target, ['lost'])).toThrow('EIO')
    expect(JSON.parse(readFileSync(target, 'utf8'))).toEqual(['keep'])
    expect(readdirSync(dir)).toEqual(['starred.json'])
  })

  it('keeps the previous state when the rename fails', () => {
    dir = mkdtempSync(join(tmpdir(), 'eu-json-'))
    const target = join(dir, 'ai-settings.json')
    writeJsonAtomic(target, { provider: 'a' })
    vi.mocked(renameSync).mockImplementationOnce(() => {
      throw errnoError('EXDEV')
    })
    expect(() => writeJsonAtomic(target, { provider: 'b' })).toThrow('EXDEV')
    expect(JSON.parse(readFileSync(target, 'utf8'))).toEqual({ provider: 'a' })
    expect(readdirSync(dir)).toEqual(['ai-settings.json'])
  })

  it('retries a locked rename, then falls back to an in-place write', () => {
    dir = mkdtempSync(join(tmpdir(), 'eu-json-'))
    const target = join(dir, 'recent.json')
    writeJsonAtomic(target, ['old'])
    const locked = () => {
      throw errnoError('EPERM')
    }
    vi.mocked(renameSync).mockImplementationOnce(locked).mockImplementationOnce(locked)
    writeJsonAtomic(target, ['retried'])
    expect(vi.mocked(renameSync)).toHaveBeenCalledTimes(4)
    for (let i = 0; i < 5; i += 1) vi.mocked(renameSync).mockImplementationOnce(locked)
    writeJsonAtomic(target, ['fallback'])
    expect(JSON.parse(readFileSync(target, 'utf8'))).toEqual(['fallback'])
    expect(readdirSync(dir)).toEqual(['recent.json'])
  })

  it('tolerates a refused fsync on the temp file', () => {
    dir = mkdtempSync(join(tmpdir(), 'eu-json-'))
    const target = join(dir, 'ai-settings.json')
    vi.mocked(fsyncSync).mockImplementationOnce(() => {
      throw errnoError('EINVAL')
    })
    writeJsonAtomic(target, { ok: true })
    expect(JSON.parse(readFileSync(target, 'utf8'))).toEqual({ ok: true })
  })
})
