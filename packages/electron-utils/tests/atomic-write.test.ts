import {
  fsyncSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  renameSync,
  rmSync,
  writeFileSync,
} from 'node:fs'
import { open, rename, writeFile, copyFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, dirname, join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  atomicCopyFile,
  atomicWriteFile,
  atomicWriteVia,
  writeJsonAtomic,
} from '../src/atomic-write'

vi.mock('node:fs/promises', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs/promises')>()
  return {
    ...actual,
    rename: vi.fn(actual.rename),
    open: vi.fn(actual.open),
    writeFile: vi.fn(actual.writeFile),
    copyFile: vi.fn(actual.copyFile),
  }
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

  // A gvfs SMB share refuses rename-over-existing with EEXIST — outside the
  // Windows retry set. The save must still land through the in-place fallback.
  it('falls back to an in-place write when the mount refuses the rename', async () => {
    dir = mkdtempSync(join(tmpdir(), 'eu-aw-'))
    const target = join(dir, 'a.docx')
    writeFileSync(target, 'old')
    vi.mocked(rename).mockRejectedValueOnce(errnoError('EEXIST'))
    await atomicWriteFile(target, Buffer.from('new'))
    expect(readFileSync(target, 'utf8')).toBe('new')
    expect(readdirSync(dir)).toEqual(['a.docx'])
  })

  it('surfaces the fallback write failure and keeps the completed temp', async () => {
    dir = mkdtempSync(join(tmpdir(), 'eu-aw-'))
    const target = join(dir, 'a.docx')
    writeFileSync(target, 'old')
    vi.mocked(rename).mockRejectedValueOnce(errnoError('EEXIST'))
    const actual = await vi.importActual<typeof import('node:fs/promises')>('node:fs/promises')
    // call 1 writes the temp for real, call 2 is the in-place fallback (fails)
    vi.mocked(writeFile).mockImplementationOnce((...args) =>
      actual.writeFile(...(args as Parameters<typeof actual.writeFile>)),
    )
    vi.mocked(writeFile).mockRejectedValueOnce(errnoError('EACCES'))
    await expect(atomicWriteFile(target, Buffer.from('new'))).rejects.toThrow('EACCES')
    expect(readFileSync(target, 'utf8')).toBe('old')
    // the fallback failed, so the completed temp is kept as the rescue copy
    const leftovers = readdirSync(dir).sort()
    expect(leftovers).toHaveLength(2)
    expect(leftovers[0]).toBe('a.docx')
    expect(leftovers[1]).toMatch(/^a\.docx\..{12}\.tmp$/)
  })
})

describe('atomicCopyFile', () => {
  it('fsyncs the copied temp file before the rename publishes it', async () => {
    dir = mkdtempSync(join(tmpdir(), 'eu-ac-'))
    const source = join(dir, 'a.pdf')
    const target = join(dir, 'b.pdf')
    writeFileSync(source, 'pdf-bytes')
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
    await atomicCopyFile(source, target)
    expect(order).toEqual(['sync', 'rename'])
    expect(readFileSync(target, 'utf8')).toBe('pdf-bytes')
    expect(readdirSync(dir).sort()).toEqual(['a.pdf', 'b.pdf'])
  })

  // the same tolerated-flush contract as atomicWriteFile
  it('tolerates a refused flush and still publishes the copy', async () => {
    dir = mkdtempSync(join(tmpdir(), 'eu-ac-'))
    const source = join(dir, 'a.pdf')
    const target = join(dir, 'b.pdf')
    writeFileSync(source, 'pdf-bytes')
    writeFileSync(target, 'old')
    vi.mocked(open).mockRejectedValueOnce(errnoError('EPERM'))
    await atomicCopyFile(source, target)
    expect(readFileSync(target, 'utf8')).toBe('pdf-bytes')
    expect(readdirSync(dir).sort()).toEqual(['a.pdf', 'b.pdf'])
  })

  it('leaves the target intact and drops no temp when the source is missing', async () => {
    dir = mkdtempSync(join(tmpdir(), 'eu-ac-'))
    const target = join(dir, 'b.pdf')
    writeFileSync(target, 'old')
    await expect(atomicCopyFile(join(dir, 'nope.pdf'), target)).rejects.toThrow()
    expect(readFileSync(target, 'utf8')).toBe('old')
    expect(readdirSync(dir)).toEqual(['b.pdf'])
  })

  // Same fallback contract as atomicWriteFile: a mount that refuses
  // rename-over-existing still receives the copy, written in place.
  it('falls back to an in-place copy when the mount refuses the rename', async () => {
    dir = mkdtempSync(join(tmpdir(), 'eu-ac-'))
    const source = join(dir, 'a.pdf')
    const target = join(dir, 'b.pdf')
    writeFileSync(source, 'pdf-bytes')
    writeFileSync(target, 'old')
    vi.mocked(rename).mockRejectedValueOnce(errnoError('EEXIST'))
    await atomicCopyFile(source, target)
    expect(readFileSync(target, 'utf8')).toBe('pdf-bytes')
    expect(readdirSync(dir).sort()).toEqual(['a.pdf', 'b.pdf'])
  })
})

// A streamed deck save (JSZip → createWriteStream) cannot go through the
// Buffer-taking atomicWriteFile; the writer fills the temp itself and the same
// sync → rename → in-place fallback contract publishes it (genoffice#2123).
describe('atomicWriteVia', () => {
  it('hands the writer a sibling temp, then fsyncs it before the rename publishes it', async () => {
    dir = mkdtempSync(join(tmpdir(), 'eu-av-'))
    const target = join(dir, 'a.pptx')
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
    await atomicWriteVia(target, async (tempPath) => {
      order.push('write')
      // beside the target, so the rename stays on one filesystem; never the target itself
      expect(dirname(tempPath)).toBe(dir)
      expect(tempPath).not.toBe(target)
      expect(basename(tempPath)).toMatch(/a\.pptx\..{12}\.tmp$/)
      expect(readFileSync(target, 'utf8')).toBe('old')
      await actual.writeFile(tempPath, 'new')
    })
    expect(order).toEqual(['write', 'sync', 'rename'])
    expect(readFileSync(target, 'utf8')).toBe('new')
    expect(readdirSync(dir)).toEqual(['a.pptx'])
  })

  // The regression: the stream used to open the deck itself with 'w', so a
  // save that died mid-way left a truncated .pptx with no copy to fall back to.
  it('a writer that dies part-way leaves the previous file intact and no temp behind', async () => {
    dir = mkdtempSync(join(tmpdir(), 'eu-av-'))
    const target = join(dir, 'a.pptx')
    writeFileSync(target, 'the whole old deck')
    const actual = await vi.importActual<typeof import('node:fs/promises')>('node:fs/promises')
    await expect(
      atomicWriteVia(target, async (tempPath) => {
        await actual.writeFile(tempPath, 'half a new d')
        throw errnoError('EIO')
      }),
    ).rejects.toThrow('EIO')
    expect(readFileSync(target, 'utf8')).toBe('the whole old deck')
    expect(readdirSync(dir)).toEqual(['a.pptx'])
  })

  it('tolerates a refused flush and still renames atomically', async () => {
    dir = mkdtempSync(join(tmpdir(), 'eu-av-'))
    const target = join(dir, 'a.pptx')
    writeFileSync(target, 'old')
    const actual = await vi.importActual<typeof import('node:fs/promises')>('node:fs/promises')
    vi.mocked(open).mockRejectedValueOnce(errnoError('EPERM'))
    await atomicWriteVia(target, (tempPath) => actual.writeFile(tempPath, 'new'))
    expect(readFileSync(target, 'utf8')).toBe('new')
    expect(readdirSync(dir)).toEqual(['a.pptx'])
  })

  // Same fallback contract as atomicWriteFile (genoffice#1878): a mount that refuses
  // rename-over-existing still receives the completed temp, copied in place.
  it('falls back to an in-place copy when the mount refuses the rename', async () => {
    dir = mkdtempSync(join(tmpdir(), 'eu-av-'))
    const target = join(dir, 'a.pptx')
    writeFileSync(target, 'old')
    const actual = await vi.importActual<typeof import('node:fs/promises')>('node:fs/promises')
    vi.mocked(rename).mockRejectedValueOnce(errnoError('EEXIST'))
    await atomicWriteVia(target, (tempPath) => actual.writeFile(tempPath, 'new'))
    expect(readFileSync(target, 'utf8')).toBe('new')
    expect(readdirSync(dir)).toEqual(['a.pptx'])
  })

  it('surfaces the fallback copy failure and keeps the completed temp', async () => {
    dir = mkdtempSync(join(tmpdir(), 'eu-av-'))
    const target = join(dir, 'a.pptx')
    writeFileSync(target, 'old')
    const actual = await vi.importActual<typeof import('node:fs/promises')>('node:fs/promises')
    vi.mocked(rename).mockRejectedValueOnce(errnoError('EEXIST'))
    vi.mocked(copyFile).mockRejectedValueOnce(errnoError('EACCES'))
    await expect(
      atomicWriteVia(target, (tempPath) => actual.writeFile(tempPath, 'new')),
    ).rejects.toThrow('EACCES')
    expect(readFileSync(target, 'utf8')).toBe('old')
    // the fallback failed, so the completed temp is kept as the rescue copy
    const leftovers = readdirSync(dir).sort((a, b) => a.length - b.length)
    expect(leftovers).toHaveLength(2)
    expect(leftovers[0]).toBe('a.pptx')
    expect(leftovers[1]).toMatch(/a\.pptx\..{12}\.tmp$/)
    expect(readFileSync(join(dir, leftovers[1]!), 'utf8')).toBe('new')
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
