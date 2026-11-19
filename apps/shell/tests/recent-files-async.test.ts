import { afterEach, beforeEach, expect, it, vi } from 'vitest'
import { pageRecentPaths } from '../src/main/recent-files'

const fs = vi.hoisted(() => ({ stat: vi.fn(), statSync: vi.fn() }))
vi.mock('node:fs', () => ({ statSync: fs.statSync }))
vi.mock('node:fs/promises', () => ({ stat: fs.stat }))

beforeEach(() => {
  fs.stat.mockReset().mockResolvedValue({ mtimeMs: 123, size: 45 })
  fs.statSync.mockReset().mockImplementation(() => {
    // Models an automount syscall that stalls the main thread, without a mount.
    Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, 50)
    return { mtimeMs: 123, size: 45 }
  })
})
afterEach(() => vi.restoreAllMocks())

it('keeps the event loop responsive while a recent path probe is pending', async () => {
  let release!: (value: { mtimeMs: number; size: number }) => void
  fs.stat.mockReturnValueOnce(
    new Promise((resolve) => {
      release = resolve
    }),
  )
  const result = pageRecentPaths(['/offline/share.docx'], {}, new Set())
  expect(fs.statSync).not.toHaveBeenCalled()
  let settled = false
  void Promise.resolve(result).then(() => {
    settled = true
  })
  await new Promise<void>((resolve) => setImmediate(resolve))
  expect(settled).toBe(false)
  release({ mtimeMs: 321, size: 67 })
  expect((await result).entries[0]).toMatchObject({
    path: '/offline/share.docx',
    mtimeMs: 321,
    sizeBytes: 67,
  })
})

it('preserves path order, starred flags and unavailable entries with asynchronous probes', async () => {
  fs.stat.mockImplementation(async (path: string) => {
    if (path.includes('offline')) throw new Error('unavailable mount')
    return { mtimeMs: 123, size: 45 }
  })
  const { entries } = await pageRecentPaths(
    ['/offline/share.docx', '/local/book.xlsx'],
    {},
    new Set(['/offline/share.docx']),
  )
  expect(fs.statSync).not.toHaveBeenCalled()
  expect(entries).toEqual([
    {
      path: '/offline/share.docx',
      name: 'share.docx',
      ext: 'docx',
      mtimeMs: 0,
      sizeBytes: 0,
      starred: true,
      missing: true,
    },
    {
      path: '/local/book.xlsx',
      name: 'book.xlsx',
      ext: 'xlsx',
      mtimeMs: 123,
      sizeBytes: 45,
      starred: false,
    },
  ])
})

it('starts probes only for the filtered page being returned', async () => {
  const result = await pageRecentPaths(
    ['/a.docx', '/offline/b.xlsx', '/c.docx'],
    { ext: 'docx', offset: 1, limit: 1 },
    new Set(),
  )
  expect(fs.stat).toHaveBeenCalledTimes(1)
  expect(fs.stat).toHaveBeenCalledWith('/c.docx')
  expect(result).toMatchObject({ total: 2, totalAll: 3, entries: [{ path: '/c.docx' }] })
})

it('counts a zero-length page without probing paths', async () => {
  expect(await pageRecentPaths(['/offline/share.docx'], { limit: 0 }, new Set())).toEqual({
    entries: [],
    total: 1,
    totalAll: 1,
  })
  expect(fs.stat).not.toHaveBeenCalled()
  expect(fs.statSync).not.toHaveBeenCalled()
})
