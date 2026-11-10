import { mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'

const trace = vi.hoisted(() => ({ order: [] as string[] }))

vi.mock('node:fs/promises', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs/promises')>()
  return {
    ...actual,
    open: vi.fn(async (...args: Parameters<typeof actual.open>) => {
      const handle = await actual.open(...args)
      const sync = handle.sync.bind(handle)
      handle.sync = async () => {
        trace.order.push('sync')
        return sync()
      }
      return handle
    }),
    rename: vi.fn(async (...args: Parameters<typeof actual.rename>) => {
      trace.order.push('rename')
      return actual.rename(...args)
    }),
  }
})

import { atomicCopyFile } from '../src/main/atomic-write'

describe('atomicCopyFile', () => {
  const dirs: string[] = []
  afterEach(() => {
    for (const d of dirs.splice(0)) rmSync(d, { recursive: true, force: true })
  })

  it('copies the bytes and leaves no temp file behind', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'atomic-copy-'))
    dirs.push(dir)
    const src = join(dir, 'a.pdf')
    writeFileSync(src, Buffer.from([1, 2, 3, 0, 255]))
    await atomicCopyFile(src, join(dir, 'a copy.pdf'))
    expect(readFileSync(join(dir, 'a copy.pdf'))).toEqual(readFileSync(src))
    expect(readdirSync(dir).sort()).toEqual(['a copy.pdf', 'a.pdf'])
  })

  it('flushes the temp file before the rename publishes the copy', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'atomic-copy-'))
    dirs.push(dir)
    const src = join(dir, 'a.pdf')
    writeFileSync(src, Buffer.from([1, 2, 3, 0, 255]))
    trace.order.length = 0
    await atomicCopyFile(src, join(dir, 'a copy.pdf'))
    expect(trace.order).toEqual(['sync', 'rename'])
  })

  it('rejects on a missing source without creating the target', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'atomic-copy-'))
    dirs.push(dir)
    await expect(atomicCopyFile(join(dir, 'nope'), join(dir, 'out'))).rejects.toThrow()
    expect(readdirSync(dir)).toEqual([])
  })
})
