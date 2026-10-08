import { mkdtempSync, readFileSync, readdirSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'

const synced = vi.hoisted(() => [] as string[])
const syncError = vi.hoisted(() => ({ code: '' as string }))

vi.mock('node:fs/promises', async (importOriginal) => {
  const real = await importOriginal<typeof import('node:fs/promises')>()
  return {
    ...real,
    open: async (path: string, flags: string) => {
      const handle = await real.open(path, flags)
      const sync = handle.sync.bind(handle)
      handle.sync = async () => {
        synced.push(path)
        if (syncError.code)
          throw Object.assign(new Error('fsync refused'), { code: syncError.code })
        await sync()
      }
      return handle
    },
  }
})

import { atomicWriteFile } from '../src/main/atomic-write'

afterEach(() => {
  synced.length = 0
  syncError.code = ''
})

describe('atomicWriteFile durability', () => {
  it('fsyncs the temporary file before the rename and the directory after it', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-fsync-'))
    const target = join(dir, 'report.docx')
    try {
      await atomicWriteFile(target, new TextEncoder().encode('new'))
      expect(readFileSync(target, 'utf8')).toBe('new')
      expect(synced).toHaveLength(process.platform === 'win32' ? 1 : 2)
      expect(synced[0]).toMatch(/\.report\.docx\.[0-9a-f]{12}\.tmp$/)
      if (process.platform !== 'win32') expect(synced[1]).toBe(dir)
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  it('still publishes the file when the filesystem refuses the fsync', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-fsync-'))
    const target = join(dir, 'report.docx')
    syncError.code = 'EINVAL'
    try {
      await atomicWriteFile(target, new TextEncoder().encode('new'))
      expect(readFileSync(target, 'utf8')).toBe('new')
      expect(synced.length).toBeGreaterThan(0)
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  it('fails the write on an unexpected fsync error and leaves no temp file behind', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-fsync-'))
    const target = join(dir, 'report.docx')
    syncError.code = 'EIO'
    try {
      await expect(atomicWriteFile(target, new TextEncoder().encode('new'))).rejects.toThrow()
      expect(readdirSync(dir)).toEqual([])
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })
})
