import { mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
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

  it('rejects on a missing source without creating the target', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'atomic-copy-'))
    dirs.push(dir)
    await expect(atomicCopyFile(join(dir, 'nope'), join(dir, 'out'))).rejects.toThrow()
    expect(readdirSync(dir)).toEqual([])
  })
})
