import { mkdirSync, mkdtempSync, rmSync, symlinkSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { scanFiles } from '../src/main/file-index/scan'

let dir: string
beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'genoffice-scan-'))
})
afterEach(() => rmSync(dir, { recursive: true, force: true }))

describe('scanFiles budget', () => {
  it('stops descending past the depth limit', () => {
    let deep = dir
    for (let level = 0; level < 6; level += 1) {
      deep = join(deep, `d${level}`)
      mkdirSync(deep)
      writeFileSync(join(deep, `f${level}.md`), 'x')
    }
    expect(scanFiles(dir)).toMatchObject({ truncated: false })
    expect(scanFiles(dir).files).toHaveLength(6)
    const shallow = scanFiles(dir, { maxDepth: 3 })
    expect(shallow.truncated).toBe(true)
    expect(shallow.files).toHaveLength(3)
    expect(shallow.files.some((f) => f.path.endsWith('f3.md'))).toBe(false)
  })

  it('stops collecting at the file budget', () => {
    for (let i = 0; i < 10; i += 1) writeFileSync(join(dir, `n${i}.md`), 'x')
    const capped = scanFiles(dir, { maxFiles: 4 })
    expect(capped.files).toHaveLength(4)
    expect(capped.truncated).toBe(true)
    expect(scanFiles(dir, { maxFiles: 10 }).truncated).toBe(false)
  })

  it('does not follow symlinked directories', () => {
    mkdirSync(join(dir, 'real'))
    writeFileSync(join(dir, 'real', 'a.md'), 'x')
    symlinkSync(join(dir, 'real'), join(dir, 'real', 'loop'), 'dir')
    symlinkSync(join(dir, 'real'), join(dir, 'alias'), 'dir')
    expect(scanFiles(dir).files.map((f) => f.path)).toEqual([join(dir, 'real', 'a.md')])
  })
})

describe('scanFiles exclusions (#2844)', () => {
  it('prunes excluded subfolders and skips their files', () => {
    mkdirSync(join(dir, 'keep'))
    writeFileSync(join(dir, 'keep', 'a.md'), 'x')
    mkdirSync(join(dir, 'keep', 'secret'))
    writeFileSync(join(dir, 'keep', 'secret', 'b.md'), 'x')
    writeFileSync(join(dir, 'keep', 'c.md'), 'x')
    const res = scanFiles(dir, {}, [join(dir, 'keep', 'secret')])
    expect(res.files.some((f) => f.path.includes('b.md'))).toBe(false)
    expect(res.files.some((f) => f.path.endsWith('a.md'))).toBe(true)
    expect(res.files.some((f) => f.path.endsWith('c.md'))).toBe(true)
    // a bare-prefix sibling is not caught by the boundary rule
    mkdirSync(join(dir, 'secret-notes'))
    writeFileSync(join(dir, 'secret-notes', 'd.md'), 'x')
    expect(
      scanFiles(dir, {}, [join(dir, 'secret')]).files.some((f) => f.path.includes('d.md')),
    ).toBe(true)
  })
})
