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
