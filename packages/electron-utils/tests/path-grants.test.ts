import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { createPathGrants, isBesideDocument } from '../src/path-grants'

describe('createPathGrants', () => {
  it('grants and verifies per view', () => {
    const grants = createPathGrants()
    grants.grant(1, ['/tmp/attachments/report.docx'])
    expect(grants.has(1, '/tmp/attachments/report.docx')).toBe(true)
    // another view does not inherit the grant
    expect(grants.has(2, '/tmp/attachments/report.docx')).toBe(false)
    // non-granted paths stay refused
    expect(grants.has(1, '/tmp/attachments/other.docx')).toBe(false)
    expect(grants.has(1, '/etc/passwd')).toBe(false)
  })

  it('resolves relative and non-normalized forms to the same grant', () => {
    const grants = createPathGrants()
    grants.grant(1, ['/tmp/attachments/./a/../report.docx'])
    expect(grants.has(1, '/tmp/attachments/report.docx')).toBe(true)
  })

  it('releases a view wholesale', () => {
    const grants = createPathGrants()
    grants.grant(7, ['/tmp/a.txt', '/tmp/b.txt'])
    grants.release(7)
    expect(grants.has(7, '/tmp/a.txt')).toBe(false)
    // re-granting after release works
    grants.grant(7, ['/tmp/c.txt'])
    expect(grants.has(7, '/tmp/c.txt')).toBe(true)
    expect(grants.has(7, '/tmp/a.txt')).toBe(false)
  })

  it('drops junk entries and soft-caps a runaway view', () => {
    const grants = createPathGrants()
    grants.grant(1, ['', 'x'.repeat(5000), '/tmp/ok.txt'] as unknown as string[])
    expect(grants.has(1, '/tmp/ok.txt')).toBe(true)
    for (let i = 0; i < 600; i++) grants.grant(1, [`/tmp/f${i}.txt`])
    // the cap holds and the earliest grants evicted FIFO
    let live = 0
    for (let i = 0; i < 600; i++) if (grants.has(1, `/tmp/f${i}.txt`)) live++
    expect(live).toBeLessThanOrEqual(512)
  })
})

describe('isBesideDocument', () => {
  const dir = '/srv/shared/notes'
  it('accepts siblings, the document itself and subdirectory notes', () => {
    expect(isBesideDocument(join(dir, 'a.md'), join(dir, 'b.md'))).toBe(true)
    expect(isBesideDocument(join(dir, 'a.md'), join(dir, 'sub', 'note.md'))).toBe(true)
    expect(isBesideDocument(join(dir, 'a.md'), join(dir, 'a.md'))).toBe(true)
  })
  it('refuses the directory itself, anything outside it, and a missing document', () => {
    expect(isBesideDocument(join(dir, 'a.md'), dir)).toBe(false)
    expect(isBesideDocument(join(dir, 'a.md'), join(dir, '..', 'elsewhere.md'))).toBe(false)
    expect(isBesideDocument(undefined, join(dir, 'b.md'))).toBe(false)
  })
})
