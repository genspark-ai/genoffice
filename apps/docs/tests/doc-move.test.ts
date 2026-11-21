import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { renamedDocPath } from '../src/main/doc-move'

const FROM = join('/docs', 'Report.docx')

describe('File > Rename', () => {
  it('keeps the folder and the extension, typed or not', () => {
    expect(renamedDocPath(FROM, 'Plan')).toBe(join('/docs', 'Plan.docx'))
    expect(renamedDocPath(FROM, '  Plan.DOCX ')).toBe(join('/docs', 'Plan.docx'))
    expect(renamedDocPath(FROM, 'v1.2 notes')).toBe(join('/docs', 'v1.2 notes.docx'))
  })

  it('refuses an empty name, a path, a hidden file and another extension', () => {
    for (const name of ['', '   ', '.docx', '.hidden', 'a/b', 'a\\b', 'a:b', 'plan.pdf', 'x.doc'])
      expect(renamedDocPath(FROM, name), name).toBeNull()
  })
})
