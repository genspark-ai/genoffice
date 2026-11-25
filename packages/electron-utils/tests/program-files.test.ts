import { describe, expect, it } from 'vitest'
import { isOpenableLinkSource, isProgramFile } from '../src/program-files'

describe('isOpenableLinkSource', () => {
  it('takes the documents and pictures a link can name', () => {
    for (const name of ['a.docx', 'A.PDF', '/tmp/b.xlsx', 'C:\\docs\\c.pptx', 'd.odt', 'e.jpeg'])
      expect(isOpenableLinkSource(name), name).toBe(true)
  })

  it('leaves out everything else, a program or not', () => {
    for (const name of [
      'run',
      'a.docm',
      'b.xlsm',
      'c.pptm',
      'd.py',
      'e.chm',
      'f.tool',
      'g.inetloc',
      'h.exe',
      'i.docx.',
      'j.docx ',
      'k.docx::$DATA',
      'Report.docx.app',
    ])
      expect(isOpenableLinkSource(name), name).toBe(false)
    expect(isProgramFile('h.exe')).toBe(true)
  })
})
