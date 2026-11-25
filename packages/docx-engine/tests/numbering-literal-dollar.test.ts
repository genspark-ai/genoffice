import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '../src'
import { buildDocx } from './helpers/build-docx'
import { packageXmlProblems } from './helpers/strict-xml'

describe.each([0, 1])('literal numbering labels at level %i', (ilvl) => {
  it.each(['$$%1', '$$$$%1', '$&%1', '$`%1', "$'%1", '%1.'])(
    'preserves %s through save and reopen',
    async (lvlText) => {
      const bytes = await buildDocx({
        bodyXml: '<w:p><w:r><w:t>Unchanged body</w:t></w:r></w:p>',
        withNumbering: true,
      })
      const saved = await saveDocx(await parseDocx(bytes), [{ kind: 'original', docxIndex: 0 }], {
        numbering: {
          levelEdits: [
            { abstractNumId: '0', ilvl, level: { numFmt: 'decimal', lvlText, indentLeft: 720 } },
          ],
        },
      })
      expect(await packageXmlProblems(saved)).toEqual([])
      const zip = await JSZip.loadAsync(saved)
      const xml = await zip.file('word/numbering.xml')!.async('string')
      expect(xml).toContain(`w:lvlText w:val="${lvlText.replaceAll('&', '&amp;')}"`)
      const reopened = await parseDocx(saved)
      expect(reopened.numbering.get('1')?.levels[ilvl].lvlText).toBe(lvlText)
      const original = await JSZip.loadAsync(bytes)
      for (const [path, entry] of Object.entries(original.files)) {
        if (entry.dir || path === 'word/numbering.xml') continue
        expect(await zip.file(path)!.async('uint8array'), path).toEqual(
          await entry.async('uint8array'),
        )
      }
    },
  )
})
