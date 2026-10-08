import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const BODY = '<w:p><w:r><w:t>Body</w:t></w:r></w:p>'
const stylesXml = (pPr: string) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}>` +
  `<w:docDefaults><w:rPrDefault/><w:pPrDefault><w:pPr>${pPr}</w:pPr></w:pPrDefault></w:docDefaults></w:styles>`

describe('pPrDefault w:widowControl', () => {
  it('an explicit off reaches docDefaults', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: BODY, stylesXml: stylesXml('<w:widowControl w:val="0"/>') }),
    )
    expect(doc.docDefaults?.widowControl).toBe(false)
  })

  it('on (the Word default) and absent stay unset', async () => {
    for (const pPr of ['<w:widowControl/>', '']) {
      const doc = await parseDocx(await buildDocx({ bodyXml: BODY, stylesXml: stylesXml(pPr) }))
      expect(doc.docDefaults?.widowControl).toBeUndefined()
    }
  })
})
