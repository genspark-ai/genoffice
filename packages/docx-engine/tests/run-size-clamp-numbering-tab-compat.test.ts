import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const settingsPart = (compat: string) => ({
  path: 'word/settings.xml',
  xml:
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:compat>${compat}</w:compat></w:settings>`,
  contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
})

describe('w:doNotUseIndentAsNumberingTabStop', () => {
  it('surfaces the compat flag', async () => {
    const on = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
        extraParts: [settingsPart('<w:doNotUseIndentAsNumberingTabStop/>')],
      }),
    )
    expect(on.indentNotNumberingTabStop).toBe(true)
    const off = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
        extraParts: [settingsPart('<w:doNotUseIndentAsNumberingTabStop w:val="0"/>')],
      }),
    )
    expect(off.indentNotNumberingTabStop).toBeUndefined()
  })
})

describe('w:sz clamp', () => {
  it('lays out w:sz=0 at one half-point like Word, not at the default size', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t>Q</w:t></w:r>' +
          '<w:r><w:rPr><w:sz w:val="0"/></w:rPr><w:t>proj</w:t></w:r>' +
          '<w:r><w:rPr><w:sz w:val="1"/></w:rPr><w:t>​</w:t></w:r></w:p>',
      }),
    )
    const runs = doc.blocks[0].runs!
    expect(runs.map((r) => r.sizeHalfPoints)).toEqual([21, 1, 1])
  })
})
