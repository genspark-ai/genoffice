import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const STYLES_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

// Normal turns widow control off (a Word 2003-era default in many CJK templates);
// Body inherits it through basedOn, Strict switches it back on explicitly
const STYLES =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${STYLES_NS}>` +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>' +
  '<w:pPr><w:widowControl w:val="0"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/><w:basedOn w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Strict"><w:name w:val="Strict"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:widowControl/></w:pPr></w:style>' +
  '</w:styles>'

describe('style-level w:widowControl', () => {
  it('is exposed on the style display and inherited through basedOn (explicit on overrides)', async () => {
    const bytes = await buildDocx({
      bodyXml:
        '<w:p><w:r><w:t>text</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:widowControl/></w:pPr><w:r><w:t>on again</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:widowControl w:val="0"/></w:pPr><w:r><w:t>off</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:widowControl w:val="off"/></w:pPr><w:r><w:t>off too</w:t></w:r></w:p>',
      stylesXml: STYLES,
    })
    const { styles, blocks } = await parseDocx(bytes)
    expect(styles.get('Normal')?.display?.widowControl).toBe(false)
    expect(styles.get('Normal')?.isDefault).toBe(true)
    expect(styles.get('Body')?.display?.widowControl).toBe(false)
    expect(styles.get('Strict')?.display?.widowControl).toBe(true)
    // the paragraph's own format stays untouched (saving must not add the property);
    // an explicit on is kept so it can override the style chain
    expect(blocks[0].format?.widowControl).toBeUndefined()
    expect(blocks[1].format?.widowControl).toBe(true)
    expect(blocks[2].format?.widowControl).toBe(false)
    expect(blocks[3].format?.widowControl).toBe(false)
  })
})
