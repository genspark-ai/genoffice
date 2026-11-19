import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const styles = (docRf: string, extra = '') =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${W}><w:docDefaults>` +
  `<w:rPrDefault><w:rPr>${docRf}<w:sz w:val="22"/></w:rPr></w:rPrDefault>` +
  '<w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  `${extra}</w:styles>`
const run = (rf: string, text: string) =>
  `<w:r><w:rPr><w:rFonts ${rf}/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`

// Word probe 2026-11-02: w:hAnsi covers the high-ANSI range only; a run or style
// without w:ascii inherits the ascii face from the chain instead of borrowing hAnsi
describe('rFonts ascii slot', () => {
  it('a lone w:hAnsi on a run leaves the Latin face to the chain', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          `<w:p>${run('w:hAnsi="Courier New" w:eastAsia="Courier New" w:cs="Courier New"', 'hansi only')}` +
          `${run('w:ascii="Courier New"', 'ascii only')}</w:p>`,
        stylesXml: styles('<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'),
      }),
    )
    const [a, b] = doc.blocks[0].runs!
    expect(a.fontAscii).toBeUndefined()
    expect(a.font).toBe('Courier New')
    expect(a.eastAsiaFont).toBe('Courier New')
    expect(a.effective?.fonts.ascii).toBe('Times New Roman')
    expect(b.fontAscii).toBe('Courier New')
  })

  it('a lone w:hAnsi on a style inherits too; docDefaults still borrow it (nothing above)', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: `<w:p><w:pPr><w:pStyle w:val="Hdr"/></w:pPr>${run('w:eastAsia="MS Mincho"', 'styled')}</w:p>`,
        stylesXml: styles(
          '<w:rFonts w:hAnsi="Garamond"/>',
          '<w:style w:type="paragraph" w:styleId="Hdr"><w:name w:val="Hdr"/><w:basedOn w:val="Normal"/>' +
            '<w:rPr><w:rFonts w:hAnsi="Courier New"/></w:rPr></w:style>',
        ),
      }),
    )
    expect(doc.styles.get('Hdr')?.display?.fontAscii).toBeUndefined()
    expect(doc.docDefaults?.asciiFont).toBe('Garamond')
    expect(doc.blocks[0].runs![0].effective?.fonts.ascii).toBe('Garamond')
  })
})
