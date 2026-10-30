import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { mergePPrFormat, parseDocx, saveDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const STYLES_OPEN =
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'

describe('linked styles: paragraph style never reads its Char twin', () => {
  it('keeps a heading in the paragraph chain font when only the Char style names one', async () => {
    const styles =
      '<w:style w:type="paragraph" w:styleId="Heading7"><w:name w:val="heading 7"/>' +
      '<w:link w:val="Heading7Char"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>' +
      '<w:style w:type="character" w:styleId="Heading7Char"><w:name w:val="Heading 7 Char"/>' +
      '<w:link w:val="Heading7"/>' +
      '<w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:sz w:val="40"/></w:rPr></w:style>'
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:pPr><w:pStyle w:val="Heading7"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>',
        extraStylesXml: styles,
      }),
    )
    const para = doc.styles.get('Heading7')!.display!
    expect(para.fontAscii).toBeUndefined()
    expect(para.sizeHalfPoints).toBe(32)
    // the Char shell still inherits what it lacks from the paragraph style
    expect(doc.styles.get('Heading7Char')!.display).toMatchObject({
      bold: true,
      fontAscii: 'Calibri Light',
      sizeHalfPoints: 40,
    })
    const run = doc.blocks[0].runs![0]
    expect(run.effective?.fonts.ascii).toBeUndefined()
    expect(run.effective?.sizeHalfPoints).toBe(32)
  })
})

describe('docDefaults w:rtl / w:bidi', () => {
  const stylesXml =
    STYLES_OPEN +
    '<w:docDefaults><w:rPrDefault><w:rPr>' +
    '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Arial"/>' +
    '<w:sz w:val="22"/><w:szCs w:val="28"/><w:rtl/>' +
    '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:bidi/></w:pPr></w:pPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
    '</w:styles>'

  it('runs without their own flag are complex-script runs: cs face and szCs', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:r><w:t>Abstract</w:t></w:r></w:p>' +
          '<w:p><w:r><w:rPr><w:rtl w:val="0"/></w:rPr><w:t>ltr</w:t></w:r></w:p>',
        stylesXml,
      }),
    )
    expect(doc.docDefaults).toMatchObject({ rtl: true, bidi: true })
    const rtlRun = doc.blocks[0].runs![0]
    expect(rtlRun.cs).toBe(true)
    expect(rtlRun.csFont).toBe('Arial')
    expect(rtlRun.effective?.sizeHalfPoints).toBe(28)
    // the flag stays inherited: the run's own rPr is not rewritten
    expect(rtlRun.rtl).toBeUndefined()
    const ltrRun = doc.blocks[1].runs![0]
    expect(ltrRun.cs).toBeUndefined()
    expect(ltrRun.csFont).toBeUndefined()
    expect(ltrRun.effective?.sizeHalfPoints).toBe(22)
  })

  it('paragraphs default to a bidi base direction with mirrored jc', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:r><w:t>plain</w:t></w:r></w:p>' +
          '<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>start</w:t></w:r></w:p>' +
          '<w:p><w:pPr><w:bidi w:val="0"/><w:jc w:val="left"/></w:pPr><w:r><w:t>ltr</w:t></w:r></w:p>',
        stylesXml,
      }),
    )
    const [plain, start, ltr] = doc.blocks
    expect(plain.format?.effective?.bidi).toBe(true)
    expect(plain.format?.bidi).toBeUndefined()
    expect(start.format?.effective?.bidi).toBe(true)
    expect(start.format?.align).toBe('right')
    expect(ltr.format?.effective?.bidi).toBe(false)
    expect(ltr.format?.align).toBe('left')
  })
})

describe('pPrDefault w:bidi: inherited direction never rewrites jc on save', () => {
  const stylesXml =
    STYLES_OPEN +
    '<w:docDefaults><w:pPrDefault><w:pPr><w:bidi/></w:pPr></w:pPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
    '</w:styles>'
  const bodyXml =
    '<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>one</w:t></w:r></w:p>' +
    '<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>two</w:t></w:r></w:p>'

  it('parses the mirrored visual align and marks the bidi as inherited only', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml, stylesXml }))
    expect(doc.blocks[0].format).toMatchObject({ align: 'right', bidiInherited: true })
    expect(doc.blocks[0].format?.bidi).toBeUndefined()
  })

  it('keeps jc left and adds no w:bidi when the pPr is rebuilt', async () => {
    const bytes = await buildDocx({ bodyXml, stylesXml })
    const doc = await parseDocx(bytes)
    const saved = await saveDocx(doc, [
      { kind: 'original', docxIndex: 0 },
      {
        kind: 'generated',
        docxIndex: 1,
        block: { type: 'paragraph', runs: [{ text: 'edited' }], format: doc.blocks[1].format },
      },
    ])
    const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    const paras = xml.match(/<w:p>[\s\S]*?<\/w:p>/g)!
    expect(paras).toHaveLength(2)
    expect(paras[0]).toContain('<w:jc w:val="left"/>')
    expect(paras[1]).toContain('edited')
    expect(paras[1]).toContain('<w:jc w:val="left"/>')
    expect(xml).not.toContain('<w:bidi')
  })

  it('a paragraph with no w:pPr mirrors a user-set align like an explicit bidi one', async () => {
    const bytes = await buildDocx({
      bodyXml:
        '<w:p><w:r><w:t>bare</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:t>explicit</w:t></w:r></w:p>',
      stylesXml,
    })
    const doc = await parseDocx(bytes)
    expect(doc.blocks[0].format).toMatchObject({ bidiInherited: true })
    expect(doc.blocks[0].format?.bidi).toBeUndefined()
    const saved = await saveDocx(
      doc,
      doc.blocks.slice(0, 2).map((block, docxIndex) => ({
        kind: 'generated' as const,
        docxIndex,
        block: {
          type: 'paragraph' as const,
          runs: [{ text: 'edited' }],
          format: { ...block.format, align: 'left' as const },
        },
      })),
    )
    const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    const paras = xml.match(/<w:p>[\s\S]*?<\/w:p>/g)!
    expect(paras).toHaveLength(2)
    expect(paras[0]).toContain('<w:jc w:val="right"/>')
    expect(paras[0]).not.toContain('<w:bidi')
    expect(paras[1]).toContain('<w:jc w:val="right"/>')
    expect(paras[1]).toContain('<w:bidi/>')
  })

  it('mergePPrFormat honours the inherited flag without emitting w:bidi', () => {
    const out = mergePPrFormat('<w:pPr><w:jc w:val="left"/></w:pPr>', {
      align: 'right',
      bidiInherited: true,
    })
    expect(out).toContain('<w:jc w:val="left"/>')
    expect(out).not.toContain('<w:bidi')
  })
})
