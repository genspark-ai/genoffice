/**
 * Dual-slot run fonts (bug: editing the font flattened all four rFonts slots,
 * losing e.g. Times New Roman when a mixed CJK/Latin run was switched to KaiTi).
 */
import { describe, expect, it } from 'vitest'
import { generateParagraphXml, parseDocx, type GenerateContext } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const GEN_CTX: GenerateContext = {
  headingStyleIds: new Map([[1, 'Heading1']]),
  allocateHyperlinkRel: () => 'rId999',
}

const MIXED_RPR =
  '<w:rPr>' +
  '<w:rFonts w:ascii="Times New Roman" w:eastAsia="SimSun" w:hAnsi="Times New Roman" w:cs="Arial" w:hint="eastAsia"/>' +
  '</w:rPr>'
const MIXED_PARA = `<w:p><w:r>${MIXED_RPR}<w:t>合同 Contract</w:t></w:r></w:p>`

async function mixedRun() {
  const doc = await parseDocx(await buildDocx({ bodyXml: MIXED_PARA }))
  return doc.blocks[0].runs![0]
}

describe('rFonts dual-slot model', () => {
  it('parses eastAsia as primary font and ascii as the Latin slot', async () => {
    const run = await mixedRun()
    expect(run.font).toBe('SimSun')
    expect(run.fontAscii).toBe('Times New Roman')
  })

  it('untouched round-trip keeps the original rFonts bytes', async () => {
    const run = await mixedRun()
    const xml = generateParagraphXml({ type: 'paragraph', runs: [{ ...run }] }, GEN_CTX)
    expect(xml).toContain(
      '<w:rFonts w:ascii="Times New Roman" w:eastAsia="SimSun" w:hAnsi="Times New Roman" w:cs="Arial" w:hint="eastAsia"/>',
    )
  })

  it('switching the CJK font keeps the Latin and cs slots (bug 54)', async () => {
    const run = await mixedRun()
    const xml = generateParagraphXml(
      { type: 'paragraph', runs: [{ ...run, font: 'KaiTi' }] },
      GEN_CTX,
    )
    expect(xml).toContain('w:eastAsia="KaiTi"')
    expect(xml).toContain('w:ascii="Times New Roman"')
    expect(xml).toContain('w:hAnsi="Times New Roman"')
    expect(xml).toContain('w:cs="Arial"')
    expect(xml).toContain('w:hint="eastAsia"')
  })

  it('switching the Latin font keeps the eastAsia and cs slots', async () => {
    const run = await mixedRun()
    const xml = generateParagraphXml(
      { type: 'paragraph', runs: [{ ...run, fontAscii: 'Georgia' }] },
      GEN_CTX,
    )
    expect(xml).toContain('w:ascii="Georgia"')
    expect(xml).toContain('w:hAnsi="Georgia"')
    expect(xml).toContain('w:eastAsia="SimSun"')
    expect(xml).toContain('w:cs="Arial"')
  })

  it('Latin edit on an ascii-only run does not invent an eastAsia slot', async () => {
    const para =
      '<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>' +
      '<w:t>latin only</w:t></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    const run = doc.blocks[0].runs![0]
    expect(run.font).toBe('Arial') // parse-side derivation, not a real eastAsia slot
    const xml = generateParagraphXml(
      { type: 'paragraph', runs: [{ ...run, fontAscii: 'Georgia' }] },
      GEN_CTX,
    )
    expect(xml).toContain('w:ascii="Georgia"')
    expect(xml).toContain('w:hAnsi="Georgia"')
    expect(xml).not.toContain('eastAsia')
  })

  it('CJK edit on an ascii-only run adds eastAsia and keeps the Latin slot', async () => {
    const para =
      '<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>' +
      '<w:t>latin only</w:t></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    const run = doc.blocks[0].runs![0]
    const xml = generateParagraphXml(
      { type: 'paragraph', runs: [{ ...run, font: 'KaiTi' }] },
      GEN_CTX,
    )
    expect(xml).toContain('w:eastAsia="KaiTi"')
    expect(xml).toContain('w:ascii="Arial"')
    expect(xml).toContain('w:hAnsi="Arial"')
  })

  it('an explicit Latin font wins over a leftover theme attribute', async () => {
    const para =
      '<w:p><w:r><w:rPr>' +
      '<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsia="SimSun" w:hAnsiTheme="minorHAnsi"/>' +
      '</w:rPr><w:t>合同 Contract</w:t></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    const run = doc.blocks[0].runs![0]
    const xml = generateParagraphXml(
      { type: 'paragraph', runs: [{ ...run, fontAscii: 'Arial' }] },
      GEN_CTX,
    )
    expect(xml).toContain('w:ascii="Arial"')
    expect(xml).toContain('w:hAnsi="Arial"')
    expect(xml).not.toContain('asciiTheme')
    expect(xml).not.toContain('hAnsiTheme')
    expect(xml).toContain('w:eastAsia="SimSun"')
  })

  it('setting only the primary font on a run without rFonts fills every slot (legacy)', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>plain text</w:t></w:r></w:p>' }),
    )
    const run = doc.blocks[0].runs![0]
    const xml = generateParagraphXml(
      { type: 'paragraph', runs: [{ ...run, font: 'KaiTi' }] },
      GEN_CTX,
    )
    expect(xml).toContain(
      '<w:rFonts w:ascii="KaiTi" w:eastAsia="KaiTi" w:hAnsi="KaiTi" w:cs="KaiTi"/>',
    )
  })

  it('setting only the Latin font on a run without rFonts leaves eastAsia empty', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>plain text</w:t></w:r></w:p>' }),
    )
    const run = doc.blocks[0].runs![0]
    const xml = generateParagraphXml(
      { type: 'paragraph', runs: [{ ...run, fontAscii: 'Arial' }] },
      GEN_CTX,
    )
    expect(xml).toContain('<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>')
    expect(xml).not.toContain('eastAsia')
  })
})

describe('eastAsia theme slot', () => {
  const themePart = (eaMajor: string) => ({
    path: 'word/theme/theme1.xml',
    xml:
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">' +
      '<a:themeElements><a:fontScheme name="T">' +
      `<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface="${eaMajor}"/><a:cs typeface=""/></a:majorFont>` +
      '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>' +
      '</a:fontScheme></a:themeElements></a:theme>',
    contentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
  })
  const HEADING_RFONTS = '<w:rFonts w:eastAsiaTheme="majorEastAsia" w:eastAsia="Noto Sans CJK SC"/>'

  it('an empty theme slot resolves to DengXian, not the leftover literal', async () => {
    const para = `<w:p><w:r><w:rPr>${HEADING_RFONTS}</w:rPr><w:t>标题</w:t></w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: para, extraParts: [themePart('')] }))
    expect(doc.blocks[0].runs![0].font).toBe('DengXian')
  })

  it('a populated theme slot wins over the literal', async () => {
    const para = `<w:p><w:r><w:rPr>${HEADING_RFONTS}</w:rPr><w:t>标题</w:t></w:r></w:p>`
    const doc = await parseDocx(
      await buildDocx({ bodyXml: para, extraParts: [themePart('SimHei')] }),
    )
    expect(doc.blocks[0].runs![0].font).toBe('SimHei')
  })

  it('a theme ref without a theme part falls back to the literal', async () => {
    const para = `<w:p><w:r><w:rPr>${HEADING_RFONTS}</w:rPr><w:t>标题</w:t></w:r></w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    expect(doc.blocks[0].runs![0].font).toBe('Noto Sans CJK SC')
  })

  it('themeFontLang ja resolves empty slots to Yu Gothic (major) / Yu Mincho (minor)', async () => {
    const settingsPart = {
      path: 'word/settings.xml',
      xml:
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:themeFontLang w:val="en-US" w:eastAsia="ja-JP"/></w:settings>',
      contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
    }
    const para =
      `<w:p><w:r><w:rPr>${HEADING_RFONTS}</w:rPr><w:t>見出し</w:t></w:r>` +
      '<w:r><w:rPr><w:rFonts w:eastAsiaTheme="minorEastAsia"/></w:rPr><w:t>本文</w:t></w:r></w:p>'
    const doc = await parseDocx(
      await buildDocx({ bodyXml: para, extraParts: [themePart(''), settingsPart] }),
    )
    expect(doc.themeFonts?.eaLang).toBe('ja-JP')
    expect(doc.blocks[0].runs![0].font).toBe('Yu Gothic')
    expect(doc.blocks[0].runs![1].font).toBe('Yu Mincho')
  })

  it('themeFontLang ko resolves empty slots to Malgun Gothic', async () => {
    const settingsPart = {
      path: 'word/settings.xml',
      xml:
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:themeFontLang w:val="en-US" w:eastAsia="ko-KR"/></w:settings>',
      contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
    }
    const para =
      `<w:p><w:r><w:rPr>${HEADING_RFONTS}</w:rPr><w:t>제목</w:t></w:r>` +
      '<w:r><w:rPr><w:rFonts w:eastAsiaTheme="minorEastAsia"/></w:rPr><w:t>본문</w:t></w:r></w:p>'
    const doc = await parseDocx(
      await buildDocx({ bodyXml: para, extraParts: [themePart(''), settingsPart] }),
    )
    expect(doc.themeFonts?.eaLang).toBe('ko-KR')
    expect(doc.blocks[0].runs![0].font).toBe('Malgun Gothic')
    expect(doc.blocks[0].runs![1].font).toBe('Malgun Gothic')
  })

  it('an empty theme slot on a style rPr also resolves to DengXian', async () => {
    const styles =
      '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>' +
      `<w:rPr>${HEADING_RFONTS}<w:b/></w:rPr></w:style>`
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>标题</w:t></w:r></w:p>',
        extraStylesXml: styles,
        extraParts: [themePart('')],
      }),
    )
    expect(doc.styles.get('Heading1')!.display?.font).toBe('DengXian')
  })
})

describe('w:cs slot (complex-script font)', () => {
  it('parses the cs slot into csFont', async () => {
    const run = await mixedRun()
    expect(run.csFont).toBe('Arial')
  })

  it('resolves w:cstheme via the theme part', async () => {
    const themeXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T">' +
      '<a:themeElements><a:fontScheme name="T">' +
      '<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface="Times New Roman"/></a:majorFont>' +
      '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface="Arabic Typesetting"/></a:minorFont>' +
      '</a:fontScheme></a:themeElements></a:theme>'
    const para =
      '<w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cstheme="minorBidi"/></w:rPr>' +
      '<w:t>latin body</w:t></w:r></w:p>'
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: para,
        extraParts: [
          {
            path: 'word/theme/theme1.xml',
            xml: themeXml,
            contentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
          },
        ],
      }),
    )
    expect(doc.blocks[0].runs![0].csFont).toBe('Arabic Typesetting')
  })

  it('no cs slot means no csFont', async () => {
    const para =
      '<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>' +
      '<w:t>latin only</w:t></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    expect(doc.blocks[0].runs![0].csFont).toBeUndefined()
  })
})
