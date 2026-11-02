import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '@genoffice/docx-engine'
import { LegacyDocError, convertLegacyDocToDocx, readLegacyDocLayout } from '../src/index'
import { buildSyntheticDoc, encodePlcfHdd, type SyntheticSection } from './helpers/synthetic-doc'
import { fixtureBytes } from './helpers/fixtures'

function sprm(opcode: number, operand: number[]): Uint8Array {
  return Uint8Array.from([opcode & 0xff, (opcode >> 8) & 0xff, ...operand])
}

function concat(...parts: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0))
  let offset = 0
  for (const part of parts) {
    out.set(part, offset)
    offset += part.length
  }
  return out
}

async function zipText(docx: Uint8Array, path: string): Promise<string | undefined> {
  const zip = await JSZip.loadAsync(docx)
  const file = zip.file(path)
  return file ? file.async('string') : undefined
}

function sectPrOf(documentXml: string): string {
  const start = documentXml.lastIndexOf('<w:sectPr')
  const end = documentXml.indexOf('</w:sectPr>', start)
  return documentXml.slice(start, end + '</w:sectPr>'.length)
}

function errorCode(run: () => unknown): string {
  try {
    run()
  } catch (error) {
    expect(error).toBeInstanceOf(LegacyDocError)
    return (error as LegacyDocError).code
  }
  throw new Error('expected a LegacyDocError')
}

/** a custom landscape/title-page section (8640x12960, 720/864/1008/1152 margins, 720 header/footer) */
const CUSTOM_SEPX = concat(
  sprm(0xb01f, [0xc0, 0x21]), // XaPage 8640
  sprm(0xb020, [0xa0, 0x32]), // YaPage 12960
  sprm(0xb021, [0xd0, 0x02]), // left 720
  sprm(0xb022, [0x60, 0x03]), // right 864
  sprm(0x9023, [0xf0, 0x03]), // top 1008
  sprm(0x9024, [0x80, 0x04]), // bottom 1152
  sprm(0xb017, [0xd0, 0x02]), // header distance 720
  sprm(0xb018, [0xd0, 0x02]), // footer distance 720
  sprm(0x301d, [0x01]), // landscape
)

function syntheticSectionDoc(options: {
  text?: string
  sections: SyntheticSection[]
  headerStories?: Array<string | undefined>
  sectionHeaderStories?: Array<Array<string | undefined>>
  lcbPlcfsed?: number
  lcbPlcfhdd?: number
  plcfHdd?: Uint8Array
}): Uint8Array {
  return buildSyntheticDoc({
    pieces: [{ text: options.text ?? 'Section body text.\r' }],
    sections: options.sections,
    ...(options.headerStories ? { headerStories: options.headerStories } : {}),
    ...(options.sectionHeaderStories ? { sectionHeaderStories: options.sectionHeaderStories } : {}),
    ...(options.lcbPlcfsed !== undefined ? { lcbPlcfsed: options.lcbPlcfsed } : {}),
    ...(options.lcbPlcfhdd !== undefined ? { lcbPlcfhdd: options.lcbPlcfhdd } : {}),
    ...(options.plcfHdd ? { plcfHdd: options.plcfHdd } : {}),
  })
}

describe('page-setup.doc: decoded section geometry', () => {
  const layout = readLegacyDocLayout(fixtureBytes('page-setup'))

  it('reads one section with the fixture page size and margins', () => {
    expect(layout.sections).toHaveLength(1)
    const [section] = layout.sections
    expect(section).toMatchObject({
      startCp: 0,
      endCp: 16,
      pageWidthTwips: 8640,
      // LibreOffice stores 229 mm (the python-docx 9in) as 12983 twips
      pageHeightTwips: 12983,
      orientation: 'portrait',
      titlePage: false,
      margins: {
        left: 720,
        right: 864,
        top: 1008,
        bottom: 1152,
        header: 708,
        footer: 708,
        gutter: 0,
      },
    })
    expect(section!.warnings).toEqual([])
    expect(layout.headerFooter).toEqual([])
    expect(layout.warnings).toEqual([])
  })

  it('writes the geometry into the generated sectPr and adds no header/footer parts', async () => {
    const { docx, warnings } = await convertLegacyDocToDocx(fixtureBytes('page-setup'))
    expect(warnings).toEqual([])
    const documentXml = (await zipText(docx, 'word/document.xml'))!
    expect(sectPrOf(documentXml)).toBe(
      '<w:sectPr><w:pgSz w:w="8640" w:h="12983"/>' +
        '<w:pgMar w:top="1008" w:right="864" w:bottom="1152" w:left="720"' +
        ' w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>',
    )
    expect(await zipText(docx, 'word/header1.xml')).toBeUndefined()
    expect(await zipText(docx, 'word/footer1.xml')).toBeUndefined()
    await expect(parseDocx(docx)).resolves.toBeDefined()
  })
})

describe('header-footer.doc: stories + writer parts', () => {
  const layout = readLegacyDocLayout(fixtureBytes('header-footer'))

  it('reads the odd header/footer stories and the unused first variants', () => {
    expect(layout.sections).toHaveLength(1)
    expect(layout.sections[0]).toMatchObject({
      pageWidthTwips: 12240,
      pageHeightTwips: 15840,
      titlePage: false,
      margins: { left: 1800, right: 1800, top: 1440, bottom: 1440, header: 720, footer: 720 },
    })
    const [stories] = layout.headerFooter
    expect(stories).toMatchObject({
      oddHeader: { text: 'Header line', paragraphs: ['Header line'] },
      oddFooter: { text: 'Footer line', paragraphs: ['Footer line'] },
    })
    expect(stories!.evenHeader).toBeUndefined()
    expect(stories!.evenFooter).toBeUndefined()
    // LibreOffice keeps first-page text in the file even though w:titlePg is off
    expect(stories!.firstHeader?.text).toBe('Header line')
    expect(stories!.firstFooter?.text).toBe('Footer line')
    expect(layout.warnings).toEqual([])
  })

  it('writes header1/footer1, references them from the sectPr, and links the package parts', async () => {
    const { docx, warnings } = await convertLegacyDocToDocx(fixtureBytes('header-footer'))
    expect(warnings).toEqual([])
    const header = (await zipText(docx, 'word/header1.xml'))!
    const footer = (await zipText(docx, 'word/footer1.xml'))!
    expect(header).toContain('<w:hdr')
    expect(header).toContain('Header line')
    expect(footer).toContain('<w:ftr')
    expect(footer).toContain('Footer line')

    const documentXml = (await zipText(docx, 'word/document.xml'))!
    const sectPr = sectPrOf(documentXml)
    expect(sectPr).toContain('<w:headerReference w:type="default" r:id="rId3"/>')
    expect(sectPr).toContain('<w:footerReference w:type="default" r:id="rId4"/>')
    expect(sectPr).not.toContain('w:titlePg')

    const rels = (await zipText(docx, 'word/_rels/document.xml.rels'))!
    expect(rels).toContain(
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
    )
    expect(rels).toContain(
      '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>',
    )
    const contentTypes = (await zipText(docx, '[Content_Types].xml'))!
    expect(contentTypes).toContain(
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>',
    )
    expect(contentTypes).toContain(
      '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>',
    )

    const parsed = await parseDocx(docx)
    expect(parsed.headerText).toBe('Header line')
    expect(parsed.footerText).toBe('Footer line')
  })
})

describe('synthetic sections', () => {
  it('falls back to the blank-template geometry for a default SEP (fcSepx 0xFFFFFFFF)', () => {
    const layout = readLegacyDocLayout(syntheticSectionDoc({ sections: [{ defaultSep: true }] }))
    expect(layout.sections[0]).toMatchObject({
      pageWidthTwips: 11906,
      pageHeightTwips: 16838,
      orientation: 'portrait',
      titlePage: false,
      margins: {
        left: 1440,
        right: 1440,
        top: 1440,
        bottom: 1440,
        header: 708,
        footer: 708,
        gutter: 0,
      },
    })
  })

  it('reads custom page size, margins, landscape and the title-page flag', () => {
    const layout = readLegacyDocLayout(
      syntheticSectionDoc({ sections: [{ sepxSprms: concat(CUSTOM_SEPX, sprm(0x300a, [0x01])) }] }),
    )
    expect(layout.sections[0]).toMatchObject({
      pageWidthTwips: 8640,
      pageHeightTwips: 12960,
      orientation: 'landscape',
      titlePage: true,
      margins: { left: 720, right: 864, top: 1008, bottom: 1152, header: 720, footer: 720 },
    })
  })

  it('writes landscape orientation and w:titlePg', async () => {
    const bytes = syntheticSectionDoc({
      sections: [{ sepxSprms: concat(CUSTOM_SEPX, sprm(0x300a, [0x01])) }],
    })
    const { docx } = await convertLegacyDocToDocx(bytes)
    const sectPr = sectPrOf((await zipText(docx, 'word/document.xml'))!)
    expect(sectPr).toContain('<w:pgSz w:w="8640" w:h="12960" w:orient="landscape"/>')
    expect(sectPr).toContain('<w:titlePg/>')
    expect(sectPr).toContain('<w:pgMar w:top="1008" w:right="864" w:bottom="1152" w:left="720"')
  })

  it('writes first header/footer parts only when the section has a title page', async () => {
    const titlePageStories = [
      undefined,
      'Odd header',
      undefined,
      'Odd footer',
      'First header',
      'First footer',
    ]
    const { docx } = await convertLegacyDocToDocx(
      syntheticSectionDoc({
        sections: [{ sepxSprms: sprm(0x300a, [0x01]) }],
        headerStories: titlePageStories,
      }),
    )
    const documentXml = (await zipText(docx, 'word/document.xml'))!
    const sectPr = sectPrOf(documentXml)
    expect(sectPr).toContain('<w:headerReference w:type="first" r:id="rId5"/>')
    expect(sectPr).toContain('<w:footerReference w:type="first" r:id="rId6"/>')
    expect(await zipText(docx, 'word/header2.xml')).toContain('First header')
    expect(await zipText(docx, 'word/footer2.xml')).toContain('First footer')
    // no even stories -> no settings.xml
    expect(await zipText(docx, 'word/settings.xml')).toBeUndefined()
  })

  it('writes even header/footer parts and settings.xml with w:evenAndOddHeaders', async () => {
    const { docx } = await convertLegacyDocToDocx(
      syntheticSectionDoc({
        sections: [{}],
        headerStories: ['Even header', undefined, 'Even footer', undefined, undefined, undefined],
      }),
    )
    const documentXml = (await zipText(docx, 'word/document.xml'))!
    const sectPr = sectPrOf(documentXml)
    expect(sectPr).toContain('<w:headerReference w:type="even" r:id="rId3"/>')
    expect(sectPr).toContain('<w:footerReference w:type="even" r:id="rId4"/>')
    expect(await zipText(docx, 'word/header1.xml')).toContain('Even header')
    expect(await zipText(docx, 'word/footer1.xml')).toContain('Even footer')
    const settings = (await zipText(docx, 'word/settings.xml'))!
    expect(settings).toContain('<w:evenAndOddHeaders/>')
    const rels = (await zipText(docx, 'word/_rels/document.xml.rels'))!
    expect(rels).toContain('Target="settings.xml"')
    const contentTypes = (await zipText(docx, '[Content_Types].xml'))!
    expect(contentTypes).toContain('<Override PartName="/word/settings.xml"')
  })

  it('flattens multiple sections to the last one with a warning', async () => {
    const bytes = syntheticSectionDoc({
      text: 'First section body.\r',
      sections: [
        { endCp: 5, sepxSprms: concat(CUSTOM_SEPX, sprm(0x300a, [0x01])) },
        { sepxSprms: sprm(0xb01f, [0x40, 0x1f]) }, // 8000 twips wide, portrait
      ],
    })
    const layout = readLegacyDocLayout(bytes)
    expect(layout.sections).toHaveLength(2)
    const { docx, warnings } = await convertLegacyDocToDocx(bytes)
    expect(warnings.some((warning) => warning.includes('sections were flattened to one'))).toBe(
      true,
    )
    const sectPr = sectPrOf((await zipText(docx, 'word/document.xml'))!)
    expect(sectPr).toContain('<w:pgSz w:w="8000" w:h="16838"/>')
    expect(sectPr).not.toContain('w:orient="landscape"')
  })

  it('keeps empty stories out of the package', async () => {
    const layout = readLegacyDocLayout(
      syntheticSectionDoc({
        sections: [{}],
        headerStories: [undefined, undefined, undefined, undefined, undefined, undefined],
      }),
    )
    expect(layout.headerFooter).toHaveLength(1)
    expect(layout.headerFooter[0]).toEqual({})
    const { docx } = await convertLegacyDocToDocx(
      syntheticSectionDoc({
        sections: [{}],
        headerStories: [undefined, undefined, undefined, undefined, undefined, undefined],
      }),
    )
    expect(await zipText(docx, 'word/header1.xml')).toBeUndefined()
    expect(await zipText(docx, 'word/footer1.xml')).toBeUndefined()
  })

  it('strips the guard paragraph mark and keeps multiple paragraphs and tabs', async () => {
    const bytes = syntheticSectionDoc({
      sections: [{}],
      headerStories: [
        undefined,
        'Line one\rLine two',
        undefined,
        'Tab\there',
        undefined,
        undefined,
      ],
    })
    const layout = readLegacyDocLayout(bytes)
    const stories = layout.headerFooter[0]!
    expect(stories.oddHeader).toEqual({
      text: 'Line one\nLine two',
      paragraphs: ['Line one', 'Line two'],
    })
    expect(stories.oddFooter).toEqual({ text: 'Tab\there', paragraphs: ['Tab\there'] })
    const { docx } = await convertLegacyDocToDocx(bytes)
    const header = (await zipText(docx, 'word/header1.xml'))!
    expect((header.match(/<w:p>/g) ?? []).length).toBe(2)
    const footer = (await zipText(docx, 'word/footer1.xml'))!
    expect(footer).toContain('<w:tab/>')
  })

  it('warns instead of failing when header text exists without a header table', async () => {
    const layout = readLegacyDocLayout(buildSyntheticDoc({ pieces: [{ text: 'x\r' }], ccpHdd: 6 }))
    expect(layout.headerFooter).toEqual([])
    expect(layout.warnings.some((warning) => warning.includes('no header table'))).toBe(true)
    const { warnings } = await convertLegacyDocToDocx(
      buildSyntheticDoc({ pieces: [{ text: 'x\r' }], ccpHdd: 6 }),
    )
    expect(warnings.some((warning) => warning.includes('headers are not converted'))).toBe(true)
  })
})

describe('section/header errors', () => {
  it('rejects a SEPX pointing outside the WordDocument stream', () => {
    const bytes = syntheticSectionDoc({ sections: [{ fcSepxOverride: 0xfffffff0 }] })
    expect(errorCode(() => readLegacyDocLayout(bytes))).toBe('bad_section')
  })

  it('rejects a section table whose length is not 16n+4', () => {
    const bytes = syntheticSectionDoc({ sections: [{}], lcbPlcfsed: 16 })
    expect(errorCode(() => readLegacyDocLayout(bytes))).toBe('bad_section')
  })

  it('rejects a header/footer table whose length is not a multiple of 4', () => {
    const bytes = syntheticSectionDoc({
      sections: [{}],
      headerStories: [undefined, 'Header', undefined, undefined, undefined, undefined],
      lcbPlcfhdd: 6,
    })
    expect(errorCode(() => readLegacyDocLayout(bytes))).toBe('bad_hdd')
  })

  it('rejects a header/footer table with too few stories for the sections', () => {
    const bytes = syntheticSectionDoc({
      sections: [{}],
      plcfHdd: encodePlcfHdd([0, 0, 0, 0, 0, 0, 0, 0]),
    })
    expect(errorCode(() => readLegacyDocLayout(bytes))).toBe('bad_hdd')
  })

  it('rejects unordered header/footer CPs', () => {
    const bytes = syntheticSectionDoc({
      sections: [{}],
      plcfHdd: encodePlcfHdd([0, 0, 0, 0, 0, 0, 0, 5, 2, 2, 2, 2, 2, 2]),
    })
    expect(errorCode(() => readLegacyDocLayout(bytes))).toBe('bad_hdd')
  })
})
