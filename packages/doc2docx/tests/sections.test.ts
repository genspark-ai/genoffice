import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '@genoffice/docx-engine'
import {
  LegacyDocError,
  convertLegacyDocToDocx,
  openLegacyDoc,
  readLegacyDocLayout,
} from '../src/index'
import {
  SYNTHETIC_TEXT_OFFSET,
  buildSyntheticDoc,
  encodeChpxFkp,
  encodePapxFkp,
  encodePlcfHdd,
  type SyntheticSection,
} from './helpers/synthetic-doc'
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
        ' w:header="708" w:footer="708" w:gutter="0"/>' +
        '<w:docGrid w:linePitch="360"/></w:sectPr>',
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
    // no even stories -> no evenAndOddHeaders; the Word 97-2003 compat block is always there
    const settings = (await zipText(docx, 'word/settings.xml'))!
    expect(settings).not.toContain('<w:evenAndOddHeaders/>')
    expect(settings).toContain('<w:balanceSingleByteDoubleByteWidth/>')
    expect(settings).toContain(
      'w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="11"',
    )
  })

  it('writes the document grid of a section that snaps to lines', async () => {
    const grid = concat(sprm(0x5032, [2, 0]), sprm(0x9031, [0x38, 0x01]))
    const layout = readLegacyDocLayout(syntheticSectionDoc({ sections: [{ sepxSprms: grid }] }))
    expect(layout.sections[0]!.docGrid).toEqual({ type: 'lines', linePitchTwips: 312 })
    const { docx } = await convertLegacyDocToDocx(
      syntheticSectionDoc({ sections: [{ sepxSprms: grid }] }),
    )
    const documentXml = (await zipText(docx, 'word/document.xml'))!
    expect(sectPrOf(documentXml)).toContain(
      '<w:docGrid w:type="lines" w:linePitch="312"/></w:sectPr>',
    )
    const plain = await convertLegacyDocToDocx(syntheticSectionDoc({ sections: [{}] }))
    expect(sectPrOf((await zipText(plain.docx, 'word/document.xml'))!)).not.toContain('docGrid')
  })

  it('assembles a character grid from sprms in ispmd order (char space before clm)', async () => {
    const grid = concat(
      sprm(0x7030, [0x40, 0x0a, 0x00, 0x00]),
      sprm(0x9031, [0x38, 0x01]),
      sprm(0x5032, [1, 0]),
    )
    const layout = readLegacyDocLayout(syntheticSectionDoc({ sections: [{ sepxSprms: grid }] }))
    expect(layout.sections[0]!.docGrid).toEqual({
      type: 'linesAndChars',
      linePitchTwips: 312,
      charSpace: 2624,
    })
    const { docx } = await convertLegacyDocToDocx(
      syntheticSectionDoc({ sections: [{ sepxSprms: grid }] }),
    )
    expect(sectPrOf((await zipText(docx, 'word/document.xml'))!)).toContain(
      '<w:docGrid w:type="linesAndChars" w:linePitch="312" w:charSpace="2624"/>',
    )
    const linesOnly = concat(
      sprm(0x7030, [0x40, 0x0a, 0, 0]),
      sprm(0x9031, [0x38, 0x01]),
      sprm(0x5032, [2, 0]),
    )
    expect(
      readLegacyDocLayout(syntheticSectionDoc({ sections: [{ sepxSprms: linesOnly }] }))
        .sections[0]!.docGrid,
    ).toEqual({ type: 'lines', linePitchTwips: 312 })
  })

  it('writes the resolved run and paragraph formatting of header stories', async () => {
    const body = 'Section body text.\r'
    const header = 'Bold header'
    const headerCp = body.length
    const bytes = buildSyntheticDoc({
      pieces: [{ text: body }],
      sections: [{}],
      headerStories: [undefined, header, undefined, undefined, undefined, undefined],
      chpxFkpPages: [
        encodeChpxFkp([
          {
            fcStart: SYNTHETIC_TEXT_OFFSET + headerCp * 2,
            fcEnd: SYNTHETIC_TEXT_OFFSET + (headerCp + 4) * 2,
            grpprl: sprm(0x0835, [1]),
          },
          {
            fcStart: SYNTHETIC_TEXT_OFFSET + (headerCp + 4) * 2,
            fcEnd: SYNTHETIC_TEXT_OFFSET + (headerCp + header.length + 1) * 2,
            grpprl: sprm(0x2a0c, [7]),
          },
        ]),
      ],
      papxFkpPages: [
        encodePapxFkp([
          {
            fcStart: SYNTHETIC_TEXT_OFFSET + headerCp * 2,
            fcEnd: SYNTHETIC_TEXT_OFFSET + (headerCp + header.length + 1) * 2,
            grpprl: new Uint8Array([...sprm(0x2461, [1]), ...sprm(0x6426, [8, 1, 6, 1])]),
          },
        ]),
      ],
    })
    const layout = readLegacyDocLayout(bytes)
    const story = layout.headerFooter[0]!.oddHeader!
    expect(story.paragraphs).toEqual([header])
    expect(story.formatted?.[0]!.runs.map((run) => run.text)).toEqual(['Bold', ' header'])
    const { docx } = await convertLegacyDocToDocx(bytes)
    const xml = (await zipText(docx, 'word/header1.xml'))!
    expect(xml).toContain('<w:b/>')
    expect(xml).toContain('<w:highlight w:val="yellow"/>')
    expect(xml).toContain('<w:jc w:val="center"/>')
    expect(xml).toContain('<w:bottom w:val="single" w:sz="8" w:space="1" w:color="FF0000"/>')
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

  it('ends every section but the last inside its last paragraph', async () => {
    const bytes = syntheticSectionDoc({
      text: 'First\u000cSecond\r',
      sections: [
        { endCp: 6, sepxSprms: concat(CUSTOM_SEPX, sprm(0x300a, [0x01])) },
        { sepxSprms: sprm(0xb01f, [0x40, 0x1f]) }, // 8000 twips wide, portrait
      ],
    })
    expect(readLegacyDocLayout(bytes).sections).toHaveLength(2)
    const { docx, warnings } = await convertLegacyDocToDocx(bytes)
    expect(warnings).toEqual([])
    const xml = (await zipText(docx, 'word/document.xml'))!
    expect(xml.match(/<w:sectPr>/g)).toHaveLength(2)
    const first = xml.slice(xml.indexOf('<w:p>'), xml.indexOf('</w:p>') + '</w:p>'.length)
    expect(first).toContain('First')
    expect(first).toMatch(
      /<w:pPr>.*<w:sectPr><w:pgSz w:w="8640" w:h="12960" w:orient="landscape"\/>/,
    )
    expect(first).toContain('<w:titlePg/>')
    expect(first).toMatch(/<\/w:sectPr><\/w:pPr>/)
    expect(xml.indexOf('Second')).toBeGreaterThan(xml.indexOf('</w:sectPr>'))
    const last = sectPrOf(xml)
    expect(last).toContain('<w:pgSz w:w="8000" w:h="16838"/>')
    expect(last).not.toContain('w:orient="landscape"')
    await expect(parseDocx(docx)).resolves.toBeDefined()
  })

  it('writes the SEPX break code as w:type', async () => {
    const bytes = syntheticSectionDoc({
      text: 'First\u000cSecond\u000cThird\r',
      sections: [
        { endCp: 6, sepxSprms: sprm(0x3009, [0x03]) },
        { endCp: 13, sepxSprms: sprm(0x3009, [0x00]) },
        { sepxSprms: sprm(0x3009, [0x02]) },
      ],
    })
    expect(readLegacyDocLayout(bytes).sections.map((section) => section.breakType)).toEqual([
      'evenPage',
      'continuous',
      'nextPage',
    ])
    const xml = (await zipText((await convertLegacyDocToDocx(bytes)).docx, 'word/document.xml'))!
    expect(xml.match(/<w:sectPr>(<w:type[^>]*>)?/g)).toEqual([
      '<w:sectPr><w:type w:val="evenPage"/>',
      '<w:sectPr><w:type w:val="continuous"/>',
      '<w:sectPr>',
    ])
  })

  it('writes per-section header parts and references them from the right sectPr', async () => {
    const { docx, warnings } = await convertLegacyDocToDocx(
      syntheticSectionDoc({
        text: 'First\u000cSecond\r',
        sections: [{ endCp: 6 }, {}],
        sectionHeaderStories: [
          [undefined, 'Header one', undefined, 'Footer one', undefined, undefined],
          [undefined, 'Header two', undefined, 'Footer two', undefined, undefined],
        ],
      }),
    )
    expect(warnings).toEqual([])
    expect(await zipText(docx, 'word/header1.xml')).toContain('Header one')
    expect(await zipText(docx, 'word/footer1.xml')).toContain('Footer one')
    expect(await zipText(docx, 'word/header2.xml')).toContain('Header two')
    expect(await zipText(docx, 'word/footer2.xml')).toContain('Footer two')
    const xml = (await zipText(docx, 'word/document.xml'))!
    const first = xml.slice(0, xml.indexOf('</w:sectPr>'))
    expect(first).toContain('<w:headerReference w:type="default" r:id="rId3"/>')
    expect(first).toContain('<w:footerReference w:type="default" r:id="rId4"/>')
    const last = sectPrOf(xml)
    expect(last).toContain('<w:headerReference w:type="default" r:id="rId5"/>')
    expect(last).toContain('<w:footerReference w:type="default" r:id="rId6"/>')
    const rels = (await zipText(docx, 'word/_rels/document.xml.rels'))!
    expect(rels).toContain(
      'Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"',
    )
    const contentTypes = (await zipText(docx, '[Content_Types].xml'))!
    expect(contentTypes).toContain('<Override PartName="/word/header2.xml"')
    expect(contentTypes).toContain('<Override PartName="/word/footer2.xml"')
    const parsed = await parseDocx(docx)
    expect(parsed.headerText).toBe('Header one')
  })

  it('omits the reference of a section that inherits its header, and writes an unlinked blank one', async () => {
    const { docx } = await convertLegacyDocToDocx(
      syntheticSectionDoc({
        text: 'First\u000cSecond\u000cThird\r',
        sections: [{ endCp: 6 }, { endCp: 13 }, {}],
        sectionHeaderStories: [
          [undefined, 'Header one', undefined, 'Footer one', undefined, undefined],
          [undefined, undefined, undefined, undefined, undefined, undefined],
          [undefined, '', undefined, undefined, undefined, undefined],
        ],
      }),
    )
    const xml = (await zipText(docx, 'word/document.xml'))!
    const sectPrs = xml.match(/<w:sectPr>.*?<\/w:sectPr>/g)!
    expect(sectPrs).toHaveLength(3)
    expect(sectPrs[0]).toContain('headerReference')
    expect(sectPrs[1]).not.toContain('Reference')
    expect(sectPrs[2]).toContain('<w:headerReference w:type="default" r:id="rId5"/>')
    expect(sectPrs[2]).not.toContain('footerReference')
    expect(await zipText(docx, 'word/header2.xml')).toMatch(/<w:hdr [^>]*><w:p><\/w:p><\/w:hdr>/)
    expect(await zipText(docx, 'word/footer2.xml')).toBeUndefined()
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
    expect(stories.oddHeader).toMatchObject({
      text: 'Line one\nLine two',
      paragraphs: ['Line one', 'Line two'],
    })
    expect(stories.oddHeader!.formatted!.map((paragraph) => paragraph.text)).toEqual([
      'Line one',
      'Line two',
    ])
    expect(stories.oddFooter).toMatchObject({ text: 'Tab\there', paragraphs: ['Tab\there'] })
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

  it('rewrites section marks to paragraph marks and keeps manual page breaks', async () => {
    const text = 'First\u000cSecond\u000cmore\r'
    const bytes = syntheticSectionDoc({
      text,
      sections: [{ endCp: 6, defaultSep: true }, { defaultSep: true }],
    })
    const doc = openLegacyDoc(bytes)
    expect(doc.rawText).toBe('First\rSecond\u000cmore\r')
    expect(doc.paragraphs).toEqual(['First', 'Second\fmore', ''])
    const { docx } = await convertLegacyDocToDocx(bytes)
    const xml = (await zipText(docx, 'word/document.xml'))!
    expect((xml.match(/<w:br w:type="page"\/>/g) ?? []).length).toBe(1)
    expect(xml.indexOf('First')).toBeLessThan(xml.indexOf('<w:br w:type="page"/>'))
  })
})
