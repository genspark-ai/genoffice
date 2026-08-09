import { describe, expect, it } from 'vitest'
import {
  applySectionSettings,
  applySectionStartType,
  PAGE_MARK,
  parseDocx,
  readSections,
  readSectionSettings,
  saveDocx,
  type SaveBlock,
} from '../src/index'
import { buildDocx } from './helpers/build-docx'

const P = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`

/** Paragraph-level sectPr (section break): landscape A4 + narrow margins + optional continuous */
const sectBreakPara = (
  opts: { landscape?: boolean; continuous?: boolean; extra?: string } = {},
) => {
  const size = opts.landscape
    ? '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
    : '<w:pgSz w:w="11906" w:h="16838"/>'
  return (
    '<w:p><w:pPr><w:sectPr>' +
    (opts.continuous ? '<w:type w:val="continuous"/>' : '') +
    size +
    '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="708" w:footer="708" w:gutter="0"/>' +
    (opts.extra ?? '') +
    '</w:sectPr></w:pPr></w:p>'
  )
}

describe('readSections enumerates all sections', () => {
  it('single-section document: one section covers all blocks, trailing sectPr provides settings', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: P('a') + P('b') }))
    const sections = readSections(parsed)
    expect(sections.length).toBe(1)
    expect(sections[0].firstBlockIndex).toBe(0)
    expect(sections[0].settings.pageWidth).toBe(11906)
    expect(sections[0].startType).toBe('nextPage')
    expect(sections[0].settings).toEqual(readSectionSettings(parsed))
  })

  it('two sections: paragraph-level sectPr ends section 1 with correct block ranges', async () => {
    const bodyXml = P('第一节') + sectBreakPara({ landscape: true }) + P('第二节') + P('尾段')
    const parsed = await parseDocx(await buildDocx({ bodyXml }))
    const sections = readSections(parsed)
    expect(sections.length).toBe(2)
    // Section 1: block 0 (paragraph) + block 1 (section-break paragraph), landscape narrow margins
    expect(sections[0].firstBlockIndex).toBe(0)
    expect(sections[0].lastBlockIndex).toBe(1)
    expect(sections[0].settings.orientation).toBe('landscape')
    expect(sections[0].settings.pageWidth).toBe(16838)
    expect(sections[0].settings.marginTop).toBe(720)
    // Section 2: blocks 2..3 + the trailing hidden section-properties block, portrait A4 default margins
    expect(sections[1].firstBlockIndex).toBe(2)
    expect(sections[1].settings.orientation).toBe('portrait')
    expect(sections[1].settings.pageWidth).toBe(11906)
    expect(sections[1].settings.marginTop).toBe(1440)
    expect(sections[1].firstBlockIndex).toBe(sections[0].lastBlockIndex + 1)
  })

  it('parses startType/titlePg/pgNumType/header-footer references per section', async () => {
    const extra =
      '<w:headerReference w:type="default" r:id="rId7"/>' +
      '<w:footerReference w:type="first" r:id="rId8"/>' +
      '<w:titlePg/><w:pgNumType w:start="5"/>'
    const bodyXml = P('a') + sectBreakPara({ continuous: true, extra }) + P('b')
    const parsed = await parseDocx(await buildDocx({ bodyXml }))
    const sections = readSections(parsed)
    expect(sections[0].startType).toBe('continuous')
    expect(sections[0].titlePg).toBe(true)
    expect(sections[0].pageNumberStart).toBe(5)
    expect(sections[0].headerRefs.default).toBe('rId7')
    expect(sections[0].footerRefs.first).toBe('rId8')
    expect(sections[1].startType).toBe('nextPage')
    expect(sections[1].titlePg).toBe(false)
    expect(sections[1].pageNumberStart).toBeUndefined()
  })

  it('three sections with mixed portrait/landscape', async () => {
    const bodyXml =
      P('纵向一') + sectBreakPara() + P('横向二') + sectBreakPara({ landscape: true }) + P('纵向三')
    const parsed = await parseDocx(await buildDocx({ bodyXml }))
    const sections = readSections(parsed)
    expect(sections.map((s) => s.settings.orientation)).toEqual([
      'portrait',
      'landscape',
      'portrait',
    ])
    expect(sections.map((s) => s.firstBlockIndex)).toEqual([0, 2, 4])
  })

  it('hfParts: parses all header/footer parts by rId (PAGE field displayed as PAGE_MARK)', async () => {
    const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    const headerXml = `${XML}<w:hdr ${W}><w:p><w:r><w:t>第一节页眉</w:t></w:r></w:p></w:hdr>`
    const footerXml =
      `${XML}<w:ftr ${W}><w:p><w:r><w:t>第 </w:t></w:r>` +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t> 页</w:t></w:r></w:p></w:ftr>'
    const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    const bytes = await buildDocx({
      bodyXml:
        P('a') +
        sectBreakPara({
          extra:
            '<w:headerReference w:type="default" r:id="rId20"/><w:footerReference w:type="default" r:id="rId21"/>',
        }) +
        P('b'),
      extraRels:
        `<Relationship Id="rId20" Type="${REL}/header" Target="header1.xml"/>` +
        `<Relationship Id="rId21" Type="${REL}/footer" Target="footer1.xml"/>`,
      extraParts: [
        {
          path: 'word/header1.xml',
          xml: headerXml,
          contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
        },
        {
          path: 'word/footer1.xml',
          xml: footerXml,
          contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
        },
      ],
    })
    const parsed = await parseDocx(bytes)
    expect(parsed.hfParts?.rId20?.text).toBe('第一节页眉')
    expect(parsed.hfParts?.rId21?.hasPageNumber).toBe(true)
    expect(parsed.hfParts?.rId21?.text).toContain(PAGE_MARK)
    const sections = readSections(parsed)
    expect(sections[0].headerRefs.default).toBe('rId20')
    expect(sections[0].footerRefs.default).toBe('rId21')
    expect(sections[1].headerRefs.default).toBeUndefined()
  })

  it('applySectionStartType: inserts/replaces/removes w:type', () => {
    const base = '<w:sectPr><w:pgSz w:w="1" w:h="2"/></w:sectPr>'
    expect(applySectionStartType(base, 'continuous')).toBe(
      '<w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="1" w:h="2"/></w:sectPr>',
    )
    const cont = '<w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="1" w:h="2"/></w:sectPr>'
    expect(applySectionStartType(cont, 'evenPage')).toContain('<w:type w:val="evenPage"/>')
    expect(applySectionStartType(cont, 'nextPage')).toBe(base)
    // Without pgSz, insert at the start of sectPr
    expect(applySectionStartType('<w:sectPr></w:sectPr>', 'oddPage')).toBe(
      '<w:sectPr><w:type w:val="oddPage"/></w:sectPr>',
    )
  })

  it('section edit round-trip: non-final sectPr rewritten via kind:xml, final section via options.section', async () => {
    const bytes = await buildDocx({
      bodyXml: P('第一节') + sectBreakPara({ landscape: true }) + P('第二节'),
    })
    const parsed = await parseDocx(bytes)
    const sections = readSections(parsed)
    // Section 1: landscape → change to portrait narrow margins (simulating layout applied
    // to the section under the cursor)
    const edited = {
      ...sections[0].settings,
      orientation: 'portrait' as const,
      pageWidth: sections[0].settings.pageHeight,
      pageHeight: sections[0].settings.pageWidth,
      marginTop: 567,
    }
    const breakBlock = parsed.blocks.find((b) => b.docxIndex === sections[0].lastBlockIndex)!
    const newXml = breakBlock.originalXml!.replace(
      sections[0].sectPrXml,
      applySectionStartType(applySectionSettings(sections[0].sectPrXml, edited), 'continuous'),
    )
    const finalBlocks: SaveBlock[] = parsed.blocks
      .filter((b) => !b.hidden)
      .map((b) =>
        b.docxIndex === sections[0].lastBlockIndex
          ? { kind: 'xml', xml: newXml }
          : { kind: 'original', docxIndex: b.docxIndex! },
      )
    const saved = await saveDocx(parsed, finalBlocks, {
      section: { ...sections[1].settings, marginLeft: 999 },
    })
    const reparsed = readSections(await parseDocx(saved))
    expect(reparsed[0].settings.orientation).toBe('portrait')
    expect(reparsed[0].settings.marginTop).toBe(567)
    expect(reparsed[0].startType).toBe('continuous')
    expect(reparsed[1].settings.marginLeft).toBe(999)
  })

  it('insert section break: new break paragraph via kind:xml + trailing sectPr gets w:type', async () => {
    const bytes = await buildDocx({ bodyXml: P('a') + P('b') })
    const parsed = await parseDocx(bytes)
    const sections = readSections(parsed)
    expect(sections.length).toBe(1)
    const breakXml = `<w:p><w:pPr>${'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'}</w:pPr></w:p>`
    const visible = parsed.blocks.filter((b) => !b.hidden)
    const finalBlocks: SaveBlock[] = [
      { kind: 'original', docxIndex: visible[0].docxIndex! },
      { kind: 'xml', xml: breakXml },
      { kind: 'original', docxIndex: visible[1].docxIndex! },
    ]
    const saved = await saveDocx(parsed, finalBlocks, { sectionStartType: 'continuous' })
    const reparsed = readSections(await parseDocx(saved))
    expect(reparsed.length).toBe(2)
    expect(reparsed[0].lastBlockIndex).toBe(1)
    expect(reparsed[1].startType).toBe('continuous')
    expect(reparsed[1].firstBlockIndex).toBe(2)
  })

  it('unedited multi-section document saves byte-identical', async () => {
    const bytes = await buildDocx({
      bodyXml: P('第一节') + sectBreakPara({ landscape: true }) + P('第二节'),
    })
    const parsed = await parseDocx(bytes)
    readSections(parsed)
    const visible: SaveBlock[] = parsed.blocks
      .filter((b) => !b.hidden)
      .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))
    expect(await saveDocx(parsed, visible)).toBe(bytes)
  })
})

describe('sectionHf per-section headers/footers', () => {
  const visibleBlocks = (parsed: Awaited<ReturnType<typeof parseDocx>>): SaveBlock[] =>
    parsed.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))

  it('section without references: new part + reference injected into that sectPr, neighbors unaffected', async () => {
    const bodyXml = P('第一节') + sectBreakPara() + P('第二节')
    const parsed = await parseDocx(await buildDocx({ bodyXml }))
    const sections = readSections(parsed)
    expect(sections[0].headerRefs.default).toBeUndefined()

    const saved = await saveDocx(parsed, visibleBlocks(parsed), {
      sectionHf: [
        { lastBlockIndex: sections[0].lastBlockIndex, kind: 'header', hf: { text: '第一节页眉' } },
      ],
    })
    const reparsed = await parseDocx(saved)
    const secs = readSections(reparsed)
    const rId = secs[0].headerRefs.default
    expect(rId).toBeDefined()
    expect(reparsed.hfParts?.[rId!]?.text).toContain('第一节页眉')
    // The last section did not get a reference stuffed in
    expect(secs[1].headerRefs.default).toBeUndefined()
    // The reference is the first child of that section's sectPr
    const zip = await (await import('jszip')).default.loadAsync(saved)
    const docXml = await zip.file('word/document.xml')!.async('string')
    expect(docXml).toMatch(/<w:sectPr[^>]*><w:headerReference w:type="default"/)
  })

  it('section with an existing reference: rewrites the referenced part instead of creating one', async () => {
    const HDR =
      '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:p><w:r><w:t>旧页眉</w:t></w:r></w:p></w:hdr>'
    const bodyXml =
      P('第一节') +
      sectBreakPara({ extra: '<w:headerReference w:type="default" r:id="rId60"/>' }) +
      P('第二节')
    const bytes = await buildDocx({
      bodyXml,
      extraRels:
        '<Relationship Id="rId60" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
      extraParts: [
        {
          path: 'word/header1.xml',
          xml: HDR,
          contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
        },
      ],
    })
    const parsed = await parseDocx(bytes)
    const sections = readSections(parsed)
    expect(sections[0].headerRefs.default).toBe('rId60')

    const saved = await saveDocx(parsed, visibleBlocks(parsed), {
      sectionHf: [
        { lastBlockIndex: sections[0].lastBlockIndex, kind: 'header', hf: { text: '新页眉' } },
      ],
    })
    const zip = await (await import('jszip')).default.loadAsync(saved)
    const hdr = await zip.file('word/header1.xml')!.async('string')
    expect(hdr).toContain('新页眉')
    expect(zip.file('word/header2.xml')).toBeNull()
    // Reference count unchanged (no duplicate injection)
    const docXml = await zip.file('word/document.xml')!.async('string')
    expect(docXml.match(/<w:headerReference/g)).toHaveLength(1)
  })

  it('kind:xml section-break block (layout rewritten in the same pass) can also receive references', async () => {
    const bodyXml = P('第一节') + sectBreakPara({ landscape: true }) + P('第二节')
    const parsed = await parseDocx(await buildDocx({ bodyXml }))
    const sections = readSections(parsed)
    const breakBlock = parsed.blocks.find((b) => b.docxIndex === sections[0].lastBlockIndex)!
    const rewritten = applySectionSettings(sections[0].sectPrXml, {
      ...sections[0].settings,
      marginTop: 720,
    })
    const finalBlocks: SaveBlock[] = parsed.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) =>
        b.docxIndex === sections[0].lastBlockIndex
          ? {
              kind: 'xml' as const,
              xml: breakBlock.originalXml!.replace(sections[0].sectPrXml, rewritten),
              docxIndex: b.docxIndex!,
            }
          : { kind: 'original' as const, docxIndex: b.docxIndex! },
      )
    const saved = await saveDocx(parsed, finalBlocks, {
      sectionHf: [
        { lastBlockIndex: sections[0].lastBlockIndex, kind: 'footer', hf: { text: '第一节页脚' } },
      ],
    })
    const reparsed = await parseDocx(saved)
    const secs = readSections(reparsed)
    expect(secs[0].settings.marginTop).toBe(720)
    const rId = secs[0].footerRefs.default
    expect(rId).toBeDefined()
    expect(reparsed.hfParts?.[rId!]?.text).toContain('第一节页脚')
  })
})

describe('pgNumType page numbering', () => {
  it('applyPageNumType inserts/replaces/removes', async () => {
    const { applyPageNumType } = await import('../src/index')
    const base =
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="425"/></w:sectPr>'
    const withFmt = applyPageNumType(base, 'lowerRoman', 3)
    expect(withFmt).toContain('<w:pgNumType w:fmt="lowerRoman" w:start="3"/><w:cols')
    // Replacement does not duplicate
    const replaced = applyPageNumType(withFmt, 'upperLetter', undefined)
    expect(replaced.match(/<w:pgNumType/g)).toHaveLength(1)
    expect(replaced).toContain('w:fmt="upperLetter"')
    expect(replaced).not.toContain('w:start=')
    // All-unset removes the tag
    expect(applyPageNumType(withFmt, undefined, undefined)).not.toContain('pgNumType')
    // Without cols it lands right before </w:sectPr>
    const noCols = base.replace('<w:cols w:space="425"/>', '')
    expect(applyPageNumType(noCols, 'decimal', undefined)).toContain(
      '<w:pgNumType w:fmt="decimal"/></w:sectPr>',
    )
  })

  it('SaveOptions.pgNumType writes the final section and round-trips', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: P('正文') }))
    const blocks: SaveBlock[] = parsed.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))
    const saved = await saveDocx(parsed, blocks, { pgNumType: { fmt: 'upperRoman', start: 5 } })
    const secs = readSections(await parseDocx(saved))
    expect(secs[0].pageNumberFmt).toBe('upperRoman')
    expect(secs[0].pageNumberStart).toBe(5)
  })
})

describe('SaveOptions.numbering write-back', () => {
  const P2 = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
  const LI = (numId: number, text: string) =>
    `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`
  const NUMBERING =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:abstractNum w:abstractNumId="3"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>' +
    '<w:num w:numId="5"><w:abstractNumId w:val="3"/></w:num>' +
    '</w:numbering>'

  it('restartNums: appends a w:num pointing at an existing abstractNum + startOverride', async () => {
    const bytes = await buildDocx({
      bodyXml: LI(5, '一') + LI(5, '二'),
      numberingXml: NUMBERING,
    })
    const parsed = await parseDocx(bytes)
    const blocks: SaveBlock[] = parsed.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))
    const saved = await saveDocx(parsed, blocks, {
      numbering: { restartNums: [{ numId: '6', abstractNumId: '3', startOverrides: { 0: 1 } }] },
    })
    const reparsed = await parseDocx(saved)
    const def = reparsed.numbering.get('6')
    expect(def).toBeDefined()
    expect(def!.abstractNumId).toBe('3')
    expect(def!.startOverrides[0]).toBe(1)
    // Existing entries keep their original bytes
    const zip = await (await import('jszip')).default.loadAsync(saved)
    const numXml = await zip.file('word/numbering.xml')!.async('string')
    expect(numXml).toContain('<w:num w:numId="5"><w:abstractNumId w:val="3"/></w:num>')
  })

  it('newDefs: allocates a new abstractNum; creates part/rel/ContentType when numbering.xml is missing', async () => {
    const bytes = await buildDocx({ bodyXml: P2('正文') })
    const parsed = await parseDocx(bytes)
    const blocks: SaveBlock[] = parsed.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))
    const saved = await saveDocx(parsed, blocks, {
      numbering: { newDefs: [{ numId: '10', kind: 'ordered' }] },
    })
    const reparsed = await parseDocx(saved)
    const def = reparsed.numbering.get('10')
    expect(def).toBeDefined()
    expect(def!.levels[0].numFmt).toBe('decimal')
    // numIds 1/2 from the blank-template base are also present
    expect(reparsed.numbering.get('1')!.levels[0].numFmt).toBe('bullet')
    const zip = await (await import('jszip')).default.loadAsync(saved)
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
    expect(rels).toContain('Target="numbering.xml"')
    const ct = await zip.file('[Content_Types].xml')!.async('string')
    expect(ct).toContain('PartName="/word/numbering.xml"')
    // The abstract id does not clash with the template's 0/1
    const numXml = await zip.file('word/numbering.xml')!.async('string')
    expect(numXml).toContain('<w:num w:numId="10"><w:abstractNumId w:val="2"/></w:num>')
  })

  it('with an existing numbering.xml, newDefs abstracts are inserted before w:num entries', async () => {
    const bytes = await buildDocx({ bodyXml: LI(5, '一'), numberingXml: NUMBERING })
    const parsed = await parseDocx(bytes)
    const blocks: SaveBlock[] = parsed.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))
    const saved = await saveDocx(parsed, blocks, {
      numbering: { newDefs: [{ numId: '9', kind: 'bullet' }] },
    })
    const zip = await (await import('jszip')).default.loadAsync(saved)
    const numXml = await zip.file('word/numbering.xml')!.async('string')
    // The new abstract (id=4) comes before the first w:num; the new num goes at the end
    expect(numXml.indexOf('w:abstractNumId="4"')).toBeLessThan(numXml.indexOf('<w:num '))
    expect(numXml).toContain('<w:num w:numId="9"><w:abstractNumId w:val="4"/></w:num>')
    const reparsed = await parseDocx(saved)
    expect(reparsed.numbering.get('9')!.levels[0].numFmt).toBe('bullet')
  })
})
