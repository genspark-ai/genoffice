import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import {
  applySectionSettings,
  applyTitlePg,
  notePropsFromXml,
  parseDocx,
  readSections,
  saveDocx,
  sectionSettingsFromXml,
} from '../src/index'
import { buildNotesXml } from '../src/notes'
import { buildDocx } from './helpers/build-docx'

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

const note = (id: number) =>
  `<w:footnote w:id="${id}"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>note ${id}</w:t></w:r></w:p></w:footnote>`

const FOOTNOTES_PART = {
  path: 'word/footnotes.xml',
  xml:
    XML_DECL +
    `<w:footnotes ${W_NS}>` +
    '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
    note(2) +
    note(3) +
    note(4) +
    '</w:footnotes>',
  contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
}

const FOOTNOTES_REL =
  '<Relationship Id="rId40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>'

const settingsPart = (inner: string) => ({
  path: 'word/settings.xml',
  xml: XML_DECL + `<w:settings ${W_NS}>${inner}</w:settings>`,
  contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
})

const PR = '<w:footnotePr><w:numStart w:val="5"/><w:numRestart w:val="eachSect"/></w:footnotePr>'

const ref = (id: number) =>
  `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`

describe('w:footnotePr / w:endnotePr', () => {
  it('reads pos, numFmt, numStart and numRestart, ignoring the settings.xml separator ids', () => {
    expect(
      notePropsFromXml(
        '<w:footnotePr><w:pos w:val="beneathText"/><w:numStart w:val="5"/>' +
          '<w:numRestart w:val="eachSect"/><w:footnote w:id="0"/><w:footnote w:id="1"/></w:footnotePr>' +
          '<w:endnotePr><w:numFmt w:val="decimal"/><w:endnote w:id="0"/></w:endnotePr>',
        'w:footnotePr',
      ),
    ).toEqual({ pos: 'beneathText', numStart: 5, numRestart: 'eachSect' })
    expect(
      notePropsFromXml('<w:endnotePr><w:numFmt w:val="decimal"/></w:endnotePr>', 'w:endnotePr'),
    ).toEqual({ numFmt: 'decimal' })
    expect(
      notePropsFromXml('<w:endnotePr><w:endnote w:id="0"/></w:endnotePr>', 'w:endnotePr'),
    ).toBeUndefined()
    expect(notePropsFromXml('<w:pgSz w:w="1" w:h="1"/>', 'w:footnotePr')).toBeUndefined()
  })

  it('effective document props: settings.xml overlaid by the final sectPr; the sectPr keeps its own', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
        sectPrExtra:
          '<w:endnotePr><w:numFmt w:val="decimal"/><w:numRestart w:val="eachSect"/></w:endnotePr>',
        extraParts: [
          settingsPart(
            '<w:footnotePr><w:pos w:val="beneathText"/></w:footnotePr>' +
              '<w:endnotePr><w:pos w:val="sectEnd"/><w:numFmt w:val="lowerLetter"/></w:endnotePr>',
          ),
        ],
      }),
    )
    expect(doc.footnoteProps).toEqual({ pos: 'beneathText' })
    expect(doc.endnoteProps).toEqual({ pos: 'sectEnd', numFmt: 'decimal', numRestart: 'eachSect' })
    const [section] = readSections(doc)
    expect(section.settings.endnotePr).toEqual({ numFmt: 'decimal', numRestart: 'eachSect' })
    expect(section.settings.footnotePr).toBeUndefined()
  })

  it('numbers references in body order from numStart, restarting where a section says so', async () => {
    const doc = await parseDocx(
      await buildDocx({
        // file order of the notes part is 2, 3, 4; the body references 3 first
        bodyXml:
          `<w:p><w:r><w:t>a</w:t></w:r>${ref(3)}</w:p>` +
          // the section-break paragraph's own reference still belongs to section 1
          `<w:p><w:pPr><w:sectPr>${PR}<w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr><w:r><w:t>b</w:t></w:r>${ref(2)}</w:p>` +
          `<w:p><w:r><w:t>c</w:t></w:r>${ref(4)}</w:p>`,
        sectPrExtra: PR,
        extraRels: FOOTNOTES_REL,
        extraParts: [FOOTNOTES_PART],
      }),
    )
    expect(doc.noteNumbers).toEqual({ 'footnote:3': 5, 'footnote:2': 6, 'footnote:4': 5 })
    const marks = doc.blocks
      .flatMap((b) => b.runs ?? [])
      .filter((r) => r.noteRef)
      .map((r) => r.text)
    expect(marks).toEqual(['5', '6', '5'])
  })

  it('each section numbers by settings.xml overlaid by its own sectPr', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          `<w:p><w:r><w:t>a</w:t></w:r>${ref(2)}</w:p>` +
          // the first section states nothing: the document's lowerLetter, counted on
          `<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr><w:r><w:t>b</w:t></w:r>${ref(3)}</w:p>` +
          `<w:p><w:r><w:t>c</w:t></w:r>${ref(4)}</w:p>`,
        sectPrExtra:
          '<w:footnotePr><w:numFmt w:val="upperRoman"/><w:numStart w:val="4"/><w:numRestart w:val="eachSect"/></w:footnotePr>',
        extraRels: FOOTNOTES_REL,
        extraParts: [
          FOOTNOTES_PART,
          settingsPart('<w:footnotePr><w:numFmt w:val="lowerLetter"/></w:footnotePr>'),
        ],
      }),
    )
    expect(doc.settingsNoteProps).toEqual({ footnote: { numFmt: 'lowerLetter' } })
    expect(doc.noteNumbers).toEqual({ 'footnote:2': 1, 'footnote:3': 2, 'footnote:4': 4 })
    expect(doc.noteNumFmts).toEqual({
      'footnote:2': 'lowerLetter',
      'footnote:3': 'lowerLetter',
      'footnote:4': 'upperRoman',
    })
    const marks = doc.blocks
      .flatMap((b) => b.runs ?? [])
      .filter((r) => r.noteRef)
      .map((r) => r.noteRef!.numFmt)
    expect(marks).toEqual(['lowerLetter', 'lowerLetter', 'upperRoman'])
  })

  it('a section that counts on shows its own start number plus the running count', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          `<w:p><w:pPr><w:sectPr w:rsidR="00A1"/></w:pPr><w:r><w:t>a</w:t></w:r>${ref(2)}${ref(3)}</w:p>` +
          `<w:p><w:r><w:t>c</w:t></w:r>${ref(4)}</w:p>`,
        sectPrExtra: '<w:footnotePr><w:numStart w:val="10"/></w:footnotePr>',
        extraRels: FOOTNOTES_REL,
        extraParts: [FOOTNOTES_PART],
      }),
    )
    expect(doc.noteNumbers).toEqual({ 'footnote:2': 1, 'footnote:3': 2, 'footnote:4': 12 })
    expect(doc.noteNumFmts).toBeUndefined()
  })

  it('a section writes its own properties after the header references and leaves equal ones alone', () => {
    const sectPr =
      '<w:sectPr><w:headerReference w:type="default" r:id="rId9"/>' +
      '<w:footnotePr><w:numFmt w:val="upperRoman"/></w:footnotePr>' +
      '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
      '<w:sectPrChange w:id="9" w:author="a"><w:sectPr><w:endnotePr><w:numFmt w:val="decimal"/></w:endnotePr></w:sectPr></w:sectPrChange>' +
      '</w:sectPr>'
    const settings = sectionSettingsFromXml(sectPr)
    // the tracked previous sectPr's properties are not the section's
    expect(settings.footnotePr).toEqual({ numFmt: 'upperRoman' })
    expect(settings.endnotePr).toBeUndefined()
    expect(applySectionSettings(sectPr, settings)).toBe(sectPr)
    const next = applySectionSettings(sectPr, {
      ...settings,
      footnotePr: { numFmt: 'upperRoman', numStart: 3, numRestart: 'eachSect' },
      endnotePr: { numFmt: 'upperLetter' },
    })
    expect(next).toContain(
      '<w:headerReference w:type="default" r:id="rId9"/>' +
        '<w:footnotePr><w:numFmt w:val="upperRoman"/><w:numStart w:val="3"/><w:numRestart w:val="eachSect"/></w:footnotePr>' +
        '<w:endnotePr><w:numFmt w:val="upperLetter"/></w:endnotePr><w:pgSz ',
    )
    expect(next).toContain(
      '<w:sectPrChange w:id="9" w:author="a"><w:sectPr><w:endnotePr><w:numFmt w:val="decimal"/></w:endnotePr></w:sectPr></w:sectPrChange></w:sectPr>',
    )
    expect(sectionSettingsFromXml(next).footnotePr).toEqual({
      numFmt: 'upperRoman',
      numStart: 3,
      numRestart: 'eachSect',
    })
  })

  it('a second reference to a note reuses its number; textbox paragraphs do not end the section-break paragraph', async () => {
    const txbx =
      '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline><a:graphic><a:graphicData>' +
      '<wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:t>box</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp>' +
      '</a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice></mc:AlternateContent></w:r>'
    const doc = await parseDocx(
      await buildDocx({
        bodyXml:
          `<w:p><w:r><w:t>a</w:t></w:r>${ref(2)}${ref(2)}${ref(3)}</w:p>` +
          // the textbox paragraph closes before the break paragraph's own reference
          `<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr>${txbx}${ref(4)}</w:p>` +
          `<w:p><w:r><w:t>c</w:t></w:r>${ref(3)}</w:p>`,
        sectPrExtra: '<w:footnotePr><w:numRestart w:val="eachSect"/></w:footnotePr>',
        docRootExtraAttrs:
          'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
          'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
        extraRels: FOOTNOTES_REL,
        extraParts: [FOOTNOTES_PART],
      }),
    )
    expect(doc.noteNumbers).toEqual({ 'footnote:2': 1, 'footnote:3': 2, 'footnote:4': 3 })
  })

  it('a tracked sectPrChange does not stand in for the live final sectPr', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
        sectPrExtra:
          '<w:footnotePr><w:numStart w:val="3"/></w:footnotePr>' +
          '<w:sectPrChange w:id="9" w:author="a" w:date="2024-01-01T00:00:00Z">' +
          '<w:sectPr><w:footnotePr><w:numStart w:val="7"/><w:pos w:val="beneathText"/></w:footnotePr></w:sectPr>' +
          '</w:sectPrChange>',
      }),
    )
    expect(doc.footnoteProps).toEqual({ numStart: 3 })
  })

  it('without note props the numbering is the plain body order', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: `<w:p><w:r><w:t>a</w:t></w:r>${ref(4)}${ref(2)}</w:p>`,
        extraRels: FOOTNOTES_REL,
        extraParts: [FOOTNOTES_PART],
      }),
    )
    expect(doc.footnoteProps).toBeUndefined()
    // the unreferenced note keeps its part-order number
    expect(doc.noteNumbers).toEqual({ 'footnote:4': 1, 'footnote:2': 2, 'footnote:3': 2 })
  })
})

const partOf = async (bytes: Uint8Array, path: string) =>
  (await (await JSZip.loadAsync(bytes)).file(path)?.async('string')) ?? ''

describe('w15:footnoteColumns', () => {
  const PAGE =
    '<w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
    '<w:cols w:num="2" w:space="720"/>'
  // where Word writes it: the last child of the sectPr, the value in the w namespace
  const stated = `<w:sectPr>${PAGE}<w:docGrid w:linePitch="360"/><w15:footnoteColumns w:val="1"/></w:sectPr>`

  it('reads the count a section states; none means as many as the section has', () => {
    expect(sectionSettingsFromXml(stated).footnoteColumns).toBe(1)
    expect(sectionSettingsFromXml(`<w:sectPr>${PAGE}</w:sectPr>`).footnoteColumns).toBeUndefined()
    expect(
      sectionSettingsFromXml(
        `<w:sectPr>${PAGE}<w15:footnoteColumns w:val="3"></w15:footnoteColumns></w:sectPr>`,
      ).footnoteColumns,
    ).toBe(3)
  })

  it('does not take a count Word does not offer, nor the one of a tracked previous sectPr', () => {
    for (const val of ['0', '5', 'x'])
      expect(
        sectionSettingsFromXml(`<w:sectPr>${PAGE}<w15:footnoteColumns w:val="${val}"/></w:sectPr>`)
          .footnoteColumns,
      ).toBeUndefined()
    expect(
      sectionSettingsFromXml(
        `<w:sectPr>${PAGE}<w:sectPrChange w:id="1" w:author="a">${stated}</w:sectPrChange></w:sectPr>`,
      ).footnoteColumns,
    ).toBeUndefined()
  })

  it('writes the count last, replaces it, and removes it to match the section', () => {
    const settings = sectionSettingsFromXml(stated)
    expect(applySectionSettings(stated, settings)).toBe(stated)
    const three = applySectionSettings(stated, { ...settings, footnoteColumns: 3 })
    expect(three).toBe(stated.replace('w:val="1"', 'w:val="3"'))
    const { footnoteColumns: _stated, ...matching } = settings
    expect(applySectionSettings(stated, matching)).toBe(
      `<w:sectPr>${PAGE}<w:docGrid w:linePitch="360"/></w:sectPr>`,
    )
    const tracked = `<w:sectPr>${PAGE}<w:sectPrChange w:id="1" w:author="a">${stated}</w:sectPrChange></w:sectPr>`
    expect(
      applySectionSettings(tracked, { ...sectionSettingsFromXml(tracked), footnoteColumns: 2 }),
    ).toBe(
      `<w:sectPr>${PAGE}<w15:footnoteColumns w:val="2"/><w:sectPrChange w:id="1" w:author="a">${stated}</w:sectPrChange></w:sectPr>`,
    )
  })

  it('a save declares the namespace and lists it as ignorable; the file reads back', async () => {
    const source = await buildDocx({ bodyXml: '<w:p><w:r><w:t>a</w:t></w:r></w:p>' })
    const parsed = await parseDocx(source)
    const section = { ...readSections(parsed)[0].settings, footnoteColumns: 2 }
    const saved = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }], { section })
    const documentXml = await partOf(saved, 'word/document.xml')
    expect(documentXml).toContain('<w15:footnoteColumns w:val="2"/></w:sectPr>')
    const root = /<w:document\b[^>]*>/.exec(documentXml)![0]
    expect(root).toContain('xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"')
    expect(root).toContain('xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"')
    expect(/mc:Ignorable="([^"]*)"/.exec(root)![1].split(' ')).toContain('w15')
    const reparsed = await parseDocx(saved)
    expect(readSections(reparsed)[0].settings.footnoteColumns).toBe(2)
    // saved again unchanged, the root keeps its bytes
    const again = await saveDocx(reparsed, [{ kind: 'original', docxIndex: 0 }], {
      section: readSections(reparsed)[0].settings,
    })
    expect(/<w:document\b[^>]*>/.exec(await partOf(again, 'word/document.xml'))![0]).toBe(root)
  })
  it('stays the last child when another property is added after it', () => {
    const bare = `<w:sectPr>${PAGE}<w15:footnoteColumns w:val="2"/></w:sectPr>`
    const settings = sectionSettingsFromXml(bare)
    const tail = (xml: string) => xml.slice(xml.indexOf('<w:cols'))
    expect(
      tail(applySectionSettings(bare, { ...settings, vAlign: 'center', formProt: false })),
    ).toBe(
      '<w:cols w:num="2" w:space="720"/><w:formProt w:val="0"/><w:vAlign w:val="center"/>' +
        '<w15:footnoteColumns w:val="2"/></w:sectPr>',
    )
    expect(tail(applyTitlePg(bare, true))).toBe(
      '<w:cols w:num="2" w:space="720"/><w:titlePg/><w15:footnoteColumns w:val="2"/></w:sectPr>',
    )
    expect(tail(applySectionSettings(bare, { ...settings, bidi: true }))).toBe(
      '<w:cols w:num="2" w:space="720"/><w:bidi/><w15:footnoteColumns w:val="2"/></w:sectPr>',
    )
  })
})

describe('saving note properties for the whole document', () => {
  it('sets them in settings.xml and in every section, keeps the rest, and reads them back', async () => {
    const source = await buildDocx({
      bodyXml:
        '<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rId9"/>' +
        '<w:footnotePr><w:numFmt w:val="upperRoman"/><w:numRestart w:val="eachPage"/></w:footnotePr>' +
        '<w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr><w:r><w:t>a</w:t></w:r></w:p>' +
        '<w:p><w:r><w:t>b</w:t></w:r></w:p>',
      extraParts: [
        settingsPart(
          '<w:zoom w:percent="100"/><w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>' +
            '<w:compat/>',
        ),
      ],
    })
    const parsed = await parseDocx(source)
    const saved = await saveDocx(
      parsed,
      [
        { kind: 'original', docxIndex: 0 },
        { kind: 'original', docxIndex: 1 },
      ],
      {
        noteProps: {
          footnote: { numFmt: 'lowerLetter', numStart: 3 },
          endnote: { numFmt: 'decimal' },
        },
      },
    )
    const settings = await partOf(saved, 'word/settings.xml')
    expect(settings).toContain(
      '<w:footnotePr><w:numFmt w:val="lowerLetter"/><w:numStart w:val="3"/><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>' +
        '<w:endnotePr><w:numFmt w:val="decimal"/></w:endnotePr><w:compat/>',
    )
    const documentXml = await partOf(saved, 'word/document.xml')
    // the section's own restart rule stays; footnotePr follows the header reference
    expect(documentXml).toContain(
      '<w:sectPr><w:headerReference w:type="default" r:id="rId9"/>' +
        '<w:footnotePr><w:numFmt w:val="lowerLetter"/><w:numStart w:val="3"/><w:numRestart w:val="eachPage"/></w:footnotePr>' +
        '<w:endnotePr><w:numFmt w:val="decimal"/></w:endnotePr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>',
    )
    expect(documentXml.match(/<w:footnotePr>/g)).toHaveLength(2)
    const reparsed = await parseDocx(saved)
    expect(reparsed.footnoteProps).toEqual({ numFmt: 'lowerLetter', numStart: 3 })
    expect(reparsed.endnoteProps).toEqual({ numFmt: 'decimal' })
  })

  it('settingsNoteProps reaches settings.xml alone', async () => {
    const source = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>a</w:t></w:r></w:p>',
      extraParts: [settingsPart('<w:zoom w:percent="100"/>')],
    })
    const saved = await saveDocx(await parseDocx(source), [{ kind: 'original', docxIndex: 0 }], {
      settingsNoteProps: { footnote: { numFmt: 'lowerRoman' }, endnote: {} },
    })
    expect(await partOf(saved, 'word/settings.xml')).toContain(
      '<w:footnotePr><w:numFmt w:val="lowerRoman"/></w:footnotePr>',
    )
    expect(await partOf(saved, 'word/settings.xml')).not.toContain('w:endnotePr')
    expect(await partOf(saved, 'word/document.xml')).not.toContain('w:footnotePr')
  })

  it('a kind the save sets nothing for gains no element', async () => {
    const source = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>a</w:t></w:r></w:p>',
      extraParts: [settingsPart('<w:zoom w:percent="100"/>')],
    })
    const saved = await saveDocx(await parseDocx(source), [{ kind: 'original', docxIndex: 0 }], {
      noteProps: { footnote: { numStart: 2 }, endnote: {} },
    })
    for (const path of ['word/settings.xml', 'word/document.xml']) {
      const xml = await partOf(saved, path)
      expect(xml).toContain('<w:footnotePr><w:numStart w:val="2"/></w:footnotePr>')
      expect(xml).not.toContain('w:endnotePr')
    }
  })
})

describe('custom note marks', () => {
  it('a w:customMarkFollows reference carries its mark, takes no number, and saves as it came', async () => {
    const custom =
      '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>' +
      '<w:footnoteReference w:customMarkFollows="1" w:id="3"/><w:t>\u2020</w:t></w:r>'
    const source = await buildDocx({
      bodyXml: `<w:p><w:r><w:t>a</w:t></w:r>${custom}${ref(2)}</w:p>`,
      extraRels: FOOTNOTES_REL,
      extraParts: [FOOTNOTES_PART],
    })
    const parsed = await parseDocx(source)
    const runs = parsed.blocks[0].runs ?? []
    expect(runs.map((r) => r.text)).toEqual(['a', '\u2020', '1'])
    expect(runs[1].noteRef).toEqual({ kind: 'footnote', id: '3', customMark: '\u2020' })
    const saved = await saveDocx(parsed, [
      { kind: 'generated', block: { type: 'paragraph', runs: [{ text: 'edited ' }, ...runs] } },
    ])
    expect(await partOf(saved, 'word/document.xml')).toContain(
      '<w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>' +
        '<w:footnoteReference w:customMarkFollows="1" w:id="3"/><w:t xml:space="preserve">\u2020</w:t></w:r>',
    )
  })

  it('a new note with a custom mark opens with the mark run instead of the self-reference', () => {
    const xml = buildNotesXml(
      'footnote',
      [{ id: '1', text: '* starred', noRefMark: true, customMark: '*' }],
      null,
    )
    expect(xml).toContain(
      '<w:footnote w:id="1"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t xml:space="preserve">*</w:t></w:r>' +
        '<w:r><w:t xml:space="preserve"> starred</w:t></w:r></w:p></w:footnote>',
    )
    expect(xml).not.toContain('<w:footnoteRef/>')
  })
})

describe('converted notes', () => {
  it('an endnote made from a footnote keeps the entry bytes under the other kind', () => {
    const footnotes =
      `<w:footnotes ${W_NS}><w:footnote w:id="2"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>` +
      '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>' +
      '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve"> bold note</w:t></w:r></w:p></w:footnote></w:footnotes>'
    const convert = (styleIds: string[]) =>
      buildNotesXml(
        'endnote',
        [{ id: '1', text: 'bold note', origin: { kind: 'footnote', id: '2' } }],
        null,
        { xml: footnotes, styleIds: new Set(styleIds) },
      )
    expect(convert(['EndnoteText', 'EndnoteReference'])).toContain(
      '<w:endnote w:id="1"><w:p><w:pPr><w:pStyle w:val="EndnoteText"/></w:pPr>' +
        '<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r>' +
        '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve"> bold note</w:t></w:r></w:p></w:endnote>',
    )
    // a document without endnote styles keeps the styles the note had
    expect(convert([])).toContain('<w:pStyle w:val="FootnoteText"/>')
    expect(convert([])).toContain('<w:endnoteRef/>')
  })
})
