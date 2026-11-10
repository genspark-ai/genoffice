import { describe, expect, it } from 'vitest'
import { notePropsFromXml, parseDocx, readSections } from '../src/index'
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

  it('numbers references in body order from numStart, restarting after a section-break paragraph', async () => {
    const doc = await parseDocx(
      await buildDocx({
        // file order of the notes part is 2, 3, 4; the body references 3 first
        bodyXml:
          `<w:p><w:r><w:t>a</w:t></w:r>${ref(3)}</w:p>` +
          // the section-break paragraph's own reference still belongs to section 1
          `<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr><w:r><w:t>b</w:t></w:r>${ref(2)}</w:p>` +
          `<w:p><w:r><w:t>c</w:t></w:r>${ref(4)}</w:p>`,
        sectPrExtra:
          '<w:footnotePr><w:numStart w:val="5"/><w:numRestart w:val="eachSect"/></w:footnotePr>',
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
