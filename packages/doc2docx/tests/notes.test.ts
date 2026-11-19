import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '@genoffice/docx-engine'
import {
  LegacyDocError,
  convertLegacyDocToDocx,
  readLegacyDocBody,
  readLegacyDocLayout,
} from '../src/index'
import {
  SYNTHETIC_TEXT_OFFSET,
  buildSyntheticDoc,
  encodeChpxFkp,
  encodePapxFkp,
  encodeTableCellPapx,
  encodeTableRowPapx,
  type SyntheticPapxEntry,
} from './helpers/synthetic-doc'

function textFc(cp: number): number {
  return SYNTHETIC_TEXT_OFFSET + cp * 2
}

function sprm(opcode: number, operand: number[]): Uint8Array {
  return Uint8Array.from([opcode & 0xff, (opcode >> 8) & 0xff, ...operand])
}

async function partsOf(docx: Uint8Array): Promise<Record<string, string>> {
  const zip = await JSZip.loadAsync(docx)
  const out: Record<string, string> = {}
  for (const name of [
    'word/document.xml',
    'word/footnotes.xml',
    'word/endnotes.xml',
    'word/settings.xml',
    'word/styles.xml',
    'word/_rels/document.xml.rels',
    '[Content_Types].xml',
  ]) {
    const file = zip.file(name)
    if (file) out[name] = await file.async('string')
  }
  return out
}

describe('footnotes', () => {
  const bytes = buildSyntheticDoc({
    pieces: [{ text: 'Hello\u0002 world\r' }],
    footnotes: [{ cp: 5, text: '\u0002Note one' }],
    headerStories: [undefined, 'Head'],
  })

  it('reads the reference mark and the note text alongside the header story', () => {
    const body = readLegacyDocBody(bytes)
    expect(body.warnings).toEqual([])
    expect(body.notes?.marks).toEqual([
      { cp: 5, kind: 'footnote', id: 1, role: 'reference', custom: false },
    ])
    const note = body.notes!.footnotes[0]!
    expect(note.custom).toBe(false)
    expect(note.paragraphs.map((paragraph) => paragraph.text)).toEqual(['Note one'])
    expect(note.paragraphs[0]!.runs[0]!.note?.role).toBe('self')
    const block = body.blocks[0]!
    if (block.type !== 'paragraph') throw new Error('expected a paragraph')
    expect(block.paragraph.runs.map((run) => [run.text, run.note?.id])).toEqual([
      ['Hello', undefined],
      ['', 1],
      [' world', undefined],
    ])
    expect(readLegacyDocLayout(bytes).headerFooter[0]?.oddHeader?.text).toBe('Head')
  })

  it('writes footnotes.xml, the reference run, the relationship and the styles', async () => {
    const { docx, warnings } = await convertLegacyDocToDocx(bytes)
    expect(warnings.some((warning) => warning.includes('footnote'))).toBe(false)
    const parts = await partsOf(docx)
    expect(parts['word/document.xml']).toContain(
      '<w:t xml:space="preserve">Hello</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>' +
        '<w:footnoteReference w:id="1"/></w:r><w:r><w:t xml:space="preserve"> world</w:t>',
    )
    expect(parts['word/footnotes.xml']).toContain('<w:footnote w:type="separator" w:id="-1">')
    expect(parts['word/footnotes.xml']).toContain(
      '<w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>' +
        '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>' +
        '<w:r><w:t xml:space="preserve">Note one</w:t></w:r></w:p></w:footnote>',
    )
    expect(parts['word/settings.xml']).toContain('<w:footnotePr><w:footnote w:id="-1"/>')
    expect(parts['word/styles.xml']).toContain('w:styleId="FootnoteText"')
    expect(parts['word/styles.xml']).toContain('w:styleId="FootnoteReference"')
    expect(parts['[Content_Types].xml']).toContain('PartName="/word/footnotes.xml"')
    expect(parts['word/_rels/document.xml.rels']).toMatch(
      /relationships\/footnotes" Target="footnotes.xml"/,
    )
    expect(parts['word/endnotes.xml']).toBeUndefined()
    const parsed = await parseDocx(docx)
    expect(parsed.footnotes.map((note) => [note.id, note.text])).toEqual([['1', 'Note one']])
    const paragraph = parsed.blocks.find((block) => block.type === 'paragraph')!
    expect(paragraph.runs?.map((run) => run.noteRef?.id ?? run.text)).toEqual([
      'Hello',
      '1',
      ' world',
    ])
  })

  it('keeps two notes in body order with ids 1 and 2 and the resolved superscript', async () => {
    const two = buildSyntheticDoc({
      pieces: [{ text: 'a\u0002b\u0002\r' }],
      footnotes: [
        { cp: 1, text: '\u0002First' },
        { cp: 3, text: '\u0002Second\rmore' },
      ],
      chpxFkpPages: [
        encodeChpxFkp([
          { fcStart: textFc(0), fcEnd: textFc(1) },
          { fcStart: textFc(1), fcEnd: textFc(2), grpprl: sprm(0x2a48, [1]) },
          { fcStart: textFc(2), fcEnd: textFc(5) },
        ]),
      ],
    })
    const { docx } = await convertLegacyDocToDocx(two)
    const parts = await partsOf(docx)
    const document = parts['word/document.xml']!
    expect(document).toContain(
      '<w:rPr><w:rStyle w:val="FootnoteReference"/><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/>',
    )
    expect(document.indexOf('w:id="1"')).toBeLessThan(document.indexOf('w:id="2"'))
    const parsed = await parseDocx(docx)
    expect(parsed.footnotes.map((note) => [note.id, note.text])).toEqual([
      ['1', 'First'],
      ['2', 'Second\nmore'],
    ])
  })

  it('writes a custom mark as customMarkFollows plus the mark text', async () => {
    const custom = buildSyntheticDoc({
      pieces: [{ text: 'A*B\r' }],
      footnotes: [{ cp: 1, text: '*Custom note', custom: true }],
    })
    const { docx } = await convertLegacyDocToDocx(custom)
    const parts = await partsOf(docx)
    expect(parts['word/document.xml']).toContain(
      '<w:r><w:footnoteReference w:customMarkFollows="1" w:id="1"/><w:t xml:space="preserve">*</w:t></w:r>',
    )
    const note = parts['word/footnotes.xml']!
    expect(note).not.toContain('<w:footnoteRef/>')
    expect(note).toContain('<w:t xml:space="preserve">*Custom note</w:t>')
    const parsed = await parseDocx(docx)
    expect(parsed.footnotes[0]?.text).toBe('*Custom note')
  })

  it('writes a reference inside a table cell', async () => {
    const text = 'C1\u0002\u0007C2\u0007\u0007After\r'
    const entries: SyntheticPapxEntry[] = [
      { fcStart: textFc(0), fcEnd: textFc(4), istd: 0, grpprl: encodeTableCellPapx(1) },
      { fcStart: textFc(4), fcEnd: textFc(7), istd: 0, grpprl: encodeTableCellPapx(1) },
      {
        fcStart: textFc(7),
        fcEnd: textFc(8),
        istd: 0,
        grpprl: encodeTableRowPapx({ gridWidths: [2880, 5760], cells: [{}, {}] }, 1),
      },
      { fcStart: textFc(8), fcEnd: textFc(text.length), istd: 0, grpprl: new Uint8Array(0) },
    ]
    const table = buildSyntheticDoc({
      pieces: [{ text }],
      papxFkpPages: [encodePapxFkp(entries)],
      footnotes: [{ cp: 2, text: '\u0002Cell note' }],
    })
    const { docx } = await convertLegacyDocToDocx(table)
    const parts = await partsOf(docx)
    expect(parts['word/document.xml']).toMatch(
      /<w:tc>.*?C1<\/w:t><\/w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"\/><\/w:rPr><w:footnoteReference w:id="1"\/><\/w:r><\/w:p><\/w:tc>/,
    )
    const parsed = await parseDocx(docx)
    expect(parsed.footnotes.map((note) => note.text)).toEqual(['Cell note'])
  })

  it('degrades to a warning when the footnote text table is missing', async () => {
    const degraded = buildSyntheticDoc({
      pieces: [{ text: 'x\u0002\r' }],
      footnotes: [{ cp: 1, text: '\u0002Lost' }],
      omitNoteTextTables: true,
    })
    const { docx, warnings } = await convertLegacyDocToDocx(degraded)
    expect(warnings).toContain(
      'the document declares footnote text but has no footnote text table; footnotes are omitted',
    )
    const parts = await partsOf(docx)
    expect(parts['word/footnotes.xml']).toBeUndefined()
    expect(parts['word/document.xml']).not.toContain('footnoteReference')
  })

  it('rejects a footnote text table that does not match the reference count', async () => {
    const broken = buildSyntheticDoc({
      pieces: [{ text: 'x\u0002\r' }],
      footnotes: [{ cp: 1, text: '\u0002Note' }],
      lcbPlcffndTxt: 8,
    })
    await expect(convertLegacyDocToDocx(broken)).rejects.toMatchObject({ code: 'bad_note' })
    await expect(convertLegacyDocToDocx(broken)).rejects.toBeInstanceOf(LegacyDocError)
  })
})

describe('endnotes', () => {
  it('writes endnotes.xml and the endnote reference', async () => {
    const bytes = buildSyntheticDoc({
      pieces: [{ text: 'E\u0002\r' }],
      endnotes: [{ cp: 1, text: '\u0002End note' }],
      headerStories: [undefined, 'Head'],
    })
    const { docx, warnings } = await convertLegacyDocToDocx(bytes)
    expect(warnings.some((warning) => warning.includes('endnote'))).toBe(false)
    const parts = await partsOf(docx)
    expect(parts['word/document.xml']).toContain(
      '<w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="1"/>',
    )
    expect(parts['word/endnotes.xml']).toContain(
      '<w:endnote w:id="1"><w:p><w:pPr><w:pStyle w:val="EndnoteText"/></w:pPr>' +
        '<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r>' +
        '<w:r><w:t xml:space="preserve">End note</w:t></w:r></w:p></w:endnote>',
    )
    expect(parts['word/settings.xml']).toContain('<w:endnotePr>')
    expect(parts['word/footnotes.xml']).toBeUndefined()
    expect(readLegacyDocLayout(bytes).headerFooter[0]?.oddHeader?.text).toBe('Head')
    const parsed = await parseDocx(docx)
    expect(parsed.endnotes.map((note) => [note.id, note.text])).toEqual([['1', 'End note']])
  })
})
