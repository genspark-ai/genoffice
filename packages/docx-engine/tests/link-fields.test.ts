import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import {
  blockLinkFields,
  breakBlockLinkField,
  editBlockLinkField,
  includeTextInstruction,
  isWordDocumentLink,
  linkResultIsPlain,
  parseDocx,
  parseLinkInstruction,
  saveDocx,
  withLinkAutomatic,
  withLinkSource,
  wordLinkInstruction,
} from '../src/index'
import { buildDocx, IMAGE_PARAGRAPH_XML, TINY_PNG_BASE64 } from './helpers/build-docx'

const begin = (instr: string, attrs = '') =>
  `<w:r><w:fldChar w:fldCharType="begin"${attrs}/></w:r>` +
  `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
const END = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
const text = (s: string) => `<w:r><w:t xml:space="preserve">${s}</w:t></w:r>`

const marks = (doc: Awaited<ReturnType<typeof parseDocx>>) =>
  doc.blocks
    .slice(0, -1)
    .map((b) =>
      b.type !== 'paragraph'
        ? b.type
        : (b.runs ?? [])
            .map((r) => (r.linkBegin !== undefined ? '[' : r.linkEnd ? ']' : r.text))
            .join(''),
    )

describe('link fields', () => {
  it('read as two markers around ordinary content, across paragraphs and a table', async () => {
    const body =
      `<w:p>${text('Before ')}${begin('INCLUDETEXT "C:\\\\docs\\\\part.docx" Intro', ' w:fldLock="1"')}${text('First')}</w:p>` +
      `<w:p>${text('Second')}</w:p>` +
      '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
      `<w:p>${END}${text('after')}</w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: body }))
    expect(marks(doc)).toEqual(['Before [First', 'Second', 'table', ']after'])
    const open = doc.blocks[0].runs!.find((r) => r.linkBegin !== undefined)!
    expect(open.fldLock).toBe(true)
    expect(parseLinkInstruction(open.linkBegin!)).toEqual({
      kind: 'INCLUDETEXT',
      app: '',
      source: 'C:\\docs\\part.docx',
      item: 'Intro',
      automatic: false,
    })
    expect(doc.blocks[0].originalXml).toContain('INCLUDETEXT')
    expect(doc.blocks[0].originalXml).not.toContain('GENOFFICE')
  })

  it('write the field characters back where the markers stand', async () => {
    const body = `<w:p>${begin('INCLUDETEXT "a.docx"')}${text('One')}</w:p><w:p>${text('Two')}${END}</w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: body }))
    const saved = await saveDocx(doc, [
      {
        kind: 'generated',
        docxIndex: 0,
        block: {
          type: 'paragraph',
          runs: doc.blocks[0].runs!.map((r) => (r.linkBegin ? { ...r, fldLock: true } : r)),
        },
      },
      {
        kind: 'generated',
        docxIndex: 1,
        block: { type: 'paragraph', runs: [{ text: 'New ' }, ...doc.blocks[1].runs!] },
      },
      { kind: 'original', docxIndex: 2 },
    ])
    const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    expect(xml).toContain(
      '<w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> INCLUDETEXT "a.docx" </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>',
    )
    expect(xml.match(/w:fldCharType="end"/g)).toHaveLength(1)
    expect(xml).not.toContain('GENOFFICE')
    expect(marks(await parseDocx(saved)).slice(0, 2)).toEqual(['[One', 'New Two]'])
  })

  it('leave a field alone when one of its paragraphs cannot stay editable', async () => {
    const body =
      `<w:p>${begin('INCLUDETEXT "a.docx"')}${text('One')}</w:p>` +
      `<w:p>${begin('TOC \\o')}${text('x')}${END}${END}</w:p>` +
      `<w:p>${begin('LINK Excel.Sheet.12 "b.xlsx" "Sheet1!R1C1" \\a \\p')}${text('7')}${END}</w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: body }))
    expect(doc.blocks.slice(0, 3).map((b) => b.type)).toEqual([
      'passthrough',
      'passthrough',
      'paragraph',
    ])
    expect(marks(doc)[2]).toBe('[7]')
  })

  it('leave a picture or object field with nothing else in its paragraph as it was', async () => {
    const body =
      `<w:p>${begin('INCLUDEPICTURE "http://x/y.png" \\d')}${END}</w:p>` +
      `<w:p>${begin('INCLUDETEXT "a.docx"')}${END}</w:p>`
    const doc = await parseDocx(await buildDocx({ bodyXml: body }))
    expect(doc.blocks.slice(0, 2).map((b) => b.type)).toEqual(['passthrough', 'paragraph'])
    expect(blockLinkFields(doc.blocks[0].originalXml!)).toHaveLength(1)
  })

  it('a link pasted from a document is read as markers even when it ends in an empty paragraph', async () => {
    const pasted = wordLinkInstruction('/tmp/source.docx', 'LinkSource1', 'r')
    expect(pasted).toBe('LINK Word.Document.12 "/tmp/source.docx" "LinkSource1" \\r')
    const info = parseLinkInstruction(pasted)!
    expect(info).toMatchObject({ kind: 'LINK', source: '/tmp/source.docx', item: 'LinkSource1' })
    expect(isWordDocumentLink(info)).toBe(true)
    expect(isWordDocumentLink(parseLinkInstruction(' LINK Excel.Sheet.12 "b.xlsx" "S!R1" ')!)).toBe(
      false,
    )
    expect(linkResultIsPlain(pasted)).toBe(false)
    expect(linkResultIsPlain(wordLinkInstruction('/tmp/source.docx', 'LinkSource1', 't'))).toBe(
      true,
    )
    const body = (instr: string) =>
      `<w:p>${text('Host')}${begin(instr)}</w:p><w:p>${text('Linked')}</w:p><w:p>${END}</w:p>`
    expect(marks(await parseDocx(await buildDocx({ bodyXml: body(pasted) })))).toEqual([
      'Host[',
      'Linked',
      ']',
    ])
    // a LINK field of another program keeps its paragraphs as they were
    const other = await parseDocx(
      await buildDocx({
        bodyXml: body('LINK Word.Document.12 "/tmp/source.docx" "OLE_LINK1" \\a \\r'),
      }),
    )
    expect(other.blocks.slice(0, 3).map((b) => b.type)).toEqual([
      'passthrough',
      'paragraph',
      'passthrough',
    ])
  })

  it('reads and rewrites instructions', () => {
    const link = ' LINK Excel.Sheet.12 "C:\\\\b.xlsx" "Sheet1!R1C1" \\a \\f 4 \\p '
    expect(parseLinkInstruction(link)).toEqual({
      kind: 'LINK',
      app: 'Excel.Sheet.12',
      source: 'C:\\b.xlsx',
      item: 'Sheet1!R1C1',
      automatic: true,
    })
    expect(parseLinkInstruction(withLinkAutomatic(link, false))!.automatic).toBe(false)
    expect(withLinkAutomatic(' LINK a "b" ', true)).toBe(' LINK a "b" \\a ')
    expect(withLinkSource(link, '/tmp/c.xlsx')).toBe(
      ' LINK Excel.Sheet.12 "/tmp/c.xlsx" "Sheet1!R1C1" \\a \\f 4 \\p ',
    )
    expect(includeTextInstruction('/Users/a/My "doc".docx', 'bm')).toBe(
      'INCLUDETEXT "/Users/a/My doc.docx" bm',
    )
    expect(
      parseLinkInstruction(' INCLUDEPICTURE "http://x/y.png" \\* MERGEFORMATINET \\d '),
    ).toEqual({
      kind: 'INCLUDEPICTURE',
      app: '',
      source: 'http://x/y.png',
      item: '',
      automatic: false,
    })
    expect(parseLinkInstruction(' DATE ')).toBeNull()
  })

  it('lists, locks, renames and breaks the link fields of a block kept as XML', () => {
    const xml =
      `<w:p>${text('a')}${begin('DATE')}${text('1 May')}${END}` +
      `${begin('INCLUDEPICTURE "p.png" \\d')}<w:r><w:drawing/></w:r>${END}</w:p>`
    expect(blockLinkFields(xml)).toEqual([
      { index: 0, instr: ' INCLUDEPICTURE "p.png" \\d ', locked: false },
    ])
    const locked = editBlockLinkField(xml, 0, { locked: true, instr: ' INCLUDEPICTURE "q.png" ' })!
    expect(blockLinkFields(locked)).toEqual([
      { index: 0, instr: ' INCLUDEPICTURE "q.png" ', locked: true },
    ])
    expect(blockLinkFields(editBlockLinkField(locked, 0, { locked: false })!)[0].locked).toBe(false)
    const broken = breakBlockLinkField(xml, 0)!
    expect(broken).toBe(
      `<w:p>${text('a')}${begin('DATE')}${text('1 May')}${END}<w:r><w:drawing/></w:r></w:p>`,
    )
    expect(blockLinkFields(broken)).toEqual([])
  })

  it('read and write the file a picture is linked to', async () => {
    const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
    const linked = IMAGE_PARAGRAPH_XML.replace('r:embed="rId10"', 'r:embed="rId10" r:link="rId77"')
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: linked + IMAGE_PARAGRAPH_XML,
        withImage: true,
        extraRels: `<Relationship Id="rId77" Type="${IMAGE_REL}" Target="file:///tmp/a.png" TargetMode="External"/>`,
      }),
    )
    expect(doc.blocks.slice(0, 2).map((b) => b.imageLink)).toEqual(['file:///tmp/a.png', undefined])
    const saved = await saveDocx(doc, [
      { kind: 'xml', xml: doc.blocks[0].originalXml!, imageLink: null },
      { kind: 'xml', xml: doc.blocks[1].originalXml!, imageLink: 'file:///tmp/b.png' },
      {
        kind: 'image',
        image: {
          base64: TINY_PNG_BASE64,
          mime: 'image/png',
          widthPx: 10,
          heightPx: 10,
          linkTarget: 'file:///tmp/c%20d.png',
        },
      },
    ])
    const zip = await JSZip.loadAsync(saved)
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
    for (const target of ['file:///tmp/b.png', 'file:///tmp/c%20d.png'])
      expect(rels).toMatch(new RegExp(`Target="${target}" TargetMode="External"`))
    const again = await parseDocx(saved)
    expect(again.blocks.slice(0, 3).map((b) => [b.type, b.imageLink])).toEqual([
      ['image', undefined],
      ['image', 'file:///tmp/b.png'],
      ['image', 'file:///tmp/c%20d.png'],
    ])
    const xml = await zip.file('word/document.xml')!.async('string')
    expect(xml.match(/<a:blip\b[^>]*r:embed="[^"]+"/g)).toHaveLength(3)
  })
})
