import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  generateTableModelXml,
  parseDocx,
  patchTableCellTexts,
  saveDocx,
  type SaveBlock,
} from '../src/index'
import { buildDocx } from './helpers/build-docx'

/**
 * A fill-in template the way Word writes one: the document is protected as
 * read-only and only the answer regions are open to everyone. Inline ranges
 * cover one paragraph; column ranges (w:colFirst/w:colLast) start in a cell
 * and end between rows, opening the whole cell.
 */
const BODY = [
  '<w:p><w:r><w:t>Locked intro</w:t></w:r></w:p>',
  '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>',
  '<w:tr><w:tc><w:p><w:r><w:t>A. TITLE</w:t></w:r></w:p></w:tc></w:tr>',
  '<w:tr><w:tc><w:p><w:permStart w:id="11" w:edGrp="everyone"/><w:r><w:t>[title]</w:t></w:r><w:permEnd w:id="11"/></w:p></w:tc></w:tr>',
  '<w:tr><w:tc><w:p><w:r><w:t>B. SUMMARY</w:t></w:r></w:p></w:tc></w:tr>',
  '<w:tr><w:tc><w:p><w:permStart w:id="22" w:edGrp="everyone" w:colFirst="0" w:colLast="0"/><w:r><w:t>[first]</w:t></w:r></w:p><w:p><w:r><w:t>second</w:t></w:r></w:p></w:tc></w:tr>',
  '<w:permEnd w:id="22"/>',
  '<w:tr><w:tc><w:p><w:r><w:t>C. LOCKED</w:t></w:r></w:p></w:tc></w:tr>',
  '</w:tbl>',
  '<w:permStart w:id="33" w:edGrp="everyone"/>',
  '<w:p><w:r><w:t>open paragraph</w:t></w:r></w:p>',
  '<w:permEnd w:id="33"/>',
  '<w:p><w:permStart w:id="44" w:ed="someone@example.com"/><w:r><w:t>named editor only</w:t></w:r><w:permEnd w:id="44"/></w:p>',
  '<w:p><w:permStart w:id="55" w:edGrp="everyone"/><w:r><w:t>inline open</w:t></w:r><w:permEnd w:id="55"/></w:p>',
].join('')

const SETTINGS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:documentProtection w:edit="readOnly" w:enforcement="1"/></w:settings>'

const template = () =>
  buildDocx({
    bodyXml: BODY,
    extraParts: [
      {
        path: 'word/settings.xml',
        xml: SETTINGS,
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
      },
    ],
  })

async function partOf(bytes: Uint8Array, path: string): Promise<string> {
  return (await JSZip.loadAsync(bytes)).file(path)!.async('string')
}

const count = (xml: string, re: RegExp) => [...xml.matchAll(re)].length

/** the trailing body w:sectPr is written by saveDocx itself, not by a plan entry */
const isBodySectPr = (b: { originalXml: string | null }) => !!b.originalXml?.startsWith('<w:sectPr')

describe('editing permission ranges: parse', () => {
  it('marks the open paragraphs and cells, and nothing else', async () => {
    const doc = await parseDocx(await template())
    expect(doc.protection).toMatchObject({ edit: 'readOnly', enforced: true })
    const texts = doc.blocks
      .filter((b) => b.type === 'paragraph')
      .map((b) => [b.runs?.map((r) => r.text).join(''), b.permEditable === true])
    expect(texts).toEqual([
      ['Locked intro', false],
      ['open paragraph', true],
      // a range granted to a named editor opens nothing here
      ['named editor only', false],
      ['inline open', true],
    ])
    const table = doc.blocks.find((b) => b.type === 'table')!.table!
    expect(table.rows.map((row) => row[0]!.permEditable === true)).toEqual([
      false,
      true, // inline range inside the cell's paragraph
      false,
      true, // column range: the whole cell, both paragraphs
      false, // the range ended between the rows
    ])
  })

  it('opens only what a range covers: an end marker before a paragraph text leaves it locked', async () => {
    // Word ends a range that selected whole paragraphs at the start of the next
    // one; that paragraph, and one after a start marker that follows all text,
    // stay locked. An empty paragraph opens only inside a range.
    const body = [
      '<w:p><w:permStart w:id="1" w:edGrp="everyone"/><w:r><w:t>open</w:t></w:r></w:p>',
      '<w:p/>',
      '<w:p><w:permEnd w:id="1"/><w:r><w:t>Locked heading</w:t></w:r></w:p>',
      '<w:p/>',
      '<w:p><w:r><w:t>Locked tail</w:t></w:r><w:permStart w:id="2" w:edGrp="everyone"/></w:p>',
      '<w:p><w:r><w:t>open again</w:t></w:r><w:permEnd w:id="2"/></w:p>',
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>',
      '<w:tr><w:tc><w:p><w:permStart w:id="3" w:edGrp="everyone"/><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr>',
      '<w:tr><w:tc><w:p><w:permEnd w:id="3"/><w:r><w:t>locked cell</w:t></w:r></w:p></w:tc></w:tr>',
      '</w:tbl>',
    ].join('')
    const doc = await parseDocx(await buildDocx({ bodyXml: body }))
    const paras = doc.blocks
      .filter((b) => b.type !== 'table' && !isBodySectPr(b))
      .map((b) => [b.runs?.map((r) => r.text).join('') ?? '', b.permEditable === true])
    expect(paras).toEqual([
      ['open', true],
      ['', true],
      ['Locked heading', false],
      ['', false],
      ['Locked tail', false],
      ['open again', true],
    ])
    const table = doc.blocks.find((b) => b.type === 'table')!.table!
    expect(table.rows.map((row) => row[0]!.permEditable === true)).toEqual([true, false])
  })

  it('leaves documents without ranges untouched', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>' }))
    expect(doc.blocks.some((b) => b.permEditable)).toBe(false)
  })
})

describe('editing permission ranges: save', () => {
  it('keeps every marker and the protection when open cells are edited', async () => {
    const source = await template()
    const doc = await parseDocx(source)
    const tableIndex = doc.blocks.findIndex((b) => b.type === 'table')
    const table = doc.blocks[tableIndex]!
    // the title is retyped; the column-range cell gets a third paragraph (Enter)
    const xml = patchTableCellTexts(table.originalXml!, [
      null,
      [['[new title]']],
      null,
      [['[first edited]', 'second', 'third added']],
      null,
    ])
    const plan: SaveBlock[] = doc.blocks.flatMap((b, i): SaveBlock[] =>
      isBodySectPr(b)
        ? []
        : i === tableIndex
          ? [{ kind: 'xml', xml, docxIndex: i }]
          : [{ kind: 'original', docxIndex: i }],
    )
    const saved = await saveDocx(doc, plan)
    const out = await partOf(saved, 'word/document.xml')
    for (const id of ['11', '22', '33', '44', '55']) {
      expect(count(out, new RegExp(`<w:permStart w:id="${id}"`, 'g'))).toBe(1)
      expect(count(out, new RegExp(`<w:permEnd w:id="${id}"`, 'g'))).toBe(1)
    }
    expect(out).toContain('[new title]')
    expect(out).toContain('third added')
    // the column range still ends between the rows
    expect(out).toMatch(/<\/w:tr><w:permEnd w:id="22"\/><w:tr>/)
    expect(await partOf(saved, 'word/settings.xml')).toBe(SETTINGS)

    const again = await parseDocx(saved)
    const rows = again.blocks.find((b) => b.type === 'table')!.table!.rows
    expect(rows.map((row) => row[0]!.permEditable === true)).toEqual([
      false,
      true,
      false,
      true,
      false,
    ])
    expect(rows[3]![0]!.paras).toEqual(['[first edited]', 'second', 'third added'])
  })

  it('keeps the markers of a regenerated top-level paragraph, once each across a split', async () => {
    const doc = await parseDocx(await template())
    const inline = doc.blocks.findIndex(
      (b) => b.runs?.map((r) => r.text).join('') === 'inline open',
    )
    const plan: SaveBlock[] = doc.blocks.flatMap((b, i): SaveBlock[] =>
      isBodySectPr(b)
        ? []
        : i === inline
          ? [
              {
                kind: 'generated',
                block: { type: 'paragraph', runs: [{ text: 'inline' }] },
                docxIndex: i,
              },
              {
                kind: 'generated',
                block: { type: 'paragraph', runs: [{ text: 'split off' }] },
                docxIndex: i,
              },
            ]
          : [{ kind: 'original', docxIndex: i }],
    )
    const out = await partOf(await saveDocx(doc, plan), 'word/document.xml')
    expect(count(out, /<w:sectPr\b/g)).toBe(1)
    expect(count(out, /<w:permStart w:id="55"/g)).toBe(1)
    expect(count(out, /<w:permEnd w:id="55"/g)).toBe(1)
    // start in the first half, end in the second
    expect(out).toMatch(/<w:permStart w:id="55"[^>]*\/>(?:(?!<\/w:p>).)*inline<\/w:t>/)
    expect(out).toMatch(/split off<\/w:t>(?:(?!<w:p[ >]).)*<w:permEnd w:id="55"\/>/)
  })

  it('closes a range whose end paragraph was joined into the one above or deleted', async () => {
    // Locked intro / one (start) / two (end) / Locked outro: Backspace at the
    // start of "two" joins it into "one"; the end must survive on the merged
    // paragraph, or on reopen the range runs to the end of the document
    const body = [
      '<w:p><w:r><w:t>Locked intro</w:t></w:r></w:p>',
      '<w:p><w:permStart w:id="7" w:edGrp="everyone"/><w:r><w:t>one</w:t></w:r></w:p>',
      '<w:p><w:r><w:t>two</w:t></w:r><w:permEnd w:id="7"/></w:p>',
      '<w:p><w:r><w:t>Locked outro</w:t></w:r></w:p>',
    ].join('')
    const doc = await parseDocx(await buildDocx({ bodyXml: body }))
    const text = (b: (typeof doc.blocks)[number]) => b.runs?.map((r) => r.text).join('') ?? ''
    const reopen = async (plan: SaveBlock[]) => {
      const saved = await saveDocx(doc, plan)
      const xml = await partOf(saved, 'word/document.xml')
      expect(count(xml, /<w:permStart w:id="7"/g)).toBe(1)
      expect(count(xml, /<w:permEnd w:id="7"/g)).toBe(1)
      const again = await parseDocx(saved)
      return again.blocks
        .filter((b) => !isBodySectPr(b))
        .map((b) => [text(b), b.permEditable === true])
    }
    const one = doc.blocks.findIndex((b) => text(b) === 'one')
    const two = doc.blocks.findIndex((b) => text(b) === 'two')
    const keep = (i: number): SaveBlock[] =>
      isBodySectPr(doc.blocks[i]!) ? [] : [{ kind: 'original', docxIndex: i }]

    // joined: "one" regenerated as "onetwo", "two" gone
    const joined = doc.blocks.flatMap((_, i): SaveBlock[] =>
      i === two
        ? []
        : i === one
          ? [
              {
                kind: 'generated',
                block: { type: 'paragraph', runs: [{ text: 'onetwo' }] },
                docxIndex: i,
              },
            ]
          : keep(i),
    )
    expect(await reopen(joined)).toEqual([
      ['Locked intro', false],
      ['onetwo', true],
      ['Locked outro', false],
    ])

    // the end's paragraph deleted outright
    const deleted = doc.blocks.flatMap((_, i): SaveBlock[] => (i === two ? [] : keep(i)))
    expect(await reopen(deleted)).toEqual([
      ['Locked intro', false],
      ['one', true],
      ['Locked outro', false],
    ])
  })

  it('puts the markers back into a table rebuilt from its model', async () => {
    const doc = await parseDocx(await template())
    const table = doc.blocks.find((b) => b.type === 'table')!
    // a paragraph added to a cell regenerates the whole table (formatting signature)
    const model = structuredClone(table.table!)
    model.rows[3]![0]!.paras = [...model.rows[3]![0]!.paras, 'added']
    model.rows[3]![0]!.richParas = undefined
    const rebuilt = generateTableModelXml(model, table.originalXml!)
    for (const id of ['11', '22']) {
      expect(count(rebuilt, new RegExp(`<w:permStart w:id="${id}"`, 'g'))).toBe(1)
      expect(count(rebuilt, new RegExp(`<w:permEnd w:id="${id}"`, 'g'))).toBe(1)
    }
    // the column range still ends between the same rows
    expect(rebuilt).toMatch(/added<\/w:t>(?:(?!<w:tr[ >]).)*<\/w:tr><w:permEnd w:id="22"\/><w:tr/)
  })

  it('saves an untouched protected template byte for byte', async () => {
    const source = await template()
    const doc = await parseDocx(source)
    const saved = await saveDocx(
      doc,
      doc.blocks.flatMap((b, i): SaveBlock[] =>
        isBodySectPr(b) ? [] : [{ kind: 'original', docxIndex: i }],
      ),
    )
    expect(saved).toEqual(source)
  })
})
