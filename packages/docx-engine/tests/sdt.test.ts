import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

// ---- SDT fixture XML fragments ----

/** Block-level SDT wrapping a text paragraph with alias "Author" */
const BLOCK_SDT_P =
  '<w:sdt>' +
  '<w:sdtPr>' +
  '<w:alias w:val="Author Name"/>' +
  '<w:tag w:val="author"/>' +
  '<w:text/>' +
  '</w:sdtPr>' +
  '<w:sdtEndPr/>' +
  '<w:sdtContent>' +
  '<w:p><w:r><w:t>Jane Doe</w:t></w:r></w:p>' +
  '</w:sdtContent>' +
  '</w:sdt>'

/** Block-level SDT with date picker */
const DATE_SDT_P =
  '<w:sdt>' +
  '<w:sdtPr>' +
  '<w:alias w:val="Date"/>' +
  '<w:date>' +
  '<w:dateFormat w:val="yyyy-MM-dd"/>' +
  '</w:date>' +
  '</w:sdtPr>' +
  '<w:sdtContent>' +
  '<w:p><w:r><w:t>2026-07-26</w:t></w:r></w:p>' +
  '</w:sdtContent>' +
  '</w:sdt>'

/** Block-level SDT with dropdown */
const DROPDOWN_SDT_P =
  '<w:sdt>' +
  '<w:sdtPr>' +
  '<w:alias w:val="Status"/>' +
  '<w:dropDownList>' +
  '<w:listItem w:value="Draft" w:displayText="Draft"/>' +
  '<w:listItem w:value="Final" w:displayText="Final"/>' +
  '</w:dropDownList>' +
  '</w:sdtPr>' +
  '<w:sdtContent>' +
  '<w:p><w:r><w:t>Draft</w:t></w:r></w:p>' +
  '</w:sdtContent>' +
  '</w:sdt>'

/** Plain paragraph (no SDT) */
const PLAIN_P = '<w:p><w:r><w:t>plain text no sdt</w:t></w:r></w:p>'

describe('SDT parsing', () => {
  it('parses block-level w:sdt and exposes alias/tag/controlType', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BLOCK_SDT_P }))
    const block = doc.blocks[0]
    expect(block.sdtShell).toBeDefined()
    expect(block.sdtShell!.alias).toBe('Author Name')
    expect(block.sdtShell!.tag).toBe('author')
    expect(block.sdtShell!.controlType).toBe('text')
  })

  it('parses date SDT as controlType "date"', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: DATE_SDT_P }))
    expect(doc.blocks[0].sdtShell?.controlType).toBe('date')
    expect(doc.blocks[0].sdtShell?.alias).toBe('Date')
  })

  it('parses dropdown SDT as controlType "dropdown"', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: DROPDOWN_SDT_P }))
    expect(doc.blocks[0].sdtShell?.controlType).toBe('dropdown')
  })

  it('extracts the sdtContent paragraph runs as editable block runs', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BLOCK_SDT_P }))
    const block = doc.blocks[0]
    expect(block.runs?.some((r) => r.text.includes('Jane Doe'))).toBe(true)
  })

  it('plain paragraph has no sdtShell', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: PLAIN_P }))
    expect(doc.blocks[0].sdtShell).toBeUndefined()
  })
})

describe('SDT roundtrip', () => {
  it('untouched SDT block saves byte-identical (sdt shell preserved)', async () => {
    const bytes = await buildDocx({ bodyXml: BLOCK_SDT_P })
    const doc = await parseDocx(bytes)
    const saved = await saveDocx(doc, [{ kind: 'original', docxIndex: 0 }])
    const reparsed = await parseDocx(saved)
    // The shell must be preserved
    expect(reparsed.blocks[0].sdtShell?.alias).toBe('Author Name')
    // sdtContent text must be preserved
    expect(reparsed.blocks[0].runs?.some((r) => r.text.includes('Jane Doe'))).toBe(true)
    // Re-saved XML must contain the sdt shell elements
    expect(reparsed.internal.documentXml).toContain('<w:sdt>')
    expect(reparsed.internal.documentXml).toContain('w:val="Author Name"')
    expect(reparsed.internal.documentXml).toContain('Jane Doe')
  })

  it('sdt shell openXml contains original sdtPr bytes', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BLOCK_SDT_P }))
    const shell = doc.blocks[0].sdtShell!
    expect(shell.openXml).toContain('<w:sdt>')
    expect(shell.openXml).toContain('<w:sdtPr>')
    expect(shell.openXml).toContain('Author Name')
    expect(shell.openXml).toContain('<w:sdtContent>')
    expect(shell.closeXml).toBe('</w:sdtContent></w:sdt>')
  })
})

// ---- Multi-paragraph SDT (Word automatic TOC) ----

const TOC_ENTRY = (n: number) =>
  '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>' +
  `<w:r><w:t>Chapter ${n}</w:t></w:r>` +
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> PAGEREF _Toc${n} \\h </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  `<w:r><w:t>${n}</w:t></w:r>` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
  '</w:p>'

const TOC_SDT =
  '<w:sdt>' +
  '<w:sdtPr><w:id w:val="-1"/>' +
  '<w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj>' +
  '</w:sdtPr>' +
  '<w:sdtEndPr/>' +
  '<w:sdtContent>' +
  '<w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr><w:r><w:t>Contents</w:t></w:r></w:p>' +
  Array.from({ length: 13 }, (_, i) => TOC_ENTRY(i + 1)).join('') +
  '<w:p><w:r><w:t>tail</w:t></w:r></w:p>' +
  '</w:sdtContent>' +
  '</w:sdt>'

const sdtOpenCount = (xml: string) => (xml.match(/<w:sdt>/g) ?? []).length
const sdtCloseCount = (xml: string) => (xml.match(/<\/w:sdt>/g) ?? []).length

describe('multi-paragraph SDT (TOC)', () => {
  it('yields one block per sdtContent child paragraph', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: TOC_SDT }))
    const visible = doc.blocks.filter((b) => !b.hidden)
    expect(visible.length).toBe(15)
    expect(visible[0].runs?.some((r) => r.text.includes('Contents'))).toBe(true)
    for (let n = 1; n <= 13; n++) {
      expect(visible[n].type).toBe('passthrough')
      expect(visible[n].previewText).toContain(`Chapter ${n}`)
    }
    expect(visible.map((b) => b.docxIndex)).toEqual(visible.map((_, i) => i))
  })

  it('splits the shell: first block opens, last block closes, middles carry none', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: TOC_SDT }))
    const visible = doc.blocks.filter((b) => !b.hidden)
    const shells = visible.map((b) => b.sdtShell!)
    expect(shells.every((s) => s !== undefined)).toBe(true)
    expect(new Set(shells.map((s) => s.group)).size).toBe(1)
    expect(shells[0].group).not.toBeUndefined()
    expect(shells[0].openXml).toContain('<w:sdt>')
    expect(shells[0].openXml).toContain('Table of Contents')
    expect(shells[0].openXml).toContain('<w:sdtContent>')
    expect(shells[0].closeXml).toBe('')
    for (const s of shells.slice(1, -1)) {
      expect(s.openXml).toBe('')
      expect(s.closeXml).toBe('')
    }
    expect(shells[14].openXml).toBe('')
    expect(shells[14].closeXml).toBe('</w:sdtContent></w:sdt>')
  })

  it('untouched multi-paragraph sdt saves byte-identical', async () => {
    const bytes = await buildDocx({ bodyXml: TOC_SDT })
    const doc = await parseDocx(bytes)
    const saved = await saveDocx(
      doc,
      doc.blocks
        .filter((b) => !b.hidden)
        .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! })),
    )
    expect(Buffer.from(saved).equals(Buffer.from(bytes))).toBe(true)
  })

  it('regenerating a middle paragraph keeps a single balanced sdt', async () => {
    const bytes = await buildDocx({ bodyXml: TOC_SDT })
    const doc = await parseDocx(bytes)
    const visible = doc.blocks.filter((b) => !b.hidden)
    const finalBlocks = visible.map((b, i) =>
      i === 7
        ? {
            kind: 'generated' as const,
            block: {
              type: 'paragraph' as const,
              runs: [{ text: 'edited entry' }],
              sdtShell: b.sdtShell,
            },
          }
        : { kind: 'original' as const, docxIndex: b.docxIndex! },
    )
    const saved = await saveDocx(doc, finalBlocks)
    const reparsed = await parseDocx(saved)
    const xml = reparsed.internal.documentXml
    expect(sdtOpenCount(xml)).toBe(1)
    expect(sdtCloseCount(xml)).toBe(1)
    const contentOpen = xml.indexOf('<w:sdtContent>')
    const contentClose = xml.indexOf('</w:sdtContent>')
    const edited = xml.indexOf('edited entry')
    expect(edited).toBeGreaterThan(contentOpen)
    expect(edited).toBeLessThan(contentClose)
    expect(reparsed.blocks.filter((b) => !b.hidden).length).toBe(15)
  })

  it('regenerating the first and last paragraphs re-emits their shell halves', async () => {
    const bytes = await buildDocx({ bodyXml: TOC_SDT })
    const doc = await parseDocx(bytes)
    const visible = doc.blocks.filter((b) => !b.hidden)
    const finalBlocks = visible.map((b, i) =>
      i === 0 || i === visible.length - 1
        ? {
            kind: 'generated' as const,
            block: {
              type: 'paragraph' as const,
              runs: [{ text: i === 0 ? 'New Title' : 'new tail' }],
              sdtShell: b.sdtShell,
            },
          }
        : { kind: 'original' as const, docxIndex: b.docxIndex! },
    )
    const saved = await saveDocx(doc, finalBlocks)
    const xml = (await parseDocx(saved)).internal.documentXml
    expect(sdtOpenCount(xml)).toBe(1)
    expect(sdtCloseCount(xml)).toBe(1)
    expect(xml.indexOf('New Title')).toBeGreaterThan(xml.indexOf('<w:sdtContent>'))
    expect(xml.indexOf('new tail')).toBeLessThan(xml.indexOf('</w:sdtContent>'))
  })

  it('sdt wrapping a table plus a paragraph yields a table block and a paragraph block', async () => {
    const MIXED_SDT =
      '<w:sdt><w:sdtPr><w:alias w:val="Report"/></w:sdtPr><w:sdtContent>' +
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>cell A</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
      '<w:p><w:r><w:t>after table</w:t></w:r></w:p>' +
      '</w:sdtContent></w:sdt>'
    const bytes = await buildDocx({ bodyXml: MIXED_SDT })
    const doc = await parseDocx(bytes)
    const visible = doc.blocks.filter((b) => !b.hidden)
    expect(visible.length).toBe(2)
    expect(visible[0].type).toBe('table')
    expect(visible[1].runs?.some((r) => r.text.includes('after table'))).toBe(true)
    const saved = await saveDocx(
      doc,
      visible.map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! })),
    )
    expect(Buffer.from(saved).equals(Buffer.from(bytes))).toBe(true)
  })
})

describe('nested SDT (Word cover-page building blocks)', () => {
  // outer docPartObj sdt wrapping several inner field sdts + plain paragraphs
  const NESTED_SDT =
    '<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Cover Pages"/></w:docPartObj></w:sdtPr><w:sdtContent>' +
    '<w:sdt><w:sdtPr><w:alias w:val="Title"/><w:text/></w:sdtPr><w:sdtContent>' +
    '<w:p><w:r><w:t>Resume Title</w:t></w:r></w:p>' +
    '</w:sdtContent></w:sdt>' +
    '<w:p><w:r><w:t>plain between</w:t></w:r></w:p>' +
    '<w:sdt><w:sdtPr><w:alias w:val="Company"/><w:text/></w:sdtPr><w:sdtContent>' +
    '<w:p><w:r><w:t>ACME Corp</w:t></w:r></w:p>' +
    '</w:sdtContent></w:sdt>' +
    '</w:sdtContent></w:sdt>'

  it('descends nested content controls: every inner paragraph becomes a block', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: NESTED_SDT }))
    const visible = doc.blocks.filter((b) => !b.hidden)
    // regression: only "Resume Title" survived; "plain between" and "ACME Corp" were dropped
    expect(visible.map((b) => b.runs?.map((r) => r.text).join('') ?? b.previewText)).toEqual([
      'Resume Title',
      'plain between',
      'ACME Corp',
    ])
  })

  it('untouched nested sdt saves byte-identical', async () => {
    const bytes = await buildDocx({ bodyXml: NESTED_SDT })
    const doc = await parseDocx(bytes)
    const saved = await saveDocx(
      doc,
      doc.blocks
        .filter((b) => !b.hidden)
        .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! })),
    )
    expect(Buffer.from(saved).equals(Buffer.from(bytes))).toBe(true)
  })

  it('sdtContent without any paragraph keeps its text as previewText', async () => {
    const RUN_ONLY_SDT =
      '<w:sdt><w:sdtPr><w:alias w:val="Odd"/></w:sdtPr><w:sdtContent>' +
      '<w:r><w:t>bare run text</w:t></w:r>' +
      '</w:sdtContent></w:sdt>'
    const doc = await parseDocx(await buildDocx({ bodyXml: RUN_ONLY_SDT }))
    const visible = doc.blocks.filter((b) => !b.hidden)
    expect(visible[0].type).toBe('passthrough')
    expect(visible[0].previewText).toContain('bare run text')
  })
})
