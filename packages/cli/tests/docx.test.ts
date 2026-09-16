import { readFileSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { run, tempDir } from './helpers'

const REPO = resolve(__dirname, '../../..')
const DOCX = join(REPO, 'apps/docs/tests/pagination-corpus/docx/01-simple-english.docx')

async function documentXml(path: string): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(path))
  return zip.file('word/document.xml')!.async('string')
}

describe('genoffice docx (docs editor under jsdom)', () => {
  it('creates a document from markdown and from a restricted-HTML fragment', async () => {
    const dir = tempDir()
    const md = join(dir, 'report.md')
    writeFileSync(
      md,
      '# Quarterly Report\n\nRevenue grew **12%**.\n\n## Items\n\n- alpha\n- beta\n\n| Name | Qty |\n| --- | --- |\n| Apple | 3 |\n',
    )
    const fromMd = join(dir, 'report.docx')
    const r = await run(['create', '--type', 'docx', '--from', md, '--out', fromMd, '--json'])
    expect(r.code).toBe(0)
    const xml = await documentXml(fromMd)
    expect(xml).toContain('Quarterly Report')
    expect(xml).toContain('<w:tbl>')
    expect(xml).toContain('<w:numPr>')

    const html = join(dir, 'brief.html')
    writeFileSync(
      html,
      '<h1>Brief</h1><p>Hello <strong>genoffice</strong>.</p><ul><li>one</li><li>two</li></ul>',
    )
    const fromHtml = join(dir, 'brief.docx')
    const h = await run(['create', '--type', 'docx', '--from', html, '--out', fromHtml, '--json'])
    expect(h.code).toBe(0)
    expect(h.json().detail).toMatchObject({ source: 'html', blocks: 4 })
    const hx = await documentXml(fromHtml)
    expect(hx).toContain('Brief')
    expect(hx).toContain('<w:pStyle w:val="Heading1"/>')
    expect(hx).toContain('<w:numPr>')
  })

  it('reads blocks and applies ops plus restricted-HTML tools to an existing document', async () => {
    const dir = tempDir()
    const copy = join(dir, 'doc.docx')
    writeFileSync(copy, readFileSync(DOCX))
    const read = await run(['docs', 'read', copy, '--json'])
    expect(read.code).toBe(0)
    const before = read.json().detail
    expect(before.blocks).toBeGreaterThan(5)
    expect(before.items[0]).toMatchObject({ index: 0, type: 'heading', level: 1 })
    const firstText = before.items[1].text as string
    expect(firstText.length).toBeGreaterThan(0)

    const ranged = await run(['docs', 'read', copy, '--range', '0-1', '--html', '--json'])
    expect(ranged.json().detail.items).toHaveLength(2)
    expect(ranged.json().detail.html).toContain('<h1')

    const ops = join(dir, 'ops.json')
    writeFileSync(
      ops,
      JSON.stringify([
        { op: 'findReplace', find: firstText.split(' ')[0], replace: 'GENOFFICE' },
        { op: 'setFont', target: { blockIndexes: [0] }, color: 'FF0000', bold: true },
        { op: 'insert_content', afterBlockIndex: 0, html: '<p>Inserted by genoffice.</p>' },
        {
          op: 'replace_blocks',
          startBlockIndex: 2,
          endBlockIndex: 2,
          html: '<h2>Replaced heading</h2>',
        },
      ]),
    )
    const dry = await run(['docs', 'apply', copy, '--ops', ops, '--dry-run', '--json'])
    expect(dry.code).toBe(0)
    expect(dry.json().detail.plan.length).toBeGreaterThan(0)
    expect(readFileSync(copy).equals(readFileSync(DOCX))).toBe(true)

    const applied = await run(['docs', 'apply', copy, '--ops', ops, '--json'])
    expect(applied.code).toBe(0)
    expect(applied.json().detail.blocks).toBe(before.blocks + 1)
    const xml = await documentXml(copy)
    expect(xml).toContain('Inserted by genoffice.')
    expect(xml).toContain('Replaced heading')
    expect(xml).toContain('GENOFFICE')
    expect(xml).toContain('<w:color w:val="FF0000"/>')

    const after = await run(['docs', 'read', copy, '--range', '1', '--json'])
    expect(after.json().detail.items[0].text).toBe('Inserted by genoffice.')
  })

  it('restructures a table: rows, columns, merges, cell format and style land in the docx', async () => {
    const dir = tempDir()
    const md = join(dir, 'grid.md')
    writeFileSync(
      md,
      '# Grid\n\n| Item | Qty | Price |\n| --- | --- | --- |\n| Apple | 3 | 1.5 |\n| Pear | 2 | 2.0 |\n',
    )
    const docx = join(dir, 'grid.docx')
    expect((await run(['create', '--type', 'docx', '--from', md, '--out', docx])).code).toBe(0)
    const read = await run(['docs', 'read', docx, '--json'])
    const tableBlock = read.json().detail.items.find((b: { type: string }) => b.type === 'table')
    expect(tableBlock).toMatchObject({ index: 1, table: { rows: 3, cols: 3 } })
    expect(tableBlock.table.merged).toBeUndefined()
    const styles = read.json().detail.tableStyles as Array<{ id: string; name: string }>
    expect(Array.isArray(styles)).toBe(true)

    const target = { nodeType: 'table' }
    const ops = join(dir, 'ops.json')
    writeFileSync(
      ops,
      JSON.stringify([
        { op: 'insertTableRow', target, row: 2 },
        { op: 'insertTableColumn', target, col: 2, position: 'before' },
        { op: 'mergeTableCells', target, range: { row: 3, col: 0, colEnd: 1 } },
        {
          op: 'setTableCellFormat',
          target,
          range: { row: 0, col: 0, colEnd: 3 },
          fill: '#1F4E79',
          vAlign: 'center',
          borders: { bottom: { style: 'double', width: 1.5, color: '#FF0000' } },
        },
        { op: 'setTableCellFormat', target, range: { row: 0, col: 3 }, width: '1in' },
        { op: 'deleteTableRow', target, row: 1 },
        ...(styles.length
          ? [{ op: 'setTableStyle', target, styleId: styles[0].id, look: { bandedRows: false } }]
          : []),
      ]),
    )
    const applied = await run(['docs', 'apply', docx, '--ops', ops, '--json'])
    expect(applied.code, applied.stdout).toBe(0)

    const after = await run(['docs', 'read', docx, '--json'])
    const t = after.json().detail.items.find((b: { type: string }) => b.type === 'table').table
    expect(t).toMatchObject({
      rows: 3,
      cols: 4,
      merged: [{ row: 2, col: 0, rowSpan: 1, colSpan: 2 }],
    })
    const xml = await documentXml(docx)
    const tbl = /<w:tbl>[\s\S]*<\/w:tbl>/.exec(xml)![0]
    expect(tbl.match(/<w:tr[ >]/g)).toHaveLength(3)
    expect(tbl.match(/<w:gridCol /g)).toHaveLength(4)
    expect(tbl).toContain('<w:gridSpan w:val="2"/>')
    expect(tbl).not.toContain('<w:vMerge')
    expect(tbl).toContain('<w:shd w:val="clear" w:color="auto" w:fill="1F4E79"/>')
    expect(tbl).toContain('<w:vAlign w:val="center"/>')
    expect(tbl).toContain('<w:bottom w:val="double" w:sz="12" w:space="0" w:color="FF0000"/>')
    expect(tbl).toContain('<w:gridCol w:w="1440"/>')
    if (styles.length) {
      expect(tbl).toContain(`<w:tblStyle w:val="${styles[0].id}"/>`)
      expect(tbl).toMatch(/<w:tblLook [^>]*w:noHBand="1"/)
    }

    const cut = join(dir, 'cut.json')
    writeFileSync(
      cut,
      JSON.stringify([
        { op: 'mergeTableCells', target, range: { row: 1, col: 1, rowEnd: 2, colEnd: 1 } },
      ]),
    )
    const refused = await run(['docs', 'apply', docx, '--ops', cut, '--json'])
    expect(refused.code).toBe(1)
    expect(refused.json().message).toContain('crosses the range')
    expect(await documentXml(docx)).toBe(xml)
  })

  it('read --full returns whole block text instead of the 200-character preview', async () => {
    const dir = tempDir()
    const long = 'lorem '.repeat(80).trim()
    writeFileSync(join(dir, 'long.md'), `# Title\n\n${long}\n`)
    const out = join(dir, 'long.docx')
    expect(
      (await run(['create', '--type', 'docx', '--from', join(dir, 'long.md'), '--out', out])).code,
    ).toBe(0)
    const preview = (await run(['docs', 'read', out, '--json'])).json().detail.items[1]
    expect(preview.text).toBe(`${long.slice(0, 200)}…(+${long.length - 200} chars)`)
    expect(preview.truncated).toBe(true)
    const short = (await run(['docs', 'read', out, '--max-chars', '50', '--json'])).json().detail
      .items[1]
    expect(short.text.startsWith(long.slice(0, 50))).toBe(true)
    expect(short.text.endsWith(`(+${long.length - 50} chars)`)).toBe(true)
    const full = (await run(['docs', 'read', out, '--full', '--json'])).json().detail.items[1]
    expect(full.text).toBe(long)
    expect(full.truncated).toBeUndefined()
    expect((await run(['docs', 'read', out, '--max-chars', '0', '--json'])).code).toBe(1)
  })

  it('inserts fields, bookmarks and notes, lists them and deletes a note', async () => {
    const dir = tempDir()
    const copy = join(dir, 'doc.docx')
    writeFileSync(copy, readFileSync(DOCX))
    const items = (await run(['docs', 'read', copy, '--full', '--json'])).json().detail
      .items as Array<{
      index: number
      type: string
      text: string
    }>
    const paras = items.filter((b) => b.type === 'paragraph' && b.text.length > 20).slice(0, 3)
    expect(paras).toHaveLength(3)
    const [a, b, c] = paras as [(typeof paras)[0], (typeof paras)[0], (typeof paras)[0]]
    const ops = join(dir, 'ops.json')
    writeFileSync(
      ops,
      JSON.stringify([
        { op: 'insertBookmark', target: { blockIndexes: [a.index] }, name: 'Opening' },
        {
          op: 'insertField',
          target: { blockIndexes: [b.index] },
          type: 'SEQ',
          args: 'Figure \\* ARABIC',
          position: 'start',
        },
        {
          op: 'insertField',
          target: { blockIndexes: [c.index] },
          type: 'REF',
          args: 'Opening \\h',
        },
        {
          op: 'insertField',
          target: { blockIndexes: [c.index] },
          type: 'DATE',
          args: '\\@ "yyyy"',
        },
        { op: 'insert_footnote', blockIndex: a.index, text: 'Source: annual report.' },
        {
          op: 'insert_endnote',
          blockIndex: b.index,
          afterText: b.text.slice(0, 8),
          text: 'See appendix.',
        },
        { op: 'updateFields' },
      ]),
    )
    const applied = await run(['docs', 'apply', copy, '--ops', ops, '--json'])
    expect(applied.code, applied.stdout).toBe(0)
    if (process.env.W5_KEEP)
      writeFileSync(join(process.env.W5_KEEP, 'fields.docx'), readFileSync(copy))

    const read = (await run(['docs', 'read', copy, '--fields', '--notes', '--json'])).json().detail
    const fields = read.fields as Array<Record<string, unknown>>
    expect(fields).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ blockIndex: c.index, type: 'REF', result: a.text }),
        expect.objectContaining({
          blockIndex: c.index,
          type: 'DATE',
          result: String(new Date().getFullYear()),
          dirty: false,
        }),
      ]),
    )
    // the SEQ paragraph re-opens protected: the field is read from its XML
    expect(fields).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          blockIndex: b.index,
          type: 'SEQ',
          result: '1',
          dirty: true,
          protected: true,
        }),
      ]),
    )
    // a field result split by character formatting is still one field
    const bold = join(dir, 'bold.json')
    writeFileSync(
      bold,
      JSON.stringify([
        { op: 'setMatchedFont', text: '202', target: { blockIndexes: [c.index] }, bold: true },
      ]),
    )
    expect((await run(['docs', 'apply', copy, '--ops', bold, '--json'])).code).toBe(0)
    const dates = (
      (await run(['docs', 'read', copy, '--fields', '--json'])).json().detail.fields as Array<
        Record<string, unknown>
      >
    ).filter((f) => f.type === 'DATE')
    expect(dates).toEqual([
      expect.objectContaining({ blockIndex: c.index, result: String(new Date().getFullYear()) }),
    ])

    const notes = read.notes as Array<Record<string, unknown>>
    expect(notes).toEqual([
      {
        kind: 'footnote',
        id: expect.any(String),
        num: 1,
        blockIndex: a.index,
        text: 'Source: annual report.',
      },
      {
        kind: 'endnote',
        id: expect.any(String),
        num: 1,
        blockIndex: b.index,
        text: 'See appendix.',
      },
    ])

    const zip = await JSZip.loadAsync(readFileSync(copy))
    const footnotes = await zip.file('word/footnotes.xml')!.async('string')
    expect(footnotes).toContain('w:type="separator"')
    expect(footnotes).toContain('Source: annual report.')
    expect(await zip.file('word/endnotes.xml')!.async('string')).toContain('See appendix.')
    expect(await zip.file('[Content_Types].xml')!.async('string')).toContain('footnotes+xml')
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
    expect(rels).toContain('relationships/footnotes')
    expect(rels).toContain('relationships/endnotes')
    const xml = await zip.file('word/document.xml')!.async('string')
    expect(xml).toContain(`<w:footnoteReference w:id="${notes[0]!.id}"/>`)
    expect(xml).toContain('w:name="Opening"')
    expect(xml).toMatch(/w:dirty="true"[\s\S]{0,60} SEQ Figure \\\* ARABIC /)

    // the endnote mark sits in the (now protected) SEQ paragraph: refused, nothing changes
    const locked = join(dir, 'locked.json')
    writeFileSync(
      locked,
      JSON.stringify([{ op: 'delete_note', kind: 'endnote', id: notes[1]!.id }]),
    )
    const refused = await run(['docs', 'apply', copy, '--ops', locked, '--json'])
    expect(refused.code).toBe(1)
    expect(refused.json().message).toContain(`protected block ${b.index}`)
    expect(
      (await run(['docs', 'read', copy, '--notes', '--json'])).json().detail.notes,
    ).toHaveLength(2)

    const del = join(dir, 'del.json')
    writeFileSync(del, JSON.stringify([{ op: 'delete_note', kind: 'footnote', id: notes[0]!.id }]))
    expect((await run(['docs', 'apply', copy, '--ops', del, '--json'])).code).toBe(0)
    const after = (await run(['docs', 'read', copy, '--notes', '--json'])).json().detail.notes
    expect(after).toHaveLength(1)
    expect(after[0].kind).toBe('endnote')
    expect(await documentXml(copy)).not.toContain('<w:footnoteReference')

    const check = await run(['docs', 'check', copy, '--json'])
    expect(check.code).toBe(0)
    const levels = (check.json().detail.issues as Array<{ level: string }>).map((i) => i.level)
    expect(levels).not.toContain('error')
  })

  it('rejects bad ops with a usage error and leaves the file untouched', async () => {
    const dir = tempDir()
    const copy = join(dir, 'doc.docx')
    writeFileSync(copy, readFileSync(DOCX))
    const bad = join(dir, 'bad.json')
    writeFileSync(bad, JSON.stringify([{ op: 'frobnicate', target: { blockIndexes: [0] } }]))
    const r = await run(['docs', 'apply', copy, '--ops', bad, '--json'])
    expect(r.code).toBe(1)
    expect(r.json().detail.failures[0]).toMatchObject({ index: 0, op: 'frobnicate' })
    expect(readFileSync(copy).equals(readFileSync(DOCX))).toBe(true)
    expect((await run(['docs', 'nope', copy])).code).toBe(1)
  })

  it('converts markdown to docx and html, and prints the docs guide', async () => {
    const dir = tempDir()
    const md = join(dir, 'notes.md')
    writeFileSync(md, '# Notes\n\nSome *text*.\n')
    const d = await run(['convert', md, '--to', 'docx', '--json'])
    expect(d.code).toBe(0)
    expect(await documentXml(join(dir, 'notes.docx'))).toContain('Notes')
    const h = await run(['convert', md, '--to', 'html', '--json'])
    expect(h.code).toBe(0)
    const html = readFileSync(join(dir, 'notes.html'), 'utf-8')
    expect(html).toContain('<h1')
    expect(html).toContain('<em>text</em>')

    const guide = await run(['guide', 'docs'])
    expect(guide.code).toBe(0)
    expect(guide.stdout).toContain('findReplace')
    expect(guide.stdout).toContain('Only these tags are allowed')
  })

  it('handles mixed batches, appends by default, and dry-runs with real index shifts', async () => {
    const dir = tempDir()
    const copy = join(dir, 'doc.docx')
    writeFileSync(copy, readFileSync(DOCX))
    const count = (await run(['docs', 'read', copy, '--json'])).json().detail.blocks as number
    const ops = join(dir, 'mixed.json')
    writeFileSync(
      ops,
      JSON.stringify([
        { op: 'insert_content', afterBlockIndex: 0, html: '<p>First insert.</p>' },
        { op: 'setFont', target: { blockIndexes: [1] }, bold: true },
        { op: 'insert_content', html: '<p>Appended.</p>' },
        {
          op: 'replace_blocks',
          startBlockIndex: count + 1,
          endBlockIndex: count + 1,
          html: '<p>Tail.</p>',
        },
      ]),
    )
    const dry = await run(['docs', 'apply', copy, '--ops', ops, '--dry-run', '--json'])
    expect(dry.code).toBe(0)
    expect(dry.json().detail.plan).toHaveLength(4)
    expect(readFileSync(copy).equals(readFileSync(DOCX))).toBe(true)

    const applied = await run(['docs', 'apply', copy, '--ops', ops, '--json'])
    expect(applied.code).toBe(0)
    const after = (await run(['docs', 'read', copy, '--json'])).json().detail
    expect(after.blocks).toBe(count + 2)
    expect(after.items[1].text).toBe('First insert.')
    expect(after.items[count + 1].text).toBe('Tail.')

    // the dry run must catch a range that only becomes invalid after an earlier insert shifts nothing
    const badLater = join(dir, 'bad-later.json')
    writeFileSync(
      badLater,
      JSON.stringify([
        { op: 'replace_blocks', startBlockIndex: 9999, endBlockIndex: 9999, html: '<p>x</p>' },
      ]),
    )
    const r = await run(['docs', 'apply', copy, '--ops', badLater, '--dry-run', '--json'])
    expect(r.code).toBe(1)
    expect(r.json().detail.failures[0].op).toBe('replace_blocks')
  })
})

describe('genoffice convert docx → md', () => {
  it('round-trips headings, marks, lists and tables through the two editors', async () => {
    const dir = tempDir()
    const md = join(dir, 'in.md')
    writeFileSync(
      md,
      '# Title\n\nSome **bold** and *italic* text with a [link](https://example.com).\n\n- alpha\n- beta\n\n1. one\n2. two\n\n| h1 | h2 |\n| --- | --- |\n| a | b |\n',
    )
    const docx = join(dir, 'in.docx')
    expect((await run(['create', '--type', 'docx', '--from', md, '--out', docx])).code).toBe(0)
    const r = await run(['convert', docx, '--to', 'md', '--out', join(dir, 'out.md'), '--json'])
    expect(r.code).toBe(0)
    const out = readFileSync(join(dir, 'out.md'), 'utf-8')
    expect(out).toContain('# Title')
    expect(out).toContain('**bold**')
    expect(out).toContain('*italic*')
    expect(out).toContain('[link](https://example.com)')
    expect(out).toMatch(/^- alpha$/m)
    expect(out).toMatch(/^1\. one$/m)
    expect(out).toMatch(/\| h1\s+\| h2\s+\|/)
    expect(out).toMatch(/\| a\s+\| b\s+\|/)
    expect(out.endsWith('|\n')).toBe(true)
  })

  it('turns Word equations into markdown math instead of protocol placeholders', async () => {
    const dir = tempDir()
    const html = join(dir, 'in.html')
    writeFileSync(
      html,
      '<h1>Physics</h1><p>From <formula>E = mc^2</formula> we get energy.</p><formula>\\frac{a}{b} = c</formula><p>Quote <formula>\\text{"q"}</formula> end.</p><p>Done.</p>',
    )
    const docx = join(dir, 'in.docx')
    expect((await run(['create', '--type', 'docx', '--from', html, '--out', docx])).code).toBe(0)
    const r = await run(['convert', docx, '--to', 'md', '--out', join(dir, 'out.md'), '--json'])
    expect(r.code).toBe(0)
    const out = readFileSync(join(dir, 'out.md'), 'utf-8')
    // Word's equation round trip normalizes the LaTeX (mc^2 → m{c}^{2})
    expect(out).toMatch(/\$E = m\{?c\}?\^\{?2\}?\$ we get energy/)
    expect(out).toMatch(/\$\$\s*\\frac\{a\}\{b\}\s*= c\s*\$\$/)
    expect(out).toMatch(/\$\\text\{("|&quot;|“)q("|&quot;|”)\}\$ end/)
    expect(out).not.toContain('Protected')
    expect(r.json().detail).toMatchObject({ skipped: { images: 0, fields: 0 } })
  })
})

describe('docs guard rails', () => {
  it('rejects a block index that matches nothing instead of saving a no-op', async () => {
    const dir = tempDir()
    const md = join(dir, 'a.md')
    writeFileSync(md, '# Title\n\nBody.\n')
    const docx = join(dir, 'a.docx')
    expect((await run(['create', '--type', 'docx', '--from', md, '--out', docx])).code).toBe(0)
    const before = readFileSync(docx)
    const ops = join(dir, 'ops.json')
    writeFileSync(
      ops,
      JSON.stringify([
        { op: 'findReplace', find: 'Body', replace: 'Text' },
        { op: 'deleteBlocks', target: { blockIndexes: [999] } },
      ]),
    )
    const r = await run(['docs', 'apply', docx, '--ops', ops, '--json'])
    expect(r.code).toBe(1)
    expect(r.json().message).toMatch(/op 1 \(deleteBlocks\) rejected/)
    expect(readFileSync(docx).equals(before)).toBe(true)
  })

  it('embeds local images referenced from the markdown', async () => {
    const dir = tempDir()
    const png = Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAIAAAADCAIAAADZSiLoAAAAEElEQVR4nGNgYGBgYPj//z8ABf4C/tP4tTUAAAAASUVORK5CYII=',
      'base64',
    )
    writeFileSync(join(dir, 'pic.png'), png)
    writeFileSync(join(dir, '100%.png'), png)
    const md = join(dir, 'a.md')
    writeFileSync(md, '# Title\n\n![logo](pic.png)\n\n![pct](100%.png)\n\nText.\n')
    const docx = join(dir, 'a.docx')
    expect((await run(['create', '--type', 'docx', '--from', md, '--out', docx])).code).toBe(0)
    const zip = await JSZip.loadAsync(readFileSync(docx))
    expect(Object.keys(zip.files).filter((n) => n.startsWith('word/media/'))).toHaveLength(2)
  })
})

const HELPERS = join(REPO, 'packages/docx-engine/tests/helpers/build-docx')

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
const COMMENTS_XML =
  XML_DECL +
  `<w:comments ${W_NS}>` +
  '<w:comment w:id="1" w:author="Alice" w:initials="A" w:date="2026-07-01T10:00:00Z">' +
  '<w:p w14:paraId="0A0A0A01"><w:r><w:t>Please shorten this</w:t></w:r></w:p></w:comment>' +
  '<w:comment w:id="2" w:author="Bob"><w:p w14:paraId="0A0A0A02"><w:r><w:t>Resolved already</w:t></w:r></w:p></w:comment>' +
  '</w:comments>'
const COMMENTED_P =
  '<w:p><w:r><w:t xml:space="preserve">before </w:t></w:r>' +
  '<w:commentRangeStart w:id="1"/><w:r><w:t>marked words</w:t></w:r><w:commentRangeEnd w:id="1"/>' +
  '<w:r><w:commentReference w:id="1"/></w:r><w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p>'
const SECOND_P =
  '<w:p><w:commentRangeStart w:id="2"/><w:r><w:t>second thread</w:t></w:r><w:commentRangeEnd w:id="2"/>' +
  '<w:r><w:commentReference w:id="2"/></w:r></w:p>'
const TRACKED_P =
  '<w:p><w:r><w:t xml:space="preserve">kept </w:t></w:r>' +
  '<w:ins w:id="7" w:author="Carol" w:date="2026-07-02T09:00:00Z"><w:r><w:t>added words</w:t></w:r></w:ins>' +
  '<w:del w:id="8" w:author="Carol" w:date="2026-07-02T09:01:00Z"><w:r><w:delText>gone words</w:delText></w:r></w:del></w:p>'

async function reviewFixture(dir: string): Promise<string> {
  const { buildDocx } = await import(HELPERS)
  const path = join(dir, 'review.docx')
  writeFileSync(
    path,
    await buildDocx({
      bodyXml: COMMENTED_P + SECOND_P + TRACKED_P,
      extraParts: [
        {
          path: 'word/comments.xml',
          xml: COMMENTS_XML,
          contentType:
            'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
        },
      ],
    }),
  )
  return path
}

async function zipOf(path: string): Promise<JSZip> {
  return JSZip.loadAsync(readFileSync(path))
}

describe('genoffice docs: comments, revisions, header/footer, images, charts', () => {
  it('reads comment threads and tracked changes with block indexes', async () => {
    const dir = tempDir()
    const path = await reviewFixture(dir)
    const r = await run(['docs', 'read', path, '--comments', '--revisions', '--json'])
    expect(r.code).toBe(0)
    const d = r.json().detail
    expect(d.comments).toEqual([
      expect.objectContaining({
        id: '1',
        author: 'Alice',
        text: 'Please shorten this',
        blockIndex: 0,
        anchorText: 'marked words',
      }),
      expect.objectContaining({ id: '2', author: 'Bob', blockIndex: 1 }),
    ])
    expect(d.revisions).toEqual([
      {
        id: 'r1',
        type: 'insertion',
        blockIndex: 2,
        kind: 'ins',
        author: 'Carol',
        date: expect.any(String),
        text: 'added words',
      },
      {
        id: 'r2',
        type: 'deletion',
        blockIndex: 2,
        kind: 'del',
        author: 'Carol',
        date: expect.any(String),
        text: 'gone words',
      },
    ])
  })

  it('accepts and rejects tracked changes by selector, and records --track edits', async () => {
    const dir = tempDir()
    const path = await reviewFixture(dir)
    const opsFile = join(dir, 'ops.json')

    // --author alone is not a request to track: the edit lands as plain text
    writeFileSync(
      opsFile,
      JSON.stringify([{ op: 'insert_content', html: '<p>Plain addition.</p>' }]),
    )
    const untracked = await run([
      'docs',
      'apply',
      path,
      '--ops',
      opsFile,
      '--author',
      'Robot',
      '--json',
    ])
    expect(untracked.code).toBe(0)
    expect(await documentXml(path)).not.toMatch(/w:author="Robot"/)

    // --track: the CLI's own edit becomes a pending insertion by the given author
    writeFileSync(
      opsFile,
      JSON.stringify([{ op: 'insert_content', html: '<p>Tracked addition.</p>' }]),
    )
    const tracked = await run([
      'docs',
      'apply',
      path,
      '--ops',
      opsFile,
      '--track',
      '--author',
      'Robot',
      '--json',
    ])
    expect(tracked.code).toBe(0)
    const listed = await run(['docs', 'read', path, '--revisions', '--json'])
    const revisions = listed.json().detail.revisions as Array<Record<string, unknown>>
    expect(revisions.map((r) => [r.id, r.type, r.author])).toEqual([
      ['r1', 'insertion', 'Carol'],
      ['r2', 'deletion', 'Carol'],
      ['r3', 'insertion', 'Robot'],
    ])
    expect(await documentXml(path)).toMatch(/<w:ins [^>]*w:author="Robot"/)

    // a selector that matches nothing fails the batch and names what is pending
    writeFileSync(opsFile, JSON.stringify([{ op: 'accept_changes', author: 'Nobody' }]))
    const miss = await run(['docs', 'apply', path, '--ops', opsFile, '--json'])
    expect(miss.code).toBe(1)
    expect(miss.json().message).toContain('authors Carol (2), Robot (1)')
    writeFileSync(opsFile, JSON.stringify([{ op: 'reject_changes' }]))
    const bare = await run(['docs', 'apply', path, '--ops', opsFile, '--json'])
    expect(bare.code).toBe(1)
    expect(bare.json().message).toContain('give a selector')

    // reject Robot's insertion, accept Carol's changes by id
    writeFileSync(
      opsFile,
      JSON.stringify([
        { op: 'reject_changes', author: 'Robot' },
        { op: 'accept_changes', ids: ['r1', 'r2'] },
      ]),
    )
    const applied = await run(['docs', 'apply', path, '--ops', opsFile, '--json'])
    expect(applied.code).toBe(0)
    expect(applied.json().detail.results[0].output).toContain('Rejected 1 tracked change(s) (r3)')
    expect(applied.json().detail.results[1].output).toContain('No tracked changes remain')
    const xml = await documentXml(path)
    expect(xml).not.toMatch(/<w:ins\b|<w:del\b|<w:delText/)
    expect(xml).toContain('added words')
    expect(xml).not.toContain('gone words')
    expect(xml).not.toContain('Tracked addition')
    const after = await run(['docs', 'read', path, '--revisions', '--json'])
    expect(after.json().detail.revisions).toEqual([])
  })

  it('replies to and resolves comment threads, saving comments.xml', async () => {
    const dir = tempDir()
    const path = await reviewFixture(dir)
    const ops = join(dir, 'ops.json')
    writeFileSync(
      ops,
      JSON.stringify([
        { op: 'reply_comment', parentId: '1', text: 'Shortened as requested.' },
        { op: 'resolve_comment', id: '1' },
        { op: 'read_comments' },
      ]),
    )
    const r = await run(['docs', 'apply', path, '--ops', ops, '--json'])
    expect(r.code).toBe(0)
    expect(r.json().detail.results[2].output).toContain('[resolved]')
    const zip = await zipOf(path)
    const comments = await zip.file('word/comments.xml')!.async('string')
    expect(comments).toContain('Shortened as requested.')
    expect(comments).toContain('w:author="AI Assistant"')
    const read = await run(['docs', 'read', path, '--comments', '--json'])
    const list = read.json().detail.comments
    expect(list.find((c: { id: string }) => c.id === '3')).toMatchObject({
      parentId: '1',
      text: 'Shortened as requested.',
      blockIndex: 0,
    })
    expect(list.find((c: { id: string }) => c.id === '1').done).toBe(true)

    const bad = join(dir, 'bad.json')
    writeFileSync(bad, JSON.stringify([{ op: 'reply_comment', parentId: '9', text: 'x' }]))
    const rejected = await run(['docs', 'apply', path, '--ops', bad, '--json'])
    expect(rejected.code).toBe(1)
    expect(rejected.json().message).toContain('no comment with id 9')
  })

  it('adds a comment on an exact text span and deletes threads, writing valid comment parts', async () => {
    const dir = tempDir()
    const path = await reviewFixture(dir)
    const ops = join(dir, 'ops.json')
    writeFileSync(
      ops,
      JSON.stringify([
        {
          op: 'add_comment',
          blockIndex: 2,
          text: 'added words',
          comment: 'Is this addition final?\nPlease confirm.',
        },
        {
          op: 'add_comment',
          blockIndex: 1,
          comment: 'Whole block note.',
          author: 'Reviewer',
          initials: 'RV',
        },
        { op: 'delete_comment', id: '2' },
      ]),
    )
    const r = await run(['docs', 'apply', path, '--ops', ops, '--json'])
    expect(r.code).toBe(0)
    expect(r.json().detail.results[0].output).toContain(
      'Added comment 3 on block 2, anchored to "added words"',
    )
    const zip = await zipOf(path)
    const commentsXml = await zip.file('word/comments.xml')!.async('string')
    expect(commentsXml).not.toContain('w:id="2"')
    expect(commentsXml).toMatch(
      /<w:comment w:id="3" w:author="AI Assistant" w:date="[^"]+"><w:p>.*Is this addition final\?.*<w:p w14:paraId="[0-9A-F]{8}">.*Please confirm\./,
    )
    expect(commentsXml).toContain('w:id="4" w:author="Reviewer" w:initials="RV"')
    const documentXml = await zip.file('word/document.xml')!.async('string')
    expect(documentXml).not.toMatch(/commentRangeStart w:id="2"|commentReference w:id="2"/)
    for (const id of ['1', '3', '4']) {
      expect(documentXml).toContain(`<w:commentRangeStart w:id="${id}"/>`)
      expect(documentXml).toContain(`<w:commentRangeEnd w:id="${id}"/>`)
      expect(documentXml).toContain(`<w:commentReference w:id="${id}"/>`)
    }
    expect(documentXml).toMatch(
      /commentRangeStart w:id="3"\/>.*added words.*commentRangeEnd w:id="3"\//,
    )

    const read = await run(['docs', 'read', path, '--comments', '--json'])
    const list = read.json().detail.comments as Record<string, unknown>[]
    expect(list.map((c) => c.id)).toEqual(['1', '3', '4'])
    expect(list[1]).toMatchObject({
      author: 'AI Assistant',
      blockIndex: 2,
      anchorText: 'added words',
      done: false,
    })
    expect(typeof list[1].date).toBe('string')

    const dup = join(dir, 'dup.json')
    writeFileSync(
      dup,
      JSON.stringify([{ op: 'add_comment', blockIndex: 0, text: 'e', comment: 'x' }]),
    )
    const rejected = await run(['docs', 'apply', path, '--ops', dup, '--json'])
    expect(rejected.code).toBe(1)
    expect(rejected.json().message).toContain('set occurrence')
    const withReplies = join(dir, 'root.json')
    writeFileSync(
      withReplies,
      JSON.stringify([
        { op: 'reply_comment', parentId: '1', text: 'r' },
        { op: 'delete_comment', id: '1' },
      ]),
    )
    const guarded = await run(['docs', 'apply', path, '--ops', withReplies, '--json'])
    expect(guarded.code).toBe(1)
    expect(guarded.json().message).toContain('withReplies: true')
  })

  it('a comment on an imported table anchors to its first cell paragraph and survives the save', async () => {
    const dir = tempDir()
    const path = join(dir, 'table.docx')
    writeFileSync(
      path,
      readFileSync(join(REPO, 'apps/docs/tests/pagination-corpus/docx/05-long-table.docx')),
    )
    const ops = join(dir, 'ops.json')
    writeFileSync(
      ops,
      JSON.stringify([{ op: 'add_comment', blockIndex: 2, comment: 'Table note.' }]),
    )
    const r = await run(['docs', 'apply', path, '--ops', ops, '--json'])
    expect(r.code).toBe(0)
    const documentXml = await (await zipOf(path)).file('word/document.xml')!.async('string')
    expect(documentXml.match(/<w:commentRangeStart w:id="1"\/>/g)).toHaveLength(1)
    expect(documentXml.match(/<w:commentRangeEnd w:id="1"\/>/g)).toHaveLength(1)
    expect(documentXml.match(/<w:commentReference w:id="1"\/>/g)).toHaveLength(1)
    expect(documentXml.indexOf('<w:tbl>')).toBeLessThan(
      documentXml.indexOf('<w:commentRangeStart w:id="1"/>'),
    )
    const read = await run(['docs', 'read', path, '--comments', '--json'])
    expect(read.json().detail.comments[0]).toMatchObject({ id: '1', blockIndex: 2 })
  })

  it('sets header and footer text, including the first-page variant', async () => {
    const dir = tempDir()
    const copy = join(dir, 'hf.docx')
    writeFileSync(copy, readFileSync(DOCX))
    const ops = join(dir, 'ops.json')
    writeFileSync(
      ops,
      JSON.stringify([
        { op: 'set_header_footer', kind: 'header', text: 'Quarterly Report' },
        { op: 'set_header_footer', kind: 'footer', text: 'Page {PAGE} of {NUMPAGES}' },
        { op: 'set_header_footer', kind: 'header', view: 'first', text: 'Cover' },
      ]),
    )
    const r = await run(['docs', 'apply', copy, '--ops', ops, '--json'])
    expect(r.code).toBe(0)
    const zip = await zipOf(copy)
    const documentXmlText = await zip.file('word/document.xml')!.async('string')
    expect(documentXmlText).toContain('<w:titlePg/>')
    expect(documentXmlText).toMatch(/<w:headerReference w:type="default"/)
    expect(documentXmlText).toMatch(/<w:headerReference w:type="first"/)
    expect(documentXmlText).toMatch(/<w:footerReference w:type="default"/)
    const parts = Object.keys(zip.files).filter((f) => /^word\/(header|footer)\d+\.xml$/.test(f))
    const texts = await Promise.all(parts.map((p) => zip.file(p)!.async('string')))
    expect(texts.some((t) => t.includes('Quarterly Report'))).toBe(true)
    expect(texts.some((t) => t.includes('Cover'))).toBe(true)
    expect(texts.some((t) => t.includes('PAGE') && t.includes('NUMPAGES'))).toBe(true)

    const read = await run(['docs', 'read', copy, '--header-footer', '--json'])
    expect(read.json().detail.headerFooter).toMatchObject({
      header: 'Quarterly Report',
      footer: 'Page {PAGE} of {NUMPAGES}',
      headerFirst: 'Cover',
      titlePg: true,
    })
  })

  it('changes the page setup, lists sections and inserts a section break', async () => {
    const dir = tempDir()
    const copy = join(dir, 'pages.docx')
    writeFileSync(copy, readFileSync(DOCX))
    const ops = join(dir, 'ops.json')
    writeFileSync(
      ops,
      JSON.stringify([
        {
          op: 'set_page_setup',
          paper: 'A4',
          orientation: 'landscape',
          margins: { top: '2cm', bottom: '2cm', left: '1in', right: 1440 },
          columns: { count: 2, spacing: '1cm' },
          titlePg: true,
          pageNumberStart: 5,
        },
      ]),
    )
    const r = await run(['docs', 'apply', copy, '--ops', ops, '--json'])
    expect(r.code).toBe(0)
    expect(r.json().detail.results[0].output).toContain('A4 landscape')
    let xml = await documentXml(copy)
    expect(xml).toContain('<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>')
    expect(xml).toMatch(/<w:pgMar w:top="1134" w:right="1440" w:bottom="1134" w:left="1440"/)
    expect(xml).toMatch(/<w:cols w:num="2" w:space="567"/)
    expect(xml).toContain('<w:titlePg/>')
    expect(xml).toContain('<w:pgNumType w:start="5"/>')

    const read = await run(['docs', 'read', copy, '--sections', '--json'])
    const sections = read.json().detail.sections
    expect(sections).toHaveLength(1)
    expect(sections[0]).toMatchObject({
      index: 0,
      firstBlock: 0,
      paper: 'A4',
      orientation: 'landscape',
      margins: { top: 1134, left: 1440 },
      columns: { count: 2, spacing: 567 },
      titlePg: true,
      pageNumberStart: 5,
    })

    const breakOps = join(dir, 'break.json')
    writeFileSync(
      breakOps,
      JSON.stringify([
        { op: 'insert_section_break', afterBlockIndex: 1, type: 'continuous' },
        { op: 'set_page_setup', section: 1, columns: { count: 1 }, orientation: 'portrait' },
        { op: 'set_page_setup', blockIndex: 0, margins: { top: '3cm' } },
      ]),
    )
    const b = await run(['docs', 'apply', copy, '--ops', breakOps, '--json'])
    expect(b.code).toBe(0)
    expect(b.json().detail.results[0].output).toContain('the break is block 2')
    xml = await documentXml(copy)
    const sectPrs = xml.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/g)!
    expect(sectPrs).toHaveLength(2)
    // the break paragraph carries the copy (still 2 columns, now a 3cm top margin) ...
    expect(sectPrs[0]).toMatch(/<w:cols w:num="2"/)
    expect(sectPrs[0]).toMatch(/<w:pgMar w:top="1701"/)
    expect(xml.indexOf('<w:p><w:pPr><w:sectPr')).toBeGreaterThan(0)
    // ... and the trailing sectPr starts continuous with the new settings
    expect(sectPrs[1]).toContain('<w:type w:val="continuous"/>')
    expect(sectPrs[1]).toContain('<w:pgSz w:w="11906" w:h="16838"/>')
    expect(sectPrs[1]).not.toMatch(/<w:cols w:num/)

    const after = await run(['docs', 'read', copy, '--sections', '--json'])
    const list = after.json().detail.sections
    expect(list).toHaveLength(2)
    expect(list[0]).toMatchObject({ firstBlock: 0, lastBlock: 2, columns: { count: 2 } })
    // the trailing sectPr keeps what the first batch wrote and this one did not touch
    expect(list[1]).toMatchObject({
      firstBlock: 3,
      startType: 'continuous',
      orientation: 'portrait',
      titlePg: true,
      pageNumberStart: 5,
    })

    const bad = join(dir, 'bad.json')
    writeFileSync(bad, JSON.stringify([{ op: 'set_page_setup', section: 4, paper: 'A4' }]))
    const rejected = await run(['docs', 'apply', copy, '--ops', bad, '--json'])
    expect(rejected.code).toBe(1)
    expect(rejected.json().message).toContain('section must be 0-1')
  })

  it('keeps a page setup change when the section-break paragraph itself was edited', async () => {
    const dir = tempDir()
    const { buildDocx } = await import(HELPERS)
    const path = join(dir, 'edited-break.docx')
    const sectPr =
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
    writeFileSync(
      path,
      await buildDocx({
        bodyXml:
          '<w:p><w:r><w:t>Part one</w:t></w:r></w:p>' +
          `<w:p><w:pPr>${sectPr}</w:pPr><w:r><w:t>Last line of part one</w:t></w:r></w:p>` +
          '<w:p><w:r><w:t>Part two</w:t></w:r></w:p>',
      }),
    )
    const ops = join(dir, 'ops.json')
    writeFileSync(
      ops,
      JSON.stringify([
        {
          op: 'replace_blocks',
          startBlockIndex: 1,
          endBlockIndex: 1,
          html: '<p>Part one ends here</p>',
        },
        { op: 'set_page_setup', section: 0, orientation: 'landscape' },
      ]),
    )
    const r = await run(['docs', 'apply', path, '--ops', ops, '--json'])
    expect(r.code).toBe(0)
    const xml = await documentXml(path)
    const sectPrs = xml.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/g)!
    expect(sectPrs).toHaveLength(2)
    expect(sectPrs[0]).toContain('<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>')
    expect(sectPrs[1]).not.toContain('landscape')
    const breakPara = /<w:p>(?:(?!<\/w:p>)[\s\S])*?landscape[\s\S]*?<\/w:p>/.exec(xml)![0]
    expect(breakPara).toContain('Part one ends here')
  })

  it('keeps a middle section landscape when its break paragraph is retyped in the same batch', async () => {
    const dir = tempDir()
    const { buildDocx } = await import(HELPERS)
    const path = join(dir, 'middle-break.docx')
    const sectPr =
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
    writeFileSync(
      path,
      await buildDocx({
        bodyXml:
          `<w:p><w:pPr>${sectPr}</w:pPr><w:r><w:t>Part one</w:t></w:r></w:p>` +
          '<w:p><w:r><w:t>Part two</w:t></w:r></w:p>' +
          `<w:p><w:pPr>${sectPr}</w:pPr><w:r><w:t>Last line of part two</w:t></w:r></w:p>` +
          '<w:p><w:r><w:t>Part three</w:t></w:r></w:p>',
      }),
    )
    const ops = join(dir, 'ops.json')
    writeFileSync(
      ops,
      JSON.stringify([
        { op: 'set_page_setup', section: 1, orientation: 'landscape' },
        {
          op: 'findReplace',
          find: 'Last line of part two',
          replace: 'Last line of part two, retyped',
          matchCase: true,
        },
      ]),
    )
    const r = await run(['docs', 'apply', path, '--ops', ops, '--json'])
    expect(r.code).toBe(0)
    const xml = await documentXml(path)
    const sectPrs = xml.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/g)!
    expect(sectPrs).toHaveLength(3)
    expect(sectPrs[0]).not.toContain('landscape')
    expect(sectPrs[1]).toContain('<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>')
    expect(sectPrs[2]).not.toContain('landscape')
    const breakPara = /<w:p>(?:(?!<\/w:p>)[\s\S])*?landscape[\s\S]*?<\/w:p>/.exec(xml)![0]
    expect(breakPara).toContain('Last line of part two, retyped')
    const read = await run(['docs', 'read', path, '--sections', '--json'])
    expect(read.json().detail.sections.map((s: { orientation: string }) => s.orientation)).toEqual([
      'portrait',
      'landscape',
      'portrait',
    ])
  })

  it('inserts a local image and a chart, then edits the chart data', async () => {
    const dir = tempDir()
    const { buildChartDocx, TINY_PNG_BASE64 } = await import(HELPERS)
    const path = join(dir, 'media.docx')
    writeFileSync(path, await buildChartDocx('<w:p><w:r><w:t>Intro paragraph.</w:t></w:r></w:p>'))
    const png = join(dir, 'dot.png')
    writeFileSync(png, Buffer.from(TINY_PNG_BASE64, 'base64'))
    const ops = join(dir, 'ops.json')
    writeFileSync(
      ops,
      JSON.stringify([
        { op: 'insert_image', url: 'dot.png', afterBlockIndex: 0, maxWidthPx: 200 },
        {
          op: 'insert_chart',
          kind: 'pie',
          title: 'Share',
          categories: ['A', 'B'],
          series: [{ name: 'S', values: [60, 40] }],
        },
        {
          op: 'edit_chart',
          blockIndex: 2,
          title: 'Renamed',
          series: [{ index: 0, values: [1, 2, 3] }],
        },
      ]),
    )
    const r = await run(['docs', 'apply', path, '--ops', ops, '--json'])
    expect(r.code).toBe(0)
    expect(r.json().detail.results[0].output).toContain('Inserted the image (1x1px)')
    const zip = await zipOf(path)
    expect(Object.keys(zip.files).some((f) => /^word\/media\/.*\.png$/.test(f))).toBe(true)
    const charts = Object.keys(zip.files).filter((f) => /^word\/charts\/chart\d+\.xml$/.test(f))
    expect(charts).toHaveLength(2)
    const chartXml = await Promise.all(charts.map((c) => zip.file(c)!.async('string')))
    expect(chartXml.some((x) => x.includes('Renamed') && x.includes('<c:v>3</c:v>'))).toBe(true)
    expect(chartXml.some((x) => x.includes('pieChart') && x.includes('Share'))).toBe(true)

    const read = await run(['docs', 'read', path, '--json'])
    const items = read.json().detail.items
    expect(items[1]).toMatchObject({ type: 'protected', kind: 'image' })
    expect(items[2]).toMatchObject({ type: 'protected', kind: 'chart' })
    expect(items[3]).toMatchObject({ type: 'protected', kind: 'chart' })

    const missing = join(dir, 'missing.json')
    writeFileSync(missing, JSON.stringify([{ op: 'insert_image', url: 'nope.png' }]))
    const bad = await run(['docs', 'apply', path, '--ops', missing, '--json'])
    expect(bad.code).toBe(1)
    expect(bad.json().message).toContain('image not found')
  })
  it('sets a picture watermark from a local file, reports it on read, and swaps it back to text', async () => {
    const dir = tempDir()
    const { TINY_PNG_BASE64 } = await import(HELPERS)
    const path = join(dir, 'picwm.docx')
    writeFileSync(path, readFileSync(DOCX))
    writeFileSync(join(dir, 'logo.png'), Buffer.from(TINY_PNG_BASE64, 'base64'))

    const ops = join(dir, 'ops.json')
    writeFileSync(ops, JSON.stringify([{ op: 'set_watermark', image: 'logo.png', scale: 200 }]))
    const r = await run(['docs', 'apply', path, '--ops', ops, '--json'])
    expect(r.code).toBe(0)
    const results = r.json().detail.results as Array<{ output: string }>
    expect(results[0].output).toContain('Picture watermark set (1x1px source)')

    const zip = await zipOf(path)
    const hdrName = Object.keys(zip.files).find((n) => /^word\/header\d+\.xml$/.test(n))!
    const hdr = await zip.file(hdrName)!.async('string')
    expect(hdr).toContain('<w:docPartGallery w:val="Watermarks"/>')
    expect(hdr).toContain('<v:shape id="WordPictureWatermark1"')
    // 1x1px at 200% = 1.5pt
    expect(hdr).toContain('width:1.5pt;height:1.5pt;')
    expect(hdr).toContain('gain="19661f"')
    const rId = /<v:imagedata r:id="([^"]+)"/.exec(hdr)![1]
    const rels = await zip
      .file(hdrName.replace(/^word\/(.*)$/, 'word/_rels/$1.rels'))!
      .async('string')
    const target = new RegExp(
      `<Relationship Id="${rId}" Type="[^"]*/image" Target="([^"]+)"/>`,
    ).exec(rels)![1]
    expect(zip.file(`word/${target}`)).not.toBeNull()

    const read = await run(['docs', 'read', path, '--header-footer', '--json'])
    expect(read.code).toBe(0)
    expect(read.json().detail.headerFooter.watermark).toEqual({
      kind: 'picture',
      text: null,
      widthPt: 1.5,
      heightPt: 1.5,
      washout: true,
    })

    writeFileSync(ops, JSON.stringify([{ op: 'set_watermark', text: 'DRAFT' }]))
    expect((await run(['docs', 'apply', path, '--ops', ops, '--json'])).code).toBe(0)
    const zip2 = await zipOf(path)
    const hdr2 = await zip2.file(hdrName)!.async('string')
    expect(hdr2).not.toContain('<v:imagedata')
    expect(hdr2).toContain('string="DRAFT"')
    const read2 = await run(['docs', 'read', path, '--header-footer', '--json'])
    expect(read2.json().detail.headerFooter.watermark).toEqual({ kind: 'text', text: 'DRAFT' })

    writeFileSync(ops, JSON.stringify([{ op: 'set_watermark', image: 'nope.png' }]))
    const bad = await run(['docs', 'apply', path, '--ops', ops, '--json'])
    expect(bad.code).toBe(1)
    expect(bad.json().message).toContain('image not found')
  })
  it('defines and applies styles, sets a watermark, floats a text box and a picture', async () => {
    const dir = tempDir()
    const { TINY_PNG_BASE64 } = await import(HELPERS)
    const path = join(dir, 'styled.docx')
    writeFileSync(path, readFileSync(DOCX))
    const png = join(dir, 'logo.png')
    writeFileSync(png, Buffer.from(TINY_PNG_BASE64, 'base64'))

    const styles = await run(['docs', 'read', path, '--styles', '--json'])
    expect(styles.code).toBe(0)
    const list = styles.json().detail.styles as Array<Record<string, unknown>>
    const h1 = list.find((s) => s.styleId === 'Heading1')
    expect(h1).toMatchObject({ type: 'paragraph', headingLevel: 1 })
    expect(Number(h1!.inUse)).toBeGreaterThan(0)

    const ops = join(dir, 'ops.json')
    writeFileSync(
      ops,
      JSON.stringify([
        {
          op: 'define_style',
          styleId: 'Callout',
          name: 'Callout',
          basedOn: 'Normal',
          paragraph: { indentLeft: '1cm', spaceAfter: 6, align: 'justify' },
          run: { italic: true, color: '#1A73E8' },
        },
        {
          op: 'define_style',
          styleId: 'Heading1',
          run: { color: '#C00000' },
          paragraph: { spaceBefore: 24 },
        },
        { op: 'applyStyle', target: { blockIndexes: [1] }, styleId: 'Callout' },
        { op: 'applyStyle', target: { blockIndexes: [2] }, styleId: 'Heading2' },
        { op: 'define_style', styleId: 'SubTitle', basedOn: 'Heading1', run: { italic: true } },
        { op: 'applyStyle', target: { blockIndexes: [3] }, styleId: 'SubTitle' },
        { op: 'setFont', target: { blockIndexes: [3], headingLevel: 1 }, bold: true },
        { op: 'set_watermark', text: 'DRAFT', color: '#FF0000', diagonal: false },
        {
          op: 'insert_text_box',
          afterBlockIndex: 0,
          text: 'Side note',
          width: '5cm',
          height: '2cm',
          x: 0,
          y: '1cm',
          fill: '#FFF2CC',
        },
        {
          op: 'insert_picture',
          url: 'logo.png',
          afterBlockIndex: 0,
          width: '1in',
          float: { anchor: 'page', x: '1cm', y: '1cm', wrap: 'behind' },
          altText: 'Company logo',
        },
      ]),
    )
    const r = await run(['docs', 'apply', path, '--ops', ops, '--json'])
    expect(r.code).toBe(0)
    const results = r.json().detail.results as Array<{ op: string; output: string }>
    expect(results[0].output).toContain('Created style Callout')
    expect(results[1].output).toContain('Updated style Heading1')
    expect(results[9].output).toContain('floating picture (96x96px)')

    const zip = await zipOf(path)
    const stylesXml = await zip.file('word/styles.xml')!.async('string')
    expect(stylesXml).toMatch(
      /<w:style w:type="paragraph" w:styleId="Callout" w:customStyle="1"><w:name w:val="Callout"\/><w:basedOn w:val="Normal"\/><w:qFormat\/><w:pPr><w:spacing w:after="120"\/><w:ind w:left="567"\/><w:jc w:val="both"\/><\/w:pPr><w:rPr><w:i\/><w:iCs\/><w:color w:val="1A73E8"\/><\/w:rPr><\/w:style>/,
    )
    const heading1 = /<w:style [^>]*w:styleId="Heading1"[\s\S]*?<\/w:style>/.exec(stylesXml)![0]
    expect(heading1).toContain('<w:color w:val="C00000"/>')
    expect(heading1).toContain('w:before="480"')
    expect(heading1).toContain('<w:name w:val="heading 1"/>')
    expect(stylesXml.match(/w:styleId="Heading1"/g)).toHaveLength(1)

    const body = await zip.file('word/document.xml')!.async('string')
    expect(body).toContain('<w:pStyle w:val="Callout"/>')
    expect(body).toContain('<w:pStyle w:val="Heading2"/>')
    expect(body).toMatch(/<w:pStyle w:val="SubTitle"\/>.*?<w:b\/>/s)
    expect(body).toContain('descr="Company logo"')
    expect(body).toMatch(
      /<wp:anchor[^>]*behindDoc="1"[^>]*>.*?<wp:positionH relativeFrom="page"><wp:posOffset>360000<\/wp:posOffset>/s,
    )
    expect(body).toContain('<wps:txbx><w:txbxContent>')
    expect(body).toContain('<a:srgbClr val="FFF2CC"/>')

    const header = Object.keys(zip.files).find((f) => /^word\/header\d*\.xml$/.test(f))!
    const headerXml = await zip.file(header)!.async('string')
    expect(headerXml).toContain('<w:docPartGallery w:val="Watermarks"/>')
    expect(headerXml).toContain('string="DRAFT"')
    expect(headerXml).toContain('fillcolor="#FF0000"')

    const read = await run(['docs', 'read', path, '--styles', '--json'])
    const after = read.json().detail
    expect(after.items[1]).toMatchObject({ type: 'protected', kind: 'image' })
    expect(after.items[2]).toMatchObject({ type: 'protected', kind: 'passthrough' })
    expect(
      (after.styles as Array<Record<string, unknown>>).find((s) => s.styleId === 'Callout'),
    ).toMatchObject({
      basedOn: 'Normal',
      inUse: 1,
    })

    const bad = join(dir, 'bad.json')
    writeFileSync(
      bad,
      JSON.stringify([{ op: 'applyStyle', target: { blockIndexes: [0] }, styleId: 'Nope' }]),
    )
    const rejected = await run(['docs', 'apply', path, '--ops', bad, '--json'])
    expect(rejected.code).toBe(1)
    expect(rejected.json().message).toContain('no style "Nope"')
  })
})
