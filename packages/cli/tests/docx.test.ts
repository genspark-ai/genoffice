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

  it('read --full returns whole block text instead of the 200-character preview', async () => {
    const dir = tempDir()
    const long = 'lorem '.repeat(80).trim()
    writeFileSync(join(dir, 'long.md'), `# Title\n\n${long}\n`)
    const out = join(dir, 'long.docx')
    expect(
      (await run(['create', '--type', 'docx', '--from', join(dir, 'long.md'), '--out', out])).code,
    ).toBe(0)
    const preview = (await run(['docs', 'read', out, '--json'])).json().detail.items[1].text
    expect(preview).toHaveLength(200)
    expect(preview.endsWith('...')).toBe(true)
    const full = (await run(['docs', 'read', out, '--full', '--json'])).json().detail.items[1].text
    expect(full).toBe(long)
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
        blockIndex: 2,
        kind: 'ins',
        author: 'Carol',
        date: expect.any(String),
        text: 'added words',
      },
      { blockIndex: 2, kind: 'del', author: 'Carol', date: expect.any(String), text: 'gone words' },
    ])
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
})
