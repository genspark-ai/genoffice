import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { mergePPrFormat, parseDocx, saveDocx, setPPrChange, type SaveBlock } from '../src/index'
import { buildDocx, TABLE_XML } from './helpers/build-docx'

/** paragraph with a tracked insertion and a tracked deletion */
const REVISED_P =
  '<w:p><w:r><w:t xml:space="preserve">保留 </w:t></w:r>' +
  '<w:ins w:id="11" w:author="Alice" w:date="2026-07-01T10:00:00Z">' +
  '<w:r><w:t>新增文字</w:t></w:r>' +
  '</w:ins>' +
  '<w:del w:id="12" w:author="Bob" w:date="2026-07-02T08:30:00Z">' +
  '<w:r><w:delText xml:space="preserve">被删文字</w:delText></w:r>' +
  '</w:del>' +
  '<w:r><w:t xml:space="preserve"> 结尾</w:t></w:r></w:p>'

/** an insertion that was later deleted: w:del nested in w:ins */
const NESTED_P =
  '<w:p><w:ins w:id="21" w:author="Alice">' +
  '<w:del w:id="22" w:author="Bob">' +
  '<w:r><w:delText>插了又删</w:delText></w:r>' +
  '</w:del></w:ins></w:p>'

/** paragraph-mark revision lives in pPr — editable via rawPPr passthrough */
const PARA_MARK_P =
  '<w:p><w:pPr><w:rPr><w:del w:id="31" w:author="Alice"/></w:rPr></w:pPr>' +
  '<w:r><w:t>段落标记被删</w:t></w:r></w:p>'

/** tracked paragraph-property change (pPrChange nests the old pPr) */
const PPR_CHANGE_P =
  '<w:p><w:pPr><w:jc w:val="center"/><w:keepNext/>' +
  '<w:pPrChange w:id="51" w:author="Alice" w:date="2026-07-01T10:00:00Z">' +
  '<w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr>' +
  '<w:r><w:t>对齐方式被改过</w:t></w:r></w:p>'

/** move-from revision: content was moved away (shown with del marks) */
const MOVE_P =
  '<w:p><w:moveFrom w:id="41" w:author="Alice" w:name="move1">' +
  '<w:r><w:delText>moved away</w:delText></w:r></w:moveFrom></w:p>'

/** move-to revision: content was moved here (shown with ins marks) */
const MOVE_TO_P =
  '<w:p><w:moveTo w:id="42" w:author="Bob" w:name="move2">' +
  '<w:r><w:t>moved here</w:t></w:r></w:moveTo></w:p>'

describe('tracked change parsing', () => {
  it('parses run-level w:ins / w:del into editable runs', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: REVISED_P }))
    const p = doc.blocks[0]
    expect(p.type).toBe('paragraph')
    expect(p.runs).toEqual([
      { text: '保留 ' },
      {
        text: '新增文字',
        ins: { author: 'Alice', date: '2026-07-01T10:00:00Z', id: '11' },
      },
      {
        text: '被删文字',
        del: { author: 'Bob', date: '2026-07-02T08:30:00Z', id: '12' },
      },
      { text: ' 结尾' },
    ])
  })

  it('parses an insertion later deleted with both markers', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: NESTED_P }))
    const p = doc.blocks[0]
    expect(p.type).toBe('paragraph')
    expect(p.runs).toEqual([
      {
        text: '插了又删',
        ins: { author: 'Alice', id: '21' },
        del: { author: 'Bob', id: '22' },
      },
    ])
  })

  it('paragraph-mark / pPrChange revisions parse as editable paragraphs with rawPPr', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: PARA_MARK_P + PPR_CHANGE_P }))
    expect(doc.blocks[0].type).toBe('paragraph')
    expect(doc.blocks[0].rawPPr).toContain('<w:del w:id="31"')
    expect(doc.blocks[1].type).toBe('paragraph')
    expect(doc.blocks[1].rawPPr).toContain('<w:pPrChange')
    expect(doc.blocks[1].format?.align).toBe('center')
    // pPrChangeInfo is extracted from pPrChange for UI badge/navigation
    expect(doc.blocks[1].pPrChangeInfo).toBeDefined()
    expect(doc.blocks[1].pPrChangeInfo!.author).toBe('Alice')
    expect(doc.blocks[1].pPrChangeInfo!.id).toBe('51')
  })

  it('move revisions parse as editable paragraphs with moveRevision marker', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: MOVE_P }))
    expect(doc.blocks[0].type).toBe('paragraph')
    expect(doc.blocks[0].moveRevision).toBe('from')
    // The run text is extracted from w:delText inside w:moveFrom with del mark
    const delRun = doc.blocks[0].runs?.find((r) => r.del !== undefined)
    expect(delRun).toBeDefined()
  })

  it('moveTo revision parses as editable with moveRevision "to" and ins mark', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: MOVE_TO_P }))
    expect(doc.blocks[0].type).toBe('paragraph')
    expect(doc.blocks[0].moveRevision).toBe('to')
    const insRun = doc.blocks[0].runs?.find((r) => r.ins !== undefined)
    expect(insRun).toBeDefined()
    expect(insRun!.text).toBe('moved here')
  })

  it('a text edit regenerates with the pPr (incl. pPrChange) byte-identical', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: PPR_CHANGE_P }))
    const block = parsed.blocks[0]
    expect(block.rawPPr).toBeDefined()
    const out = await saveDocx(parsed, [
      {
        kind: 'generated',
        block: { type: 'paragraph', rawPPr: block.rawPPr, runs: [{ text: '改过的文字' }] },
      },
    ])
    const zip = await JSZip.loadAsync(out)
    const xml = await zip.file('word/document.xml')!.async('string')
    expect(xml).toContain('改过的文字')
    expect(xml).toContain(block.rawPPr!)
    expect(xml).toContain('<w:keepNext/>')
    expect(xml).toContain('<w:pPrChange w:id="51"')
  })
})

describe('tracked change save fidelity', () => {
  it('a paragraph revision rides the paragraph mark, a table one keeps the wrapper', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>base</w:t></w:r></w:p>' }),
    )
    const revision = {
      kind: 'ins' as const,
      author: 'Structural Reviewer',
      date: '2026-07-30T10:00:00Z',
    }
    // a whole paragraph takes the shape Word's Compare writes: the mark
    // carries the record, so readers resolve it (a w:ins wrapper around w:p
    // is the shape for tables)
    // the caller marks the runs (compare and --track both do) and asks for
    // the block revision on top
    const paragraph = await saveDocx(parsed, [
      {
        kind: 'generated',
        block: {
          type: 'paragraph',
          runs: [{ text: 'inserted paragraph', ins: { ...revision, id: '9002' } }],
        },
        revision: { ...revision, id: '9001' },
      },
    ])
    const reopenedParagraph = await parseDocx(paragraph)
    expect(reopenedParagraph.blocks[0].blockRevision).toBeUndefined()
    expect(reopenedParagraph.blocks[0].paraMarkIns).toMatchObject({
      author: revision.author,
      date: revision.date,
    })
    expect(reopenedParagraph.blocks[0].runs?.[0].text).toBe('inserted paragraph')
    expect(reopenedParagraph.blocks[0].runs?.[0].ins).toMatchObject({
      author: revision.author,
    })

    const table = await saveDocx(parsed, [{ kind: 'xml', xml: TABLE_XML, revision }])
    const reopenedTable = await parseDocx(table)
    expect(reopenedTable.blocks[0].type).toBe('table')
    expect(reopenedTable.blocks[0].blockRevision).toMatchObject(revision)
  })

  it('re-emits w:ins / w:del wrappers when the paragraph is regenerated', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: REVISED_P }))
    const block = parsed.blocks[0]
    if (block.type !== 'paragraph') throw new Error('expected paragraph')
    // edit the plain leading run, keep the revised runs untouched
    const runs = block.runs!.map((r, i) => (i === 0 ? { ...r, text: '改过 ' } : r))
    const saveBlocks: SaveBlock[] = [{ kind: 'generated', block: { type: 'paragraph', runs } }]
    const out = await saveDocx(parsed, saveBlocks)

    const zip = await JSZip.loadAsync(out)
    const xml = await zip.file('word/document.xml')!.async('string')
    expect(xml).toContain('改过 ')
    expect(xml).toContain('<w:ins w:id="11" w:author="Alice" w:date="2026-07-01T10:00:00Z">')
    expect(xml).toContain('<w:del w:id="12" w:author="Bob" w:date="2026-07-02T08:30:00Z">')
    expect(xml).toContain('<w:delText xml:space="preserve">被删文字</w:delText>')

    // round-trip: parse the saved file again
    const again = await parseDocx(out)
    expect(again.blocks[0].runs).toEqual([
      { text: '改过 ' },
      { text: '新增文字', ins: { author: 'Alice', date: '2026-07-01T10:00:00Z', id: '11' } },
      { text: '被删文字', del: { author: 'Bob', date: '2026-07-02T08:30:00Z', id: '12' } },
      { text: ' 结尾' },
    ])
  })

  it('mints ids for new revisions without one', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>base</w:t></w:r></w:p>' }),
    )
    const saveBlocks: SaveBlock[] = [
      {
        kind: 'generated',
        block: {
          type: 'paragraph',
          runs: [
            { text: 'base' },
            { text: '新输入', ins: { author: '用户', date: '2026-07-22T04:00:00Z' } },
            { text: '被删掉', del: { author: '用户', date: '2026-07-22T04:00:00Z' } },
          ],
        },
      },
    ]
    const out = await saveDocx(parsed, saveBlocks)
    const zip = await JSZip.loadAsync(out)
    const xml = await zip.file('word/document.xml')!.async('string')
    expect(xml).toMatch(/<w:ins w:id="\d+" w:author="用户" w:date="2026-07-22T04:00:00Z">/)
    expect(xml).toMatch(/<w:del w:id="\d+" w:author="用户" w:date="2026-07-22T04:00:00Z">/)
    expect(xml).toContain('<w:delText xml:space="preserve">被删掉</w:delText>')
  })

  it('leaves untouched revised paragraphs byte-identical', async () => {
    const bytes = await buildDocx({ bodyXml: REVISED_P })
    const parsed = await parseDocx(bytes)
    const saveBlocks: SaveBlock[] = parsed.blocks
      .filter((b) => !b.hidden)
      .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))
    const out = await saveDocx(parsed, saveBlocks)
    expect(out).toBe(bytes)
  })
})

describe('comment writing', () => {
  const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
  const COMMENTS_XML =
    XML_DECL +
    '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:comment w:id="1" w:author="Alice"><w:p><w:r><w:t>已有批注</w:t></w:r></w:p></w:comment>' +
    '</w:comments>'

  it('creates word/comments.xml with rel + content type when absent', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>hello</w:t></w:r></w:p>' }),
    )
    const block = parsed.blocks[0]
    if (block.type !== 'paragraph') throw new Error('expected paragraph')
    const runs = [{ ...block.runs![0], commentIds: ['1'] }]
    const out = await saveDocx(
      parsed,
      [{ kind: 'generated', block: { type: 'paragraph', runs } }],
      {
        comments: [
          { id: '1', author: '用户', date: '2026-07-22T04:00:00Z', text: '第一条\n第二行' },
        ],
      },
    )

    const zip = await JSZip.loadAsync(out)
    const comments = await zip.file('word/comments.xml')!.async('string')
    expect(comments).toContain('w:id="1"')
    expect(comments).toContain('w:author="用户"')
    expect(comments).toContain('<w:t xml:space="preserve">第一条</w:t>')
    expect(comments).toContain('<w:t xml:space="preserve">第二行</w:t>')

    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
    expect(rels).toContain('relationships/comments')
    expect(rels).toContain('Target="comments.xml"')

    const types = await zip.file('[Content_Types].xml')!.async('string')
    expect(types).toContain('PartName="/word/comments.xml"')

    const docXml = await zip.file('word/document.xml')!.async('string')
    expect(docXml).toContain('<w:commentRangeStart w:id="1"/>')
    expect(docXml).toContain('<w:commentReference w:id="1"/>')

    // the saved file parses back with the comment attached
    const again = await parseDocx(out)
    expect(again.comments).toHaveLength(1)
    expect(again.comments[0].text).toBe('第一条\n第二行')
    expect(again.blocks[0].runs![0].commentIds).toEqual(['1'])
  })

  it('rewrites an existing comments.xml from the full list', async () => {
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>hello</w:t></w:r></w:p>',
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
    const out = await saveDocx(
      parsed,
      parsed.blocks
        .filter((b) => !b.hidden)
        .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! })),
      {
        comments: [
          { id: '1', author: 'Alice', text: '已有批注' },
          { id: '2', author: '用户', text: '新加的' },
        ],
      },
    )
    const zip = await JSZip.loadAsync(out)
    const comments = await zip.file('word/comments.xml')!.async('string')
    expect(comments).toContain('w:id="1"')
    expect(comments).toContain('w:id="2"')
    expect(comments).toContain('新加的')
  })
})

describe('document protection', () => {
  const SETTINGS_XML =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:zoom w:percent="100"/></w:settings>'

  const withSettings = () =>
    buildDocx({
      bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
      extraParts: [
        {
          path: 'word/settings.xml',
          xml: SETTINGS_XML,
          contentType:
            'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
        },
      ],
    })

  it('reports no protection for an unprotected document', async () => {
    const doc = await parseDocx(await withSettings())
    expect(doc.protection).toBeNull()
  })

  it('writes and re-reads w:documentProtection', async () => {
    const parsed = await parseDocx(await withSettings())
    const blocks: SaveBlock[] = parsed.blocks
      .filter((b) => !b.hidden)
      .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))
    const out = await saveDocx(parsed, blocks, {
      protection: { edit: 'readOnly', enforced: true },
    })
    const again = await parseDocx(out)
    expect(again.protection).toEqual({ edit: 'readOnly', enforced: true })

    // removing it round-trips too
    const cleared = await saveDocx(again, blocks, { protection: null })
    expect((await parseDocx(cleared)).protection).toBeNull()
  })
})

describe('mergePPrFormat', () => {
  const RAW =
    '<w:pPr><w:pStyle w:val="Body"/><w:keepNext/><w:tabs><w:tab w:val="left" w:pos="420"/></w:tabs>' +
    '<w:spacing w:before="120" w:beforeAutospacing="1"/><w:jc w:val="left"/>' +
    '<w:rPr><w:del w:id="31" w:author="Alice"/></w:rPr>' +
    '<w:pPrChange w:id="51" w:author="Alice"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr>'

  it('replaces managed children in place, keeps everything else byte-identical', () => {
    const merged = mergePPrFormat(RAW, { align: 'center', spaceBefore: 240 })
    expect(merged).toContain('<w:pStyle w:val="Body"/>')
    expect(merged).toContain('<w:keepNext/>')
    expect(merged).toContain('<w:tabs><w:tab w:val="left" w:pos="420"/></w:tabs>')
    expect(merged).toContain('<w:rPr><w:del w:id="31" w:author="Alice"/></w:rPr>')
    expect(merged).toContain('<w:pPrChange w:id="51"')
    // rebuilt from the format model (old spacing/jc replaced)
    expect(merged).toContain('<w:spacing w:before="240"/>')
    expect(merged).toContain('<w:jc w:val="center"/>')
    expect(merged).not.toContain('beforeAutospacing')
    // schema order: tabs < spacing < jc < rPr
    const order = [
      '<w:keepNext/>',
      '<w:tabs>',
      '<w:spacing',
      '<w:jc',
      '<w:rPr>',
      '<w:pPrChange',
    ].map((t) => merged.indexOf(t))
    expect([...order].sort((a, b) => a - b)).toEqual(order)
  })

  it('clearing all format keeps only unmanaged children', () => {
    const merged = mergePPrFormat(RAW, undefined)
    expect(merged).toContain('<w:keepNext/>')
    expect(merged).not.toContain('<w:spacing')
    // the only remaining jc is the old value recorded inside pPrChange
    expect(merged.indexOf('<w:jc')).toBeGreaterThan(merged.indexOf('<w:pPrChange'))
  })
})

describe('stripPPrChange', () => {
  it('removes pPrChange from a raw pPr slice', async () => {
    const { stripPPrChange } = await import('../src/generate')
    const raw =
      '<w:pPr><w:jc w:val="center"/>' +
      '<w:pPrChange w:id="51" w:author="Alice">' +
      '<w:pPr><w:jc w:val="left"/></w:pPr>' +
      '</w:pPrChange></w:pPr>'
    const result = stripPPrChange(raw)
    expect(result).toContain('<w:jc w:val="center"/>')
    expect(result).not.toContain('<w:pPrChange')
    expect(result).not.toContain('Alice')
  })

  it('returns unchanged string if no pPrChange', async () => {
    const { stripPPrChange } = await import('../src/generate')
    const raw = '<w:pPr><w:jc w:val="left"/></w:pPr>'
    expect(stripPPrChange(raw)).toBe(raw)
  })
})

describe('setPPrChange', () => {
  it('writes a schema-ordered paragraph revision that parses after reopen', async () => {
    const raw = '<w:pPr><w:spacing w:after="240"/><w:jc w:val="center"/></w:pPr>'
    const revised = setPPrChange(
      raw,
      JSON.stringify({
        author: 'Paragraph Reviewer',
        date: '2026-07-30T09:00:00Z',
        old: {
          type: 'docParagraph',
          styleId: 'BodyText',
          format: {
            align: 'left',
            indentLeft: 120,
            lineSpacing: 1,
            spaceAfter: 80,
          },
        },
      }),
    )
    expect(revised.indexOf('<w:pPrChange')).toBeGreaterThan(revised.indexOf('<w:jc'))
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: `<w:p>${revised}<w:r><w:t>paragraph</w:t></w:r></w:p>` }),
    )
    expect(parsed.blocks[0].pPrChangeInfo).toMatchObject({
      author: 'Paragraph Reviewer',
      old: {
        align: 'left',
        indentLeft: 120,
        lineSpacing: 1,
        spaceAfter: 80,
      },
    })
  })
})

describe('paired moves (w:moveFrom / w:moveTo)', () => {
  const MOVE_PAIR_XML =
    '<w:p><w:r><w:t xml:space="preserve">A </w:t></w:r>' +
    '<w:moveFromRangeStart w:id="60" w:author="Alice" w:date="2026-07-01T10:00:00Z" w:name="move77"/>' +
    '<w:moveFrom w:id="61" w:author="Alice" w:date="2026-07-01T10:00:00Z">' +
    '<w:r><w:delText>gone</w:delText></w:r></w:moveFrom>' +
    '<w:moveFromRangeEnd w:id="60"/></w:p>' +
    '<w:p><w:r><w:t xml:space="preserve">B </w:t></w:r>' +
    '<w:moveToRangeStart w:id="62" w:author="Alice" w:date="2026-07-01T10:00:00Z" w:name="move77"/>' +
    '<w:moveTo w:id="63" w:author="Alice" w:date="2026-07-01T10:00:00Z">' +
    '<w:r><w:t>gone</w:t></w:r></w:moveTo>' +
    '<w:moveToRangeEnd w:id="62"/></w:p>'

  it('parses the range name onto both sides', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: MOVE_PAIR_XML }))
    const from = doc.blocks[0].runs!.find((r) => r.del)!
    const to = doc.blocks[1].runs!.find((r) => r.ins)!
    expect(from.del).toEqual({
      author: 'Alice',
      date: '2026-07-01T10:00:00Z',
      id: '61',
      move: 'move77',
    })
    expect(to.ins!.move).toBe('move77')
  })

  it('a range opened in one paragraph still names a move in the next', async () => {
    const xml =
      '<w:p><w:moveFromRangeStart w:id="1" w:author="A" w:name="mv"/>' +
      '<w:moveFrom w:id="2" w:author="A"><w:r><w:delText>one</w:delText></w:r></w:moveFrom></w:p>' +
      '<w:p><w:moveFrom w:id="3" w:author="A"><w:r><w:delText>two</w:delText></w:r></w:moveFrom>' +
      '<w:moveFromRangeEnd w:id="1"/></w:p>' +
      '<w:p><w:moveFrom w:id="4" w:author="A"><w:r><w:delText>loose</w:delText></w:r></w:moveFrom></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: xml }))
    expect(doc.blocks[0].runs![0].del!.move).toBe('mv')
    expect(doc.blocks[1].runs![0].del!.move).toBe('mv')
    expect(doc.blocks[2].runs![0].del!.move).toBe('move4')
  })

  it('saves editor moves as named ranges that pair again on reopen', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>' }),
    )
    const rev = { author: 'Alice', date: '2026-07-01T10:00:00Z', move: 'move12345' }
    const out = await saveDocx(parsed, [
      {
        kind: 'generated',
        block: {
          type: 'paragraph',
          runs: [
            { text: 'A ' },
            { text: 'go', del: rev },
            { text: 'ne', del: rev },
            { text: ' B' },
          ],
        },
      },
      {
        kind: 'generated',
        block: { type: 'paragraph', runs: [{ text: 'C ' }, { text: 'gone', ins: rev }] },
      },
    ])
    const xml = await (await JSZip.loadAsync(out)).file('word/document.xml')!.async('string')
    const from =
      /<w:moveFromRangeStart w:id="(\d+)" w:author="Alice" w:date="2026-07-01T10:00:00Z" w:name="move12345"\/><w:moveFrom w:id="(\d+)"[^>]*><w:r><w:delText[^>]*>go<\/w:delText><\/w:r><w:r><w:delText[^>]*>ne<\/w:delText><\/w:r><\/w:moveFrom><w:moveFromRangeEnd w:id="\1"\/>/.exec(
        xml,
      )
    expect(from).not.toBeNull()
    expect(from![1]).not.toBe(from![2])
    expect(xml).toMatch(
      /<w:moveToRangeStart w:id="(\d+)"[^>]*w:name="move12345"\/><w:moveTo w:id="\d+"[^>]*><w:r><w:t[^>]*>gone<\/w:t><\/w:r><\/w:moveTo><w:moveToRangeEnd w:id="\1"\/>/,
    )
    expect(xml).not.toMatch(/<w:(?:ins|del)\b/)
    const back = await parseDocx(out)
    expect(back.blocks[0].runs!.find((r) => r.del)!.del!.move).toBe('move12345')
    expect(back.blocks[1].runs!.find((r) => r.ins)!.ins!.move).toBe('move12345')
  })
})

describe('rPrChange keeps the recorded previous rPr verbatim', () => {
  const OLD_RPR =
    '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:i/><w:color w:val="0000FF"/>' +
    '<w:spacing w:val="20"/><w:kern w:val="28"/><w:sz w:val="28"/>' +
    '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/><w:lang w:val="fr-FR"/>'
  const RECORD = `<w:rPrChange w:id="9" w:author="Reviewer" w:date="2026-07-01T00:00:00Z"><w:rPr>${OLD_RPR}</w:rPr></w:rPrChange>`
  const P = `<w:p><w:r><w:rPr><w:b/><w:color w:val="FF0000"/>${RECORD}</w:rPr><w:t>x</w:t></w:r></w:p>`

  const documentXml = async (bytes: Uint8Array) =>
    (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')

  it('parses oldRPr next to the modeled subset', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: P }))
    const run = parsed.blocks[0].runs![0]
    expect(run.rPrChange?.oldRPr).toBe(OLD_RPR)
    expect(run.rPrChange?.old).toMatchObject({
      italic: true,
      color: '0000FF',
      fontAscii: 'Arial',
      charSpacingTwips: 20,
      kernHalfPoints: 28,
      sizeHalfPoints: 28,
      shading: 'FFFF00',
    })
  })

  it('a rebuilt record re-emits the previous rPr including unmodeled children', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: P }))
    const run = { ...parsed.blocks[0].runs![0], text: 'edited', rawRPr: undefined }
    const xml = await documentXml(
      await saveDocx(parsed, [
        { kind: 'generated', block: { type: 'paragraph', runs: [run] }, docxIndex: 0 },
      ]),
    )
    expect(xml).toContain(RECORD)
  })

  it('reject (run rebuilt from oldRPr) saves the previous rPr byte-for-byte; accept only drops the record', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: P }))
    const current = parsed.blocks[0].runs![0]
    const { rPrChange, ...rest } = current
    const rejected = {
      ...rest,
      bold: undefined,
      color: rPrChange!.old!.color,
      italic: true,
      fontAscii: 'Arial',
      font: 'Arial',
      charSpacingTwips: 20,
      kernHalfPoints: 28,
      sizeHalfPoints: 28,
      shading: 'FFFF00',
      rawRPr: `<w:rPr>${rPrChange!.oldRPr}</w:rPr>`,
    }
    const rejectedXml = await documentXml(
      await saveDocx(parsed, [
        { kind: 'generated', block: { type: 'paragraph', runs: [rejected] }, docxIndex: 0 },
      ]),
    )
    expect(rejectedXml).toContain(`<w:r><w:rPr>${OLD_RPR}</w:rPr><w:t`)

    const accepted = { ...rest, rawRPr: current.rawRPr!.replace(RECORD, '') }
    const acceptedXml = await documentXml(
      await saveDocx(parsed, [
        { kind: 'generated', block: { type: 'paragraph', runs: [accepted] }, docxIndex: 0 },
      ]),
    )
    expect(acceptedXml).toContain('<w:r><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr><w:t')
    expect(acceptedXml).not.toContain('w:rPrChange')
  })
})
