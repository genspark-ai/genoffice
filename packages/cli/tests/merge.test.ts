import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { xlsxSidecarPath } from '../src/resources'
import { run, tempDir } from './helpers'

const INCH = 914400
const sidecar = Boolean(xlsxSidecarPath())

const DATA = { name: 'Ada', amount: 12.5, due: { date: '2026-10-01' }, item: 'Widget', extra: 'x' }

async function documentXml(path: string): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(path))
  return zip.file('word/document.xml')!.async('string')
}

/** an outer group holding an inner group (sp X) and a direct text child (sp Y), as PowerPoint nests them */
async function nestGroup(path: string): Promise<void> {
  const sp = (id: number, name: string, text: string) =>
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`
  const grp = (id: number, name: string, body: string) =>
    `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/>` +
    `<a:chOff x="0" y="0"/><a:chExt cx="2000000" cy="1000000"/></a:xfrm></p:grpSpPr>${body}</p:grpSp>`
  const zip = await JSZip.loadAsync(readFileSync(path))
  const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const outer = grp(
    40,
    'outer',
    grp(41, 'inner', sp(42, 'X', 'Deep {{name}}')) + sp(43, 'Y', 'Near {{item}}'),
  )
  zip.file('ppt/slides/slide1.xml', xml.replace('</p:spTree>', `${outer}</p:spTree>`))
  writeFileSync(path, await zip.generateAsync({ type: 'nodebuffer' }))
}

function dataFile(dir: string, data: unknown): string {
  const path = join(dir, 'data.json')
  writeFileSync(path, JSON.stringify(data))
  return path
}

async function templateDocx(dir: string): Promise<string> {
  const md = join(dir, 'template.md')
  writeFileSync(
    md,
    '# Invoice for {{name}}\n\nAmount due: **{{ amount }}** by {{due.date}}.\n\nRef {{na**me**}} is split.\n\n| Item | Cost |\n| --- | --- |\n| {{item}} | {{missing}} |\n\n| Untouched |\n| --- |\n| {{absent}} |\n',
  )
  const out = join(dir, 'template.docx')
  expect((await run(['create', '--type', 'docx', '--from', md, '--out', out])).code).toBe(0)
  return out
}

describe('genoffice merge', () => {
  it('fills a docx template, keeps table cells and reports split and unknown placeholders', async () => {
    const dir = tempDir()
    const template = await templateDocx(dir)
    const out = join(dir, 'invoice.docx')
    const r = await run(['merge', template, '--data', dataFile(dir, DATA), '--out', out, '--json'])
    expect(r.code).toBe(0)
    const j = r.json()
    expect(j.summary).toBe('4 placeholders filled, 3 unresolved')
    expect(j.detail).toMatchObject({
      format: 'docx',
      filled: 4,
      used_keys: ['name', 'amount', 'due.date', 'item'],
      unused_keys: ['extra'],
    })
    expect(j.detail.unresolved_placeholders).toEqual([
      {
        placeholder: '{{name}}',
        key: 'name',
        reason: 'split_placeholder',
        location: { block: 2 },
      },
      { placeholder: '{{missing}}', key: 'missing', reason: 'no_key', location: { block: 3 } },
      { placeholder: '{{absent}}', key: 'absent', reason: 'no_key', location: { block: 4 } },
    ])
    expect(j.warnings[0]).toMatchObject({ code: 'unresolved_placeholder' })
    expect(j.warnings[0].message).toContain('{{missing}} (block 3)')
    // a table nothing fills is never round-tripped through HTML: its XML is the template's
    const tables = async (path: string) =>
      [...(await documentXml(path)).matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)].map((m) => m[0])
    expect((await tables(out))[1]).toBe((await tables(template))[1])

    const read = await run(['docs', 'read', out, '--full', '--json'])
    const texts = read.json().detail.items.map((b: { text: string }) => b.text)
    expect(texts[0]).toBe('Invoice for Ada')
    expect(texts[1]).toBe('Amount due: 12.5 by 2026-10-01.')
    expect(texts[2]).toBe('Ref {{name}} is split.')
    expect(texts[3]).toContain('Widget')
    expect(texts[3]).toContain('{{missing}}')
    expect(read.json().detail.items[3].table.rows).toBe(2)
    expect(texts[4]).toContain('{{absent}}')
  })

  it('refuses to write with --strict, takes inline data and flattens nested objects', async () => {
    const dir = tempDir()
    const template = await templateDocx(dir)
    const out = join(dir, 'strict.docx')
    const strict = await run([
      'merge',
      template,
      '--data',
      JSON.stringify(DATA),
      '--out',
      out,
      '--strict',
      '--json',
    ])
    expect(strict.code).toBe(1)
    expect(strict.json()).toMatchObject({ status: 'error', error: 'unresolved_placeholder' })
    expect(strict.json().detail.unresolved_placeholders).toHaveLength(3)
    expect(existsSync(out)).toBe(false)

    const full = { ...DATA, missing: 'now here', absent: 'here too', name: 'Grace' }
    const ok = await run([
      'merge',
      template,
      '--data',
      JSON.stringify(full),
      '--out',
      out,
      '--json',
    ])
    expect(ok.code).toBe(0)
    expect(ok.json().detail.unresolved_placeholders).toEqual([
      expect.objectContaining({ key: 'name', reason: 'split_placeholder' }),
    ])
    expect(ok.json().detail.used_keys).toContain('missing')

    const again = await run(['merge', template, '--data', JSON.stringify(full), '--out', out])
    expect(again.code).toBe(2)
    expect(again.stderr).toContain('output exists')
  })

  it('treats placeholder-shaped text inside a value as output, in every format', async () => {
    const dir = tempDir()
    const md = join(dir, 'lit.md')
    writeFileSync(md, 'Hello {{name}} and {{missing}}\n')
    const docx = join(dir, 'lit.docx')
    expect((await run(['create', '--type', 'docx', '--from', md, '--out', docx])).code).toBe(0)
    const ops = join(dir, 'lit-ops.json')
    writeFileSync(
      ops,
      JSON.stringify([
        {
          op: 'addElement',
          target: { slide: 0 },
          kind: 'textbox',
          offset: { x: INCH, y: INCH, cx: 8 * INCH, cy: INCH },
          paragraphs: [{ runs: [{ text: 'Hello {{name}} and {{missing}}' }] }],
        },
      ]),
    )
    const pptx = join(dir, 'lit.pptx')
    expect((await run(['create', '--type', 'pptx', '--ops', ops, '--out', pptx])).code).toBe(0)
    const table = join(dir, 'lit.json')
    writeFileSync(table, JSON.stringify([['Hello {{name}} and {{missing}}']]))
    const xlsx = join(dir, 'lit.xlsx')
    expect((await run(['create', '--type', 'xlsx', '--from', table, '--out', xlsx])).code).toBe(0)

    const data = JSON.stringify({ name: '{{literal}}' })
    const both = JSON.stringify({ name: '{{literal}}', missing: 'm' })
    for (const template of sidecar ? [docx, pptx, xlsx] : [docx, pptx]) {
      const out = template.replace('lit.', 'lit-out.')
      const r = await run(['merge', template, '--data', data, '--out', out, '--json'])
      expect(r.code, template).toBe(0)
      expect(r.json().summary, template).toBe('1 placeholders filled, 1 unresolved')
      expect(r.json().detail.used_keys, template).toEqual(['name'])
      expect(r.json().detail.unresolved_placeholders.map((u: { key: string }) => u.key)).toEqual([
        'missing',
      ])
      const strict = await run([
        'merge',
        template,
        '--data',
        data,
        '--out',
        out,
        '--strict',
        '--force',
        '--json',
      ])
      expect(strict.json().error, template).toBe('unresolved_placeholder')
      expect(
        strict.json().detail.unresolved_placeholders.map((u: { key: string }) => u.key),
      ).toEqual(['missing'])
      const ok = await run([
        'merge',
        template,
        '--data',
        both,
        '--out',
        out,
        '--strict',
        '--force',
        '--json',
      ])
      expect(ok.code, template).toBe(0)
      expect(ok.json().summary, template).toBe('2 placeholders filled, 0 unresolved')
    }
    const docText = (
      await run(['docs', 'read', docx.replace('lit.', 'lit-out.'), '--full', '--json'])
    ).json().detail.items[0].text
    expect(docText).toBe('Hello {{literal}} and m')
    const deck = (
      await run(['slides', 'read', pptx.replace('lit.', 'lit-out.'), '--full', '--json'])
    ).json()
    expect(deck.detail.pages[0].elements.map((e: { text?: string }) => e.text)).toContain(
      'Hello {{literal}} and m',
    )
    if (sidecar) {
      const book = (await run(['sheet', 'read', xlsx.replace('lit.', 'lit-out.'), '--json'])).json()
      expect(book.detail.rows[0]).toEqual(['Hello {{literal}} and m'])
    }
  })

  it('fills a body paragraph with a multi-line value as line breaks and strips NUL', async () => {
    const dir = tempDir()
    const md = join(dir, 'addr.md')
    writeFileSync(md, 'Ship to: **{{addr}}**\n\nRef {{ref}}\n')
    const template = join(dir, 'addr.docx')
    expect((await run(['create', '--type', 'docx', '--from', md, '--out', template])).code).toBe(0)
    const out = join(dir, 'addr-out.docx')
    const data = JSON.stringify({ addr: 'Line 1\nLine 2', ref: 'A\u0000B' })
    const r = await run(['merge', template, '--data', data, '--out', out, '--strict', '--json'])
    expect(r.code).toBe(0)
    expect(r.json().summary).toBe('2 placeholders filled, 0 unresolved')
    const read = await run(['docs', 'read', out, '--full', '--json'])
    const texts = read.json().detail.items.map((b: { text: string }) => b.text)
    expect(texts[0]).toContain('Line 1')
    expect(texts[0]).toContain('Line 2')
    expect(texts[1]).toBe('Ref AB')
    const xml = await documentXml(out)
    expect(xml).toContain('<w:br/>')
    expect(xml).not.toContain('{{')
  })

  it('keeps a multi-line value literal when its text names a placeholder filled elsewhere', async () => {
    const dir = tempDir()
    const md = join(dir, 'routes.md')
    writeFileSync(md, '{{addr}}\n\n{{name}}\n')
    const template = join(dir, 'routes.docx')
    expect((await run(['create', '--type', 'docx', '--from', md, '--out', template])).code).toBe(0)
    const out = join(dir, 'routes-out.docx')
    const data = JSON.stringify({ addr: 'A\nB {{name}}', name: 'X' })
    const r = await run(['merge', template, '--data', data, '--out', out, '--strict', '--json'])
    expect(r.code).toBe(0)
    expect(r.json().summary).toBe('2 placeholders filled, 0 unresolved')
    expect(r.json().detail.used_keys).toEqual(['addr', 'name'])
    const read = await run(['docs', 'read', out, '--full', '--json'])
    const texts = read.json().detail.items.map((b: { text: string }) => b.text)
    expect(texts[0]).toMatch(/^A.?B \{\{name\}\}$/)
    expect(texts[1]).toBe('X')
    expect(await documentXml(out)).toContain('<w:br/>')
  })

  it('substitutes in one pass: a value that reads like another placeholder is output', async () => {
    const dir = tempDir()
    const md = join(dir, 'pass.md')
    writeFileSync(md, '{{a}} {{b}}\n')
    const docx = join(dir, 'pass.docx')
    expect((await run(['create', '--type', 'docx', '--from', md, '--out', docx])).code).toBe(0)
    const ops = join(dir, 'pass-ops.json')
    writeFileSync(
      ops,
      JSON.stringify([
        {
          op: 'addElement',
          target: { slide: 0 },
          kind: 'textbox',
          offset: { x: INCH, y: INCH, cx: 8 * INCH, cy: INCH },
          paragraphs: [{ runs: [{ text: '{{a}} {{b}}' }] }],
        },
      ]),
    )
    const pptx = join(dir, 'pass.pptx')
    expect((await run(['create', '--type', 'pptx', '--ops', ops, '--out', pptx])).code).toBe(0)
    const table = join(dir, 'pass.json')
    writeFileSync(table, JSON.stringify([['{{a}} {{b}}']]))
    const xlsx = join(dir, 'pass.xlsx')
    expect((await run(['create', '--type', 'xlsx', '--from', table, '--out', xlsx])).code).toBe(0)

    const data = JSON.stringify({ a: '{{b}}', b: 'X' })
    for (const template of sidecar ? [docx, pptx, xlsx] : [docx, pptx]) {
      const out = template.replace('pass.', 'pass-out.')
      const r = await run(['merge', template, '--data', data, '--out', out, '--strict', '--json'])
      expect(r.code, template).toBe(0)
      expect(r.json().summary, template).toBe('2 placeholders filled, 0 unresolved')
      expect(r.json().detail.used_keys, template).toEqual(['a', 'b'])
    }
    const doc = (
      await run(['docs', 'read', docx.replace('pass.', 'pass-out.'), '--full', '--json'])
    ).json()
    expect(doc.detail.items[0].text).toBe('{{b}} X')
    const deck = (
      await run(['slides', 'read', pptx.replace('pass.', 'pass-out.'), '--full', '--json'])
    ).json()
    expect(deck.detail.pages[0].elements.map((e: { text?: string }) => e.text)).toContain('{{b}} X')
    if (sidecar) {
      const book = (
        await run(['sheet', 'read', xlsx.replace('pass.', 'pass-out.'), '--json'])
      ).json()
      expect(book.detail.rows[0]).toEqual(['{{b}} X'])
    }
  })

  it('rejects data that is not an object and templates of other types', async () => {
    const dir = tempDir()
    const template = await templateDocx(dir)
    const list = await run([
      'merge',
      template,
      '--data',
      '[1,2]',
      '--out',
      join(dir, 'x.docx'),
      '--json',
    ])
    expect(list.json()).toMatchObject({ status: 'error', error: 'invalid_argument' })
    const bad = await run([
      'merge',
      template,
      '--data',
      '{oops',
      '--out',
      join(dir, 'x.docx'),
      '--json',
    ])
    expect(bad.json()).toMatchObject({ status: 'error', error: 'invalid_json' })
    const none = await run(['merge', template, '--out', join(dir, 'x.docx'), '--json'])
    expect(none.json()).toMatchObject({ status: 'error', error: 'missing_argument' })
    const md = join(dir, 'template.md')
    const other = await run(['merge', md, '--data', '{}', '--out', join(dir, 'x.md'), '--json'])
    expect(other.json()).toMatchObject({ status: 'error', error: 'unsupported' })
    expect(other.json().detail.supported).toEqual(['docx', 'pptx', 'xlsx'])
  })

  it('fills a pptx template across text boxes, table cells and speaker notes', async () => {
    const dir = tempDir()
    const create = join(dir, 'create.json')
    writeFileSync(
      create,
      JSON.stringify([
        {
          op: 'addElement',
          target: { slide: 0 },
          kind: 'textbox',
          offset: { x: INCH, y: INCH, cx: 8 * INCH, cy: INCH },
          paragraphs: [{ runs: [{ text: 'Dear {{ name }}, your total is {{amount}}' }] }],
        },
        {
          op: 'addElement',
          target: { slide: 0 },
          kind: 'textbox',
          offset: { x: INCH, y: 3 * INCH, cx: 8 * INCH, cy: INCH },
          paragraphs: [{ runs: [{ text: 'Split {{na' }, { text: 'me}}', bold: true }] }],
        },
        {
          op: 'addTable',
          target: { slide: 0 },
          rows: 1,
          cols: 2,
          offset: { x: INCH, y: 4 * INCH, cx: 6 * INCH, cy: INCH },
        },
        { op: 'setNotes', target: { slide: 0 }, text: 'Call {{name}} about {{topic}}' },
      ]),
    )
    const template = join(dir, 'template.pptx')
    expect((await run(['create', '--type', 'pptx', '--ops', create, '--out', template])).code).toBe(
      0,
    )
    const before = await run(['slides', 'read', template, '--json'])
    const table = before
      .json()
      .detail.pages[0].elements.find((e: { type: string }) => e.type === 'table')
    const cell = join(dir, 'cell.json')
    writeFileSync(
      cell,
      JSON.stringify([
        {
          op: 'setTableCell',
          target: { slide: 0, el: table.id },
          row: 0,
          col: 1,
          paragraphs: [{ runs: [{ text: 'Item: {{item}}' }] }],
        },
      ]),
    )
    expect((await run(['slides', 'apply', template, '--ops', cell])).code).toBe(0)
    await nestGroup(template)

    const out = join(dir, 'deck.pptx')
    const r = await run(['merge', template, '--data', dataFile(dir, DATA), '--out', out, '--json'])
    expect(r.code).toBe(0)
    // the deck engine fills a placeholder split over runs (genoffice#1052); only the
    // nested-group and missing-key cases stay unresolved
    expect(r.json().summary).toBe('6 placeholders filled, 2 unresolved')
    expect(r.json().detail.used_keys).toEqual(['name', 'amount', 'item'])
    expect(r.json().detail.unresolved_placeholders).toEqual([
      expect.objectContaining({
        placeholder: '{{name}}',
        reason: 'unreachable_nested',
        location: { slide: 's_1', element: expect.any(String) },
      }),
      {
        placeholder: '{{topic}}',
        key: 'topic',
        reason: 'no_key',
        location: { slide: 's_1', notes: true },
      },
    ])
    expect(r.json().warnings[0].suggestion).toContain('ungroup the nested groups')

    const read = await run(['slides', 'read', out, '--full', '--json'])
    const page = read.json().detail.pages[0]
    const texts = page.elements.map((e: { text?: string }) => e.text)
    expect(texts).toContain('Dear Ada, your total is 12.5')
    expect(texts).toContain('Split Ada')
    expect(texts.some((t: string) => t?.includes('Item: Widget'))).toBe(true)
    const outer = page.elements.find((e: { type: string }) => e.type === 'group')
    const deep = outer.children.find((e: { type: string }) => e.type === 'group').children[0]
    expect(deep.text).toBe('Deep {{name}}')
    expect(outer.children.find((e: { text?: string }) => e.text?.startsWith('Near')).text).toBe(
      'Near Widget',
    )
    expect(page.notes).toBe('Call Ada about {{topic}}')
  })

  it.skipIf(!sidecar)('fills an xlsx template, typing whole-cell numbers as numbers', async () => {
    const dir = tempDir()
    const table = join(dir, 'table.json')
    writeFileSync(
      table,
      JSON.stringify({
        sheets: [
          {
            name: 'Invoice',
            rows: [
              ['Customer', '{{name}}'],
              ['Amount', ' {{ amount }} '],
              ['Note', 'Paid by {{name}} ({{ref}})'],
              ['Total', '=B2*2'],
              ['Formula-looking', '{{formula}}'],
            ],
          },
          { name: 'Meta', rows: [['{{due.date}}', 'plain']] },
        ],
      }),
    )
    const template = join(dir, 'template.xlsx')
    expect((await run(['create', '--type', 'xlsx', '--from', table, '--out', template])).code).toBe(
      0,
    )
    const out = join(dir, 'book.xlsx')
    const data = dataFile(dir, { ...DATA, formula: '=SUM(A1:A2)' })
    const r = await run(['merge', template, '--data', data, '--out', out, '--json'])
    expect(r.code).toBe(0)
    expect(r.json().summary).toBe('5 placeholders filled, 1 unresolved')
    expect(r.json().detail.used_keys).toEqual(['name', 'amount', 'formula', 'due.date'])
    expect(r.json().detail.unresolved_placeholders).toEqual([
      {
        placeholder: '{{ref}}',
        key: 'ref',
        reason: 'no_key',
        location: { sheet: 'Invoice', cell: 'B3' },
      },
    ])

    const invoice = await run(['sheet', 'read', out, '--sheet', 'Invoice', '--json'])
    const rows = invoice.json().detail.rows
    expect(rows[0]).toEqual(['Customer', 'Ada'])
    expect(rows[1]).toEqual(['Amount', 12.5])
    expect(rows[2]).toEqual(['Note', 'Paid by Ada ({{ref}})'])
    expect(rows[3][0]).toBe('Total')
    expect(rows[4]).toEqual(['Formula-looking', '=SUM(A1:A2)'])
    expect(invoice.json().detail.formulas).toEqual({ B4: '=B2*2' })
    const meta = await run(['sheet', 'read', out, '--sheet', 'Meta', '--json'])
    expect(meta.json().detail.rows[0]).toEqual(['2026-10-01', 'plain'])
  })

  it.skipIf(!sidecar)(
    'fills an xlsx template whose used range exceeds the DSL expansion cap',
    async () => {
      const dir = tempDir()
      const rows = Array.from({ length: 250 }, (_, r) =>
        Array.from({ length: 10 }, (_, c) => `r${r + 1}c${c + 1}`),
      )
      rows[0]![0] = 'Report for {{name}}'
      rows[124]![9] = '{{amount}}'
      rows[249]![4] = 'Due {{due.date}}'
      const table = join(dir, 'wide.json')
      writeFileSync(table, JSON.stringify(rows))
      const template = join(dir, 'wide.xlsx')
      expect(
        (await run(['create', '--type', 'xlsx', '--from', table, '--out', template])).code,
      ).toBe(0)
      const out = join(dir, 'wide-out.xlsx')
      const r = await run([
        'merge',
        template,
        '--data',
        dataFile(dir, DATA),
        '--out',
        out,
        '--json',
      ])
      expect(r.code).toBe(0)
      expect(r.json().summary).toBe('3 placeholders filled, 0 unresolved')
      const read = await run(['sheet', 'read', out, '--range', 'A1:J250', '--json'])
      const got = read.json().detail.rows
      expect(got[0][0]).toBe('Report for Ada')
      expect(got[124][9]).toBe(12.5)
      expect(got[249][4]).toBe('Due 2026-10-01')
      expect(got[1][1]).toBe('r2c2')
    },
  )
})
