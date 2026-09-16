import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { run, tempDir } from './helpers'

async function patchZip(
  path: string,
  edits: Record<string, (xml: string) => string>,
): Promise<void> {
  const zip = await JSZip.loadAsync(readFileSync(path))
  for (const [name, edit] of Object.entries(edits)) {
    zip.file(name, edit(await zip.file(name)!.async('string')))
  }
  writeFileSync(path, await zip.generateAsync({ type: 'nodebuffer' }))
}

describe('sheet check', () => {
  it('reports nothing on a plain table', async () => {
    const dir = tempDir()
    const table = join(dir, 't.json')
    writeFileSync(
      table,
      JSON.stringify([
        ['a', 'b'],
        [1, 2],
      ]),
    )
    const out = join(dir, 'clean.xlsx')
    expect((await run(['create', '--type', 'xlsx', '--from', table, '--out', out])).code).toBe(0)
    const r = await run(['sheet', 'check', out, '--json'])
    expect(r.code).toBe(0)
    expect(r.json().detail.issues).toEqual([])
    expect(r.json().summary).toContain('no issues found')
  })

  it('finds formula errors, missing sheets, broken names, ### columns and placeholders', async () => {
    const dir = tempDir()
    const table = join(dir, 't.json')
    writeFileSync(
      table,
      JSON.stringify([
        ['a', 'b'],
        [1, 2],
      ]),
    )
    const out = join(dir, 'bad.xlsx')
    expect((await run(['create', '--type', 'xlsx', '--from', table, '--out', out])).code).toBe(0)
    const zip = await JSZip.loadAsync(readFileSync(out))
    zip.file(
      'xl/charts/chart1.xml',
      '<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numRef><c:f>[1]Budget!$B$2:$B$5</c:f></c:numRef></c:val></c:ser><c:ser><c:val><c:numRef><c:f>\'C:\\Data\\[Book.xlsx]Far Sheet\'!$B$2:$B$5</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>',
    )
    writeFileSync(out, await zip.generateAsync({ type: 'nodebuffer' }))
    await patchZip(out, {
      'xl/worksheets/sheet1.xml': (xml) =>
        xml
          .replace(
            '<sheetData',
            '<cols><col min="3" max="3" width="4" customWidth="1"/></cols><sheetData',
          )
          .replace(
            '</sheetData>',
            '<row r="3"><c r="A3" t="inlineStr"><is><t>{{customer}}</t></is></c>' +
              '<c r="B3" t="e"><f>1/0</f><v>#DIV/0!</v></c>' +
              '<c r="C3"><v>123456789</v></c></row>' +
              '<row r="4"><c r="B4"><f>Missing!A1+1</f></c>' +
              "<c r=\"D4\"><f>'Also Gone'!A1+[1]Other!A1+'[Book.xlsx]Far Sheet'!B2+'C:\\Data\\[Book.xlsx]Far Sheet'!B3</f><v>3</v></c>" +
              '<c r="E4"><f>SUM(T!A1:A2)</f><v>1</v></c>' +
              '<c r="F4" t="str"><f>IF(A1="","","")</f><v></v></c>' +
              '<c r="C4"><f>SUM(A1:A2)</f><v>1</v></c></row>' +
              '<row hidden="1" r="5"><c r="D5"><v>987654321</v></c></row>' +
              '<row r="7"><c r="F7" s="2"><v>44927</v></c></row>' +
              '<row r="6"><c r="E6" s="1"><v>1234567.5</v></c></row></sheetData>',
          ),
      'xl/styles.xml': (xml) =>
        xml
          .replace(
            /<cellXfs count="(\d+)">([\s\S]*?)<\/cellXfs>/,
            (_m, n: string, body: string) =>
              `<cellXfs count="${Number(n) + 2}">${body}<xf numFmtId="7" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/><xf numFmtId="200" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/></cellXfs>`,
          )
          .replace(
            /<numFmts count="(\d+)">/,
            (_m, n: string) =>
              `<numFmts count="${Number(n) + 1}"><numFmt formatCode="dddd, mmmm dd, yyyy" numFmtId="200"/>`,
          )
          .replace(/<fonts\b/, (m) =>
            /<numFmts\b/.test(xml)
              ? m
              : `<numFmts count="1"><numFmt formatCode="dddd, mmmm dd, yyyy" numFmtId="200"/></numFmts>${m}`,
          ),
      'xl/workbook.xml': (xml) =>
        xml.replace(
          '</sheets>',
          '</sheets><definedNames><definedName name="Bad">#REF!</definedName>' +
            '<definedName name="Gone">Nope!$A$1</definedName></definedNames>',
        ),
    })
    const r = await run(['sheet', 'check', out, '--json'])
    expect(r.code).toBe(0)
    const issues = r.json().detail.issues as {
      id: string
      code: string
      level: string
      path: string
      context?: string
      message: string
      suggest?: Record<string, unknown>
    }[]
    const codes = issues.map((i) => i.code)
    expect(codes).toContain('formula_error')
    expect(codes).toContain('formula_not_evaluated')
    expect(codes).toContain('defined_name_broken')
    expect(codes).toContain('number_overflow')
    expect(codes).toContain('placeholder_left')
    const missingSheets = issues.filter((i) => i.code === 'missing_sheet_ref').map((i) => i.message)
    expect(missingSheets).toHaveLength(3)
    expect(missingSheets.some((m) => m.endsWith("'Also Gone'"))).toBe(true)
    expect(missingSheets.some((m) => m.includes('Far Sheet') || m.includes('Other'))).toBe(false)
    expect(codes).not.toContain('chart_ref')
    expect(issues.find((i) => i.code === 'formula_error')).toMatchObject({
      level: 'error',
      path: 't!B3',
      context: '=1/0',
    })
    expect(issues.find((i) => i.code === 'formula_not_evaluated')?.context).toBe('B4')
    const overflow = issues.find((i) => i.code === 'number_overflow')!
    expect(overflow.suggest).toMatchObject({ op: 'set_col_width', sheet: 't', column: 'C' })
    const columns = issues.filter((i) => i.code === 'number_overflow').map((i) => i.path)
    expect(columns).toContain('t!E:E')
    expect(columns).toContain('t!F:F')
    expect(columns).not.toContain('t!D:D')
    expect(Number(overflow.suggest!.widthPx)).toBeGreaterThan(60)
    expect(issues[0]!.level).toBe('error')
    expect(issues[0]!.id).toBe('E1')
    expect(r.json().detail.counts.error).toBe(5)
  })
})

describe('docs check', () => {
  it('finds unevaluated TOC fields, heading skips, placeholders and broken anchors', async () => {
    const dir = tempDir()
    const md = join(dir, 'doc.md')
    writeFileSync(
      md,
      '# Report\n\n### Deep dive\n\nTODO fill this in\n\n## Findings\n\nAll good.\n',
    )
    const out = join(dir, 'doc.docx')
    expect((await run(['convert', md, '--to', 'docx', '--out', out])).code).toBe(0)
    const ops = join(dir, 'ops.json')
    writeFileSync(ops, JSON.stringify([{ op: 'insertToc', afterBlockIndex: 0 }]))
    expect((await run(['docs', 'apply', out, '--ops', ops, '--json'])).code).toBe(0)
    await patchZip(out, {
      'word/document.xml': (xml) =>
        xml.replace(
          '<w:sectPr',
          '<w:p><w:hyperlink w:anchor="nowhere"><w:r><w:t>see above</w:t></w:r></w:hyperlink></w:p>' +
            '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGEREF </w:instrText></w:r>' +
            '<w:r><w:instrText xml:space="preserve">split_target \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
            '<w:r><w:t>3</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' +
            '<w:p><w:fldSimple w:instr=" REF &quot;quoted_target&quot; \\h "><w:r><w:t>4</w:t></w:r></w:fldSimple></w:p>' +
            '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r>' +
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p><w:sectPr',
        ),
    })
    const r = await run(['docs', 'check', out, '--json'])
    expect(r.code).toBe(0)
    const issues = r.json().detail.issues as {
      code: string
      level: string
      blockIndex?: number
      message: string
    }[]
    const codes = issues.map((i) => i.code)
    expect(codes).toContain('field_not_evaluated')
    expect(codes).toContain('heading_skip')
    expect(codes).toContain('placeholder_left')
    expect(codes).toContain('broken_ref')
    expect(codes).not.toContain('toc_stale')
    expect(issues.find((i) => i.code === 'broken_ref')).toMatchObject({
      level: 'error',
      message: 'hyperlink points at bookmark "nowhere", which does not exist',
    })
    expect(issues.find((i) => i.code === 'heading_skip')?.message).toContain('from 1 to 3')
    expect(issues.map((i) => i.message)).toContain(
      'PAGEREF points at bookmark "split_target", which does not exist',
    )
    expect(issues.map((i) => i.message)).toContain(
      'REF points at bookmark "quoted_target", which does not exist',
    )
    const empty = issues.filter((i) => i.message.includes('has no cached result'))
    expect(empty.map((i) => i.message)).toEqual([
      'PAGE has no cached result; it shows empty until Word updates fields',
    ])
    const clean = await run(['docs', 'check', join(dir, 'doc.docx')])
    expect(clean.code).toBe(0)
    expect(clean.stdout).toContain('issue(s)')
  })
})
