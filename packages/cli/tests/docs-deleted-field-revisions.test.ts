import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../docx-engine/tests/helpers/build-docx'
import { run, tempDir } from './helpers'

const DATE = '2026-01-01T00:00:00Z'

/** genoffice#2808: a field (here the usual HYPERLINK) deleted with track changes —
 *  Word wraps the field's runs in one w:del: fldChar begin, delInstrText,
 *  fldChar separate, the w:delText result runs, fldChar end. */
const deletedFieldRuns =
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:delInstrText xml:space="preserve"> HYPERLINK "https://example.com" </w:delInstrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:delText>link text</w:delText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
const deletedField = (id: string, author = 'A') =>
  `<w:del w:id="${id}" w:author="${author}" w:date="${DATE}">${deletedFieldRuns}</w:del>`

/** the same field inserted with track changes: w:ins around fldChar and instrText */
const insertedField = (id: string, author = 'A') =>
  `<w:ins w:id="${id}" w:author="${author}" w:date="${DATE}">` +
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> HYPERLINK "https://example.com" </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:t>link text</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:ins>'

async function fixture(dir: string, body: string) {
  const source = join(dir, 'source.docx')
  writeFileSync(
    source,
    await buildDocx({
      bodyXml: body,
      extraRels:
        '<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>',
      extraParts: [
        {
          path: 'customXml/sentinel.xml',
          contentType: 'application/xml',
          xml: '<untouched>Exact sentinel bytes</untouched>',
        },
      ],
    }),
  )
  return source
}

interface ListedRevision {
  id: string
  type: string
  kind: string
  author: string
  blockIndex: number
  text: string
}

async function revisions(path: string) {
  const result = await run(['docs', 'read', path, '--revisions', '--json'])
  expect(result.code, result.stderr).toBe(0)
  return result.json().detail.revisions as ListedRevision[]
}

async function apply(source: string, out: string, op: Record<string, unknown>) {
  const ops = join(tempDir(), 'ops.json')
  writeFileSync(ops, JSON.stringify([op]))
  return run(['docs', 'apply', source, '--ops', ops, '--out', out, '--json'])
}

describe('deleted fields through the real Docs CLI (genoffice#2808)', () => {
  it('a deleted field is one deletion revision and accept removes it with the wrapper', async () => {
    const dir = tempDir(),
      out = join(dir, 'accepted.docx')
    const source = await fixture(
      dir,
      `<w:p><w:r><w:t xml:space="preserve">A </w:t></w:r>${deletedField('1')}</w:p>`,
    )
    const listed = await revisions(source)
    expect(listed).toHaveLength(1)
    expect(listed[0]).toMatchObject({
      type: 'deletion',
      kind: 'del',
      author: 'A',
      text: 'link text',
    })
    const result = await apply(source, out, { op: 'accept_changes', all: true })
    expect(result.code, result.stderr || result.stdout).toBe(0)
    expect(await revisions(out)).toEqual([])
    const before = await JSZip.loadAsync(readFileSync(source)),
      after = await JSZip.loadAsync(readFileSync(out))
    const xml = await after.file('word/document.xml')!.async('string')
    expect(xml).not.toMatch(/<w:(?:ins|del|fldChar|hyperlink)\b/)
    expect(xml).not.toContain('link text')
    expect(xml).toContain('A ')
    expect(await after.file('customXml/sentinel.xml')!.async('string')).toBe(
      await before.file('customXml/sentinel.xml')!.async('string'),
    )
  })

  it('rejecting the deletion brings the hyperlink back as live text', async () => {
    const dir = tempDir(),
      out = join(dir, 'rejected.docx')
    const source = await fixture(
      dir,
      `<w:p><w:r><w:t xml:space="preserve">A </w:t></w:r>${deletedField('1')}</w:p>`,
    )
    expect(await revisions(source)).toHaveLength(1)
    const result = await apply(source, out, { op: 'reject_changes', all: true })
    expect(result.code, result.stderr || result.stdout).toBe(0)
    expect(await revisions(out)).toEqual([])
    const xml = await JSZip.loadAsync(readFileSync(out)).then((z) =>
      z.file('word/document.xml')!.async('string'),
    )
    expect(xml).not.toMatch(/<w:(?:ins|del)\b/)
    expect(xml).toContain('link text')
    expect(xml).toContain('<w:hyperlink')
  })

  it('an unrelated w:ins in the same paragraph is listed beside the deleted field', async () => {
    const dir = tempDir(),
      out = join(dir, 'mixed.docx')
    const source = await fixture(
      dir,
      `<w:p><w:r><w:t xml:space="preserve">A </w:t></w:r>${deletedField('1')}` +
        `<w:ins w:id="90" w:author="A" w:date="${DATE}"><w:r><w:t> other</w:t></w:r></w:ins></w:p>`,
    )
    const listed = await revisions(source)
    expect(listed).toHaveLength(2)
    expect(listed[0]).toMatchObject({ type: 'deletion', author: 'A', text: 'link text' })
    expect(listed[1]).toMatchObject({ type: 'insertion', author: 'A', text: 'other' })
    // accepting only the deletion keeps the insertion pending
    const first = await apply(source, out, { op: 'accept_changes', ids: ['r1'] })
    expect(first.code, first.stderr || first.stdout).toBe(0)
    const still = await revisions(out)
    expect(still).toHaveLength(1)
    expect(still[0]).toMatchObject({ type: 'insertion', text: 'other' })
    const second = await apply(out, join(dir, 'final.docx'), { op: 'accept_changes', all: true })
    expect(second.code, second.stderr || second.stdout).toBe(0)
    expect(await revisions(join(dir, 'final.docx'))).toEqual([])
    const xml = await JSZip.loadAsync(readFileSync(join(dir, 'final.docx'))).then((z) =>
      z.file('word/document.xml')!.async('string'),
    )
    expect(xml).not.toMatch(/<w:(?:ins|del)\b/)
    expect(xml).toContain('A ')
    expect(xml).toContain('other')
    expect(xml).not.toContain('link text')
  })

  it('an inserted field still lists as one insertion revision', async () => {
    const dir = tempDir(),
      out = join(dir, 'inserted.docx')
    const source = await fixture(
      dir,
      `<w:p><w:r><w:t xml:space="preserve">A </w:t></w:r>${insertedField('1')}</w:p>`,
    )
    const listed = await revisions(source)
    expect(listed).toHaveLength(1)
    expect(listed[0]).toMatchObject({
      type: 'insertion',
      kind: 'ins',
      author: 'A',
      text: 'link text',
    })
    const result = await apply(source, out, { op: 'accept_changes', all: true })
    expect(result.code, result.stderr || result.stdout).toBe(0)
    expect(await revisions(out)).toEqual([])
    const xml = await JSZip.loadAsync(readFileSync(out)).then((z) =>
      z.file('word/document.xml')!.async('string'),
    )
    expect(xml).not.toMatch(/<w:(?:ins|del)\b/)
    expect(xml).toContain('link text')
    expect(xml).toContain('<w:hyperlink')
  })
})
