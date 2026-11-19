import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { buildDocx } from '../../docx-engine/tests/helpers/build-docx'
import { run, tempDir } from './helpers'

/**
 * The last byte-confinement gaps from genoffice#524: an edit must not rewrite
 * paragraphs it did not target, and the attributes a regenerated revision
 * element carries must survive the round trip.
 *
 * - <w:proofErr> markers (Word's grammar/spell ranges) used to be dropped when
 *   an untouched heading was regenerated.
 * - w16du:dateUtc on w:ins / w:del used to be lost whenever the revision was
 *   re-emitted, even when the edit that regenerated it was elsewhere.
 */

const W16DU = 'xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du"'
const AT = 'w:author="Author A" w:date="2026-01-01T10:00:00Z" w16du:dateUtc="2026-01-01T10:00:00Z"'

const BODY =
  '<w:p><w:r><w:t>Edit target paragraph with plain text to replace.</w:t></w:r></w:p>' +
  '<w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr>' +
  '<w:proofErr w:type="gramStart"/>' +
  '<w:r><w:rPr><w:rStyle w:val="NormalTextRun"/></w:rPr><w:t>Heading</w:t></w:r>' +
  '<w:proofErr w:type="gramEnd"/>' +
  '<w:r><w:t xml:space="preserve"> tail</w:t></w:r></w:p>' +
  '<w:p>' +
  `<w:ins w:id="501" ${AT}><w:r><w:t xml:space="preserve">Inserted sentence. </w:t></w:r></w:ins>` +
  '<w:del w:id="502" w:author="Author B" w:date="2026-01-01T11:00:00Z" ' +
  'w16du:dateUtc="2026-01-01T11:00:00Z"><w:r><w:delText xml:space="preserve">' +
  'Deleted middle. </w:delText></w:r></w:del></w:p>'

async function documentXml(path: string): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(path))
  return zip.file('word/document.xml')!.async('string')
}

/** the <w:p>…</w:p> element containing the first occurrence of needle */
function paragraphWith(xml: string, needle: string): string {
  const at = xml.indexOf(needle)
  if (at < 0) return ''
  let start = -1
  for (const match of xml.matchAll(/<w:p[\s>]/g)) {
    if ((match.index ?? -1) > at) break
    start = match.index ?? -1
  }
  const end = xml.indexOf('</w:p>', at)
  return start >= 0 && end >= 0 ? xml.slice(start, end + '</w:p>'.length) : ''
}

async function makeDocx(dir: string): Promise<string> {
  const path = join(dir, 'in.docx')
  writeFileSync(path, await buildDocx({ bodyXml: BODY, docRootExtraAttrs: W16DU }))
  return path
}

async function applyOps(dir: string, input: string, ops: unknown[], name: string): Promise<string> {
  const opsPath = join(dir, `${name}.json`)
  const out = join(dir, `${name}.docx`)
  writeFileSync(opsPath, JSON.stringify(ops))
  const result = await run(['docs', 'apply', input, '--ops', opsPath, '--out', out, '--json'])
  expect(result.code, result.stderr).toBe(0)
  return out
}

it('keeps an untouched paragraph with <w:proofErr> byte-identical', async () => {
  const dir = tempDir()
  const input = await makeDocx(dir)
  const out = await applyOps(
    dir,
    input,
    [{ op: 'findReplace', find: 'plain text', replace: 'PLAIN TEXT' }],
    'fr',
  )

  const before = await documentXml(input)
  const after = await documentXml(out)
  const headingBefore = paragraphWith(before, 'Heading')
  const headingAfter = paragraphWith(after, 'Heading')
  expect(headingBefore).not.toBe('')
  expect(headingAfter).toBe(headingBefore)
  expect((after.match(/<w:proofErr\b/g) ?? []).length).toBe(2)
  expect(after).toContain('PLAIN TEXT')
})

it('keeps w16du:dateUtc on an untouched tracked paragraph and on a regenerated one', async () => {
  const dir = tempDir()
  const input = await makeDocx(dir)

  // the edit lands elsewhere: the tracked paragraph is copied through verbatim
  const out = await applyOps(
    dir,
    input,
    [{ op: 'findReplace', find: 'plain text', replace: 'PLAIN TEXT' }],
    'fr-elsewhere',
  )
  const before = await documentXml(input)
  const after = await documentXml(out)
  expect(paragraphWith(after, 'Inserted sentence')).toBe(paragraphWith(before, 'Inserted sentence'))
  expect((after.match(/w16du:dateUtc=/g) ?? []).length).toBe(2)

  // the edit lands inside the tracked paragraph: it is regenerated, and the
  // revision attributes (author/date/dateUtc) must survive
  const out2 = await applyOps(
    dir,
    input,
    [{ op: 'findReplace', find: 'Inserted sentence', replace: 'INSERTED sentence' }],
    'fr-inside',
  )
  const after2 = await documentXml(out2)
  expect(after2).toContain('INSERTED sentence')
  const ins = /<w:ins\b[^>]*>/.exec(after2)?.[0] ?? ''
  expect(ins).toContain('w:id="501"')
  expect(ins).toContain('w:author="Author A"')
  expect(ins).toContain('w:date="2026-01-01T10:00:00Z"')
  expect(ins).toContain('w16du:dateUtc="2026-01-01T10:00:00Z"')
})
