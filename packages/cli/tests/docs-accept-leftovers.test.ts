/**
 * genoffice#2337 (move-range half): genoffice docs apply accept_changes / reject_changes
 * {all:true} used to leave w:moveFromRangeEnd / w:moveToRangeEnd behind in the
 * paragraphs Word wrote them into (the following, untouched paragraphs). The
 * fixture is the issue's synthetic document (JSZip + XML strings: python-docx
 * is not available in CI).
 *
 * The section-property half (the body w:sectPrChange record) is genoffice#2344 and
 * is deliberately not covered here.
 */
import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { run, tempDir } from './helpers'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

const ISSUE_BODY =
  '<w:p><w:r><w:t>P0 intro</w:t></w:r></w:p>' +
  '<w:p><w:moveFromRangeStart w:id="1" w:author="X" w:date="2026-01-01T00:00:00Z" w:name="move1"/>' +
  '<w:moveFrom w:id="2" w:author="X" w:date="2026-01-01T00:00:00Z"><w:r><w:t xml:space="preserve">P1 moved paragraph</w:t></w:r></w:moveFrom></w:p>' +
  '<w:p><w:moveFromRangeEnd w:id="1"/><w:r><w:t xml:space="preserve">P2 stays</w:t></w:r></w:p>' +
  '<w:p><w:moveToRangeStart w:id="3" w:author="X" w:date="2026-01-01T00:00:00Z" w:name="move1"/>' +
  '<w:moveTo w:id="4" w:author="X" w:date="2026-01-01T00:00:00Z"><w:r><w:t xml:space="preserve">P1 moved paragraph</w:t></w:r></w:moveTo></w:p>' +
  '<w:p><w:moveToRangeEnd w:id="3"/><w:r><w:t xml:space="preserve">P4 tail</w:t></w:r></w:p>'

/** a still-pending move: both range markers ride through an unrelated edit */
const PENDING_BODY =
  '<w:p><w:r><w:t>Edit target paragraph</w:t></w:r></w:p>' +
  '<w:p><w:moveFromRangeStart w:id="1" w:author="X" w:date="2026-01-01T00:00:00Z" w:name="move1"/>' +
  '<w:moveFrom w:id="2" w:author="X" w:date="2026-01-01T00:00:00Z"><w:r><w:t>pending moved</w:t></w:r></w:moveFrom></w:p>' +
  '<w:p><w:moveFromRangeEnd w:id="1"/><w:r><w:t>after pending</w:t></w:r></w:p>'

function documentXml(body: string): string {
  return (
    XML_DECL +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
    body +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720" w:gutter="0"/>' +
    '</w:sectPr></w:body></w:document>'
  )
}

async function buildFixture(path: string, body: string): Promise<string> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    XML_DECL +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    XML_DECL +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>',
  )
  zip.file('word/document.xml', documentXml(body))
  zip.forEach((_path, entry) => {
    entry.date = new Date(Date.UTC(2026, 0, 1))
  })
  writeFileSync(path, await zip.generateAsync({ type: 'nodebuffer' }))
  return path
}

async function documentXmlOf(path: string): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(path))
  return zip.file('word/document.xml')!.async('string')
}

const MOVE_MARKUP = /<w:move(?:From|To)?[A-Za-z]*[\s/>]/

function writeOps(dir: string, ops: unknown[]): string {
  const opsPath = join(dir, 'ops.json')
  writeFileSync(opsPath, JSON.stringify(ops))
  return opsPath
}

describe('genoffice docs apply accept/reject move-range leftovers', () => {
  it('accept all removes the cross-paragraph move range ends', async () => {
    const dir = tempDir()
    const file = await buildFixture(join(dir, 'accept.docx'), ISSUE_BODY)
    const out = join(dir, 'accepted.docx')
    const ops = writeOps(dir, [{ op: 'accept_changes', all: true }])
    const result = await run(['docs', 'apply', file, '--ops', ops, '--out', out, '--json'])
    expect(result.code).toBe(0)
    const output = result.json().detail.results[0].output as string
    expect(output).toContain('Accepted 2 tracked change(s) (r1, r2)')
    expect(output).toContain('No tracked changes remain.')
    const xml = await documentXmlOf(out)
    expect(xml).not.toMatch(MOVE_MARKUP)
    // moved text survives once, at its destination; the neighbours survive
    expect(xml.match(/P1 moved paragraph/g)).toHaveLength(1)
    expect(xml).toContain('P2 stays')
    expect(xml).toContain('P4 tail')
  })

  it('reject all removes the move markup and keeps the text at its source', async () => {
    const dir = tempDir()
    const file = await buildFixture(join(dir, 'reject.docx'), ISSUE_BODY)
    const out = join(dir, 'rejected.docx')
    const ops = writeOps(dir, [{ op: 'reject_changes', all: true }])
    const result = await run(['docs', 'apply', file, '--ops', ops, '--out', out, '--json'])
    expect(result.code).toBe(0)
    const output = result.json().detail.results[0].output as string
    expect(output).toContain('Rejected 2 tracked change(s) (r1, r2)')
    expect(output).toContain('No tracked changes remain.')
    const xml = await documentXmlOf(out)
    expect(xml).not.toMatch(MOVE_MARKUP)
    expect(xml.match(/P1 moved paragraph/g)).toHaveLength(1)
    expect(xml).toContain('P2 stays')
    expect(xml).toContain('P4 tail')
  })

  it('keeps a still-pending move range when an unrelated edit saves', async () => {
    const dir = tempDir()
    const file = await buildFixture(join(dir, 'pending.docx'), PENDING_BODY)
    const out = join(dir, 'edited.docx')
    const ops = writeOps(dir, [
      { op: 'findReplace', find: 'Edit target', replace: 'EDITED target' },
    ])
    const result = await run(['docs', 'apply', file, '--ops', ops, '--out', out, '--json'])
    expect(result.code).toBe(0)
    const xml = await documentXmlOf(out)
    expect(xml).toContain('EDITED target')
    expect(xml).toContain('<w:moveFromRangeStart w:id="1"')
    expect(xml).toContain('<w:moveFromRangeEnd w:id="1"/>')
  })
})

/**
 * genoffice#2555: a w:ins / w:del whose only run holds an atom (a footnote
 * reference, a line break) used to be invisible to the revision list, so
 * accept_changes left the wrapper in the file — and a document whose only
 * revision was such a wrapper was rejected outright as having no pending
 * changes. genoffice#2554 (case 3): the same rejection hit files whose only
 * markup is orphan move-range markers, which the save already strips
 * (genoffice#2349) — accept / reject now succeed and let the save clean them.
 */
describe('accept/reject wrappers around atomic runs and orphan-only files', () => {
  const FN_XML =
    XML_DECL +
    '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
    '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
    '<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Note.</w:t></w:r></w:p></w:footnote></w:footnotes>'
  const REV = 'w:author="A" w:date="2026-01-01T00:00:00Z"'

  async function buildFnFixture(path: string, body: string): Promise<string> {
    const zip = new JSZip()
    zip.file(
      '[Content_Types].xml',
      XML_DECL +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>' +
        '</Types>',
    )
    zip.file(
      '_rels/.rels',
      XML_DECL +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
        '</Relationships>',
    )
    zip.file('word/document.xml', documentXml(body))
    zip.file('word/footnotes.xml', FN_XML)
    zip.file(
      'word/_rels/document.xml.rels',
      XML_DECL +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>' +
        '</Relationships>',
    )
    zip.forEach((_path, entry) => {
      entry.date = new Date(Date.UTC(2026, 0, 1))
    })
    writeFileSync(path, await zip.generateAsync({ type: 'nodebuffer' }))
    return path
  }

  const FNREF_INS = `<w:p><w:r><w:t>Text.</w:t></w:r><w:ins w:id="1" ${REV}><w:r><w:footnoteReference w:id="1"/></w:r></w:ins><w:ins w:id="9" ${REV}><w:r><w:t> other change</w:t></w:r></w:ins></w:p>`
  const FNREF_INS_ONLY = `<w:p><w:r><w:t>Text.</w:t></w:r><w:ins w:id="1" ${REV}><w:r><w:footnoteReference w:id="1"/></w:r></w:ins></w:p>`
  const FNREF_DEL = `<w:p><w:r><w:t>Text.</w:t></w:r><w:del w:id="1" ${REV}><w:r><w:footnoteReference w:id="1"/></w:r></w:del><w:ins w:id="9" ${REV}><w:r><w:t> other change</w:t></w:r></w:ins></w:p>`

  it('accept all clears a wrapper that only holds a footnote reference', async () => {
    const dir = tempDir()
    const file = await buildFnFixture(join(dir, 'fnref-ins.docx'), FNREF_INS)
    const out = join(dir, 'accepted.docx')
    const ops = writeOps(dir, [{ op: 'accept_changes', all: true }])
    const result = await run(['docs', 'apply', file, '--ops', ops, '--out', out, '--json'])
    expect(result.code).toBe(0)
    const xml = await documentXmlOf(out)
    expect(xml).not.toMatch(/<w:(?:ins|del) [^>]*w:id="/)
    // the accepted insertion keeps its marker, now unwrapped
    expect(xml).toContain('<w:footnoteReference w:id="1"/>')
    expect(xml).toContain('other change')
  })

  it('a file whose only revision is such a wrapper is accepted, not rejected', async () => {
    const dir = tempDir()
    const file = await buildFnFixture(join(dir, 'fnref-only.docx'), FNREF_INS_ONLY)
    const out = join(dir, 'accepted.docx')
    const ops = writeOps(dir, [{ op: 'accept_changes', all: true }])
    const result = await run(['docs', 'apply', file, '--ops', ops, '--out', out, '--json'])
    expect(result.code).toBe(0)
    const output = result.json().detail.results[0].output as string
    expect(output).toContain('No tracked changes remain.')
    const xml = await documentXmlOf(out)
    expect(xml).not.toMatch(/<w:(?:ins|del) [^>]*w:id="/)
    expect(xml).toContain('<w:footnoteReference w:id="1"/>')
  })

  it('accepting a deleted footnote reference removes it; the wrapper is gone', async () => {
    const dir = tempDir()
    const file = await buildFnFixture(join(dir, 'fnref-del.docx'), FNREF_DEL)
    const out = join(dir, 'accepted.docx')
    const ops = writeOps(dir, [{ op: 'accept_changes', all: true }])
    const result = await run(['docs', 'apply', file, '--ops', ops, '--out', out, '--json'])
    expect(result.code).toBe(0)
    const xml = await documentXmlOf(out)
    expect(xml).not.toMatch(/<w:(?:ins|del) [^>]*w:id="/)
    expect(xml).not.toContain('<w:footnoteReference')
  })

  it('accept / reject succeeds on a file whose only markup is orphan move-range markers', async () => {
    const orphanBody = '<w:p><w:moveFromRangeEnd w:id="1"/><w:r><w:t>P1</w:t></w:r></w:p>'
    for (const op of [
      { op: 'accept_changes', all: true },
      { op: 'reject_changes', all: true },
    ]) {
      const dir = tempDir()
      const file = await buildFixture(join(dir, 'orphan.docx'), orphanBody)
      const out = join(dir, 'cleaned.docx')
      const ops = writeOps(dir, [op])
      const result = await run(['docs', 'apply', file, '--ops', ops, '--out', out, '--json'])
      expect(result.code).toBe(0)
      const output = result.json().detail.results[0].output as string
      expect(output).toContain('No tracked changes remain.')
      expect(output).toContain('orphan move-range marker')
      const xml = await documentXmlOf(out)
      expect(xml).not.toMatch(MOVE_MARKUP)
      expect(xml).toContain('P1')
    }
  })
})
