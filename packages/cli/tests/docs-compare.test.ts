import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { run, tempDir } from './helpers'

const DOC_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

/** a minimal but valid .docx holding the given body paragraphs */
async function buildDocx(bodyXml: string): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'word/styles.xml',
    `${XML_DECL}<w:styles ${DOC_NS}>` +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
      '</w:styles>',
  )
  zip.file(
    'word/document.xml',
    `${XML_DECL}<w:document ${DOC_NS}><w:body>${bodyXml}<w:sectPr/></w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'uint8array' })
}

const P = (text: string): string => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`

const UNCHANGED = P('Hello world')
const ORIGINAL_BODY =
  UNCHANGED + P('The quick brown fox jumps.') + P('This paragraph was removed entirely.')
const REVISED_BODY =
  UNCHANGED + P('The quick red fox jumps.') + P('A brand new paragraph appears here.')

async function documentXml(path: string): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(path))
  return zip.file('word/document.xml')!.async('string')
}

/** writes the two fixtures into a fresh folder and returns their paths */
async function fixtures(): Promise<{ dir: string; original: string; revised: string }> {
  const dir = tempDir()
  const original = join(dir, 'original.docx')
  const revised = join(dir, 'revised.docx')
  writeFileSync(original, await buildDocx(ORIGINAL_BODY))
  writeFileSync(revised, await buildDocx(REVISED_BODY))
  return { dir, original, revised }
}

describe('genoffice docs compare', () => {
  it('writes a redline with the differences marked as tracked changes', async () => {
    const { dir, original, revised } = await fixtures()
    const out = join(dir, 'redline.docx')
    const r = await run(['docs', 'compare', original, revised, '--out', out, '--json'])
    expect(r.code).toBe(0)

    const xml = await documentXml(out)
    expect(xml).toContain('<w:ins ')
    expect(xml).toContain('<w:del ')
    expect(xml).toContain('brown</w:delText>')
    expect(xml).toContain('>red</w:t>')
    // the redline is a document the editor opens, and it sees the pending changes
    const read = await run(['docs', 'read', out, '--revisions', '--json'])
    expect(read.code).toBe(0)
    expect(
      read
        .json()
        .detail.items.map((i: { text: string }) => i.text)
        .join(' '),
    ).toContain('red fox jumps.')
    const revisions = read.json().detail.revisions as { author: string }[]
    expect(revisions.length).toBeGreaterThan(0)
    expect(revisions.every((r) => r.author === 'GenOffice')).toBe(true)
  })

  it('reports the counts and the output path as JSON', async () => {
    const { dir, original, revised } = await fixtures()
    const out = join(dir, 'redline.docx')
    const r = await run(['docs', 'compare', original, revised, '--out', out, '--json'])
    expect(r.code).toBe(0)
    const payload = r.json()
    expect(payload.output_path).toBe(out)
    expect(payload.detail).toMatchObject({
      insertions: 2,
      deletions: 2,
      inserted_paragraphs: 1,
      deleted_paragraphs: 1,
      changed_paragraphs: 1,
      unchanged_paragraphs: 1,
    })
  })

  it('records the author and the date on every revision', async () => {
    const { dir, original, revised } = await fixtures()
    const out = join(dir, 'redline.docx')
    const r = await run([
      'docs',
      'compare',
      original,
      revised,
      '--out',
      out,
      '--author',
      'Opposing Counsel',
      '--date',
      '2026-10-16T00:00:00Z',
      '--json',
    ])
    expect(r.code).toBe(0)
    const xml = await documentXml(out)
    const revisions = [...xml.matchAll(/<w:(ins|del)\b[^>]*>/g)].map((m) => m[0]!)
    expect(revisions.length).toBeGreaterThan(0)
    for (const tag of revisions) {
      expect(tag).toContain('w:author="Opposing Counsel"')
      expect(tag).toContain('w:date="2026-10-16T00:00:00Z"')
    }
    const ids = [...xml.matchAll(/<w:(?:ins|del)\b[^>]*\bw:id="(\d+)"/g)].map((m) => m[1]!)
    expect(new Set(ids).size).toBe(ids.length)
  })

  it('diffs a changed paragraph at word level and at character level', async () => {
    const { dir } = await fixtures()
    const original = join(dir, 'word-original.docx')
    const revised = join(dir, 'word-revised.docx')
    writeFileSync(original, await buildDocx(P('Hello world')))
    writeFileSync(revised, await buildDocx(P('Hello worl')))

    const wordOut = join(dir, 'word.docx')
    const word = await run(['docs', 'compare', original, revised, '--out', wordOut, '--json'])
    expect(word.code).toBe(0)
    expect(await documentXml(wordOut)).toContain('world</w:delText>')

    const charOut = join(dir, 'char.docx')
    const char = await run([
      'docs',
      'compare',
      original,
      revised,
      '--out',
      charOut,
      '--granularity',
      'char',
      '--json',
    ])
    expect(char.code).toBe(0)
    const charXml = await documentXml(charOut)
    expect(charXml).toContain('d</w:delText>')
    expect(charXml).not.toContain('world</w:delText>')
  })

  it('ignores a case-only or whitespace-only difference when asked to', async () => {
    const { dir } = await fixtures()
    const original = join(dir, 'case-original.docx')
    const revised = join(dir, 'case-revised.docx')
    writeFileSync(original, await buildDocx(P('the quick fox')))
    writeFileSync(revised, await buildDocx(P('The quick  fox')))

    const plain = join(dir, 'plain.docx')
    const before = await run(['docs', 'compare', original, revised, '--out', plain, '--json'])
    expect(before.json().detail.changed_paragraphs).toBe(1)

    const ignored = join(dir, 'ignored.docx')
    const after = await run([
      'docs',
      'compare',
      original,
      revised,
      '--out',
      ignored,
      '--ignore',
      'case,whitespace',
      '--json',
    ])
    expect(after.code).toBe(0)
    expect(after.json().detail.changed_paragraphs).toBe(0)
    expect(await documentXml(ignored)).not.toContain('<w:ins ')
  })

  it('refuses a missing --out, an unknown granularity and an existing output file', async () => {
    const { dir, original, revised } = await fixtures()
    const noOut = await run(['docs', 'compare', original, revised, '--json'])
    expect(noOut.code).toBe(1)
    expect(noOut.json().error).toBe('missing_argument')

    const badGranularity = await run([
      'docs',
      'compare',
      original,
      revised,
      '--out',
      join(dir, 'x.docx'),
      '--granularity',
      'sentence',
      '--json',
    ])
    expect(badGranularity.code).toBe(1)
    expect(badGranularity.json().error).toBe('invalid_argument')

    const out = join(dir, 'taken.docx')
    writeFileSync(out, 'x')
    const exists = await run(['docs', 'compare', original, revised, '--out', out, '--json'])
    expect(exists.code).toBe(2)
    expect(exists.json().error).toBe('output_exists')
  })
})
