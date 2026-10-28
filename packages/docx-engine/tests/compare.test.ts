import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { compareDocuments, parseDocx, type CompareOptions } from '../src/index'
import { applyTrackedChangesXml, countTrackedChangesXml } from '../src/revision-xml'
import { buildDocx } from './helpers/build-docx'

/** the paragraph both documents share, unchanged */
const UNCHANGED_P = '<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>'
/** the paragraph the revised document edited in the middle */
const CHANGED_P = '<w:p><w:r><w:t>The quick brown fox jumps.</w:t></w:r></w:p>'
const CHANGED_EDITED_P = '<w:p><w:r><w:t>The quick red fox jumps.</w:t></w:r></w:p>'
/** a paragraph only the original has */
const REMOVED_P = '<w:p><w:r><w:t>This paragraph was removed entirely.</w:t></w:r></w:p>'
/** a paragraph only the revised document has */
const ADDED_P = '<w:p><w:r><w:t>A brand new paragraph appears here.</w:t></w:r></w:p>'

const ORIGINAL_BODY = UNCHANGED_P + CHANGED_P + REMOVED_P
const REVISED_BODY = UNCHANGED_P + CHANGED_EDITED_P + ADDED_P

const AUTHOR = 'Opposing Counsel'
const DATE = '2026-10-16T00:00:00Z'

async function documentXml(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  return zip.file('word/document.xml')!.async('string')
}

/** the revision ids on the w:ins / w:del elements of a document */
function revisionIds(xml: string): string[] {
  return [...xml.matchAll(/<w:(?:ins|del)\b[^>]*\bw:id="(\d+)"/g)].map((m) => m[1]!)
}

async function compare(
  originalBody: string,
  revisedBody: string,
  options: CompareOptions = {},
): Promise<{
  xml: string
  bytes: Uint8Array
  summary: Awaited<ReturnType<typeof compareDocuments>>['summary']
}> {
  const original = await parseDocx(await buildDocx({ bodyXml: originalBody }))
  const revised = await parseDocx(await buildDocx({ bodyXml: revisedBody }))
  const result = await compareDocuments(original, revised, {
    author: AUTHOR,
    date: DATE,
    ...options,
  })
  return { xml: await documentXml(result.bytes), bytes: result.bytes, summary: result.summary }
}

describe('document comparison (tracked changes)', () => {
  it('marks an inserted paragraph as a tracked insertion', async () => {
    const { xml, summary } = await compare(UNCHANGED_P, UNCHANGED_P + ADDED_P)
    expect(xml).toContain(`<w:ins w:id="`)
    expect(xml).toContain(`w:author="${AUTHOR}"`)
    expect(xml).toContain(`w:date="${DATE}"`)
    expect(xml).toContain('A brand new paragraph appears here.')
    expect(summary.insertedParagraphs).toBe(1)
    expect(summary.deletedParagraphs).toBe(0)
    expect(summary.changedParagraphs).toBe(0)
  })

  it('marks a wholly inserted paragraph on the runs and the paragraph mark, not a body wrapper', async () => {
    const { xml, summary } = await compare(UNCHANGED_P, UNCHANGED_P + ADDED_P)
    // w:ins/w:del are not resolved as tracked changes around a w:p — Word's
    // Compare marks the paragraph mark instead (w:pPr/w:rPr)
    expect(xml).not.toMatch(/<w:ins\b[^>]*>\s*<w:p[\s/>]/)
    expect(xml).toContain('<w:pPr><w:rPr><w:ins ')
    expect(xml).toMatch(
      new RegExp(`<w:ins w:id="\\d+" w:author="${AUTHOR}" w:date="${DATE}"/>`),
    )
    // the text runs still carry their own insertion
    expect(xml).toMatch(/<w:ins\b[^>]*>\s*<w:r>\s*<w:t(?:\s[^>]*)?>A brand new paragraph/)
    expect(summary.insertedParagraphs).toBe(1)
  })

  it('marks a wholly deleted paragraph on the runs and the paragraph mark, not a body wrapper', async () => {
    const { xml, summary } = await compare(UNCHANGED_P + REMOVED_P, UNCHANGED_P)
    expect(xml).not.toMatch(/<w:del\b[^>]*>\s*<w:p[\s/>]/)
    expect(xml).toContain('<w:pPr><w:rPr><w:del ')
    expect(xml).toContain('<w:delText')
    expect(summary.deletedParagraphs).toBe(1)
  })

  it('a non-paragraph block keeps the block-level wrapper (the shape Word writes for tables)', async () => {
    const tableA =
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>old cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    const tableB =
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>new cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    const { xml } = await compare(UNCHANGED_P + tableA, UNCHANGED_P + tableB)
    expect(xml).toMatch(/<w:del\b[^>]*>\s*<w:tbl[\s>]/)
    expect(xml).toMatch(/<w:ins\b[^>]*>\s*<w:tbl[\s>]/)
  })

  it('accepting every change resolves the redline to the revised document, rejecting to the original', async () => {
    const { bytes, summary } = await compare(ORIGINAL_BODY, REVISED_BODY)
    expect(summary.changedParagraphs).toBe(1)
    expect(summary.insertedParagraphs).toBe(1)
    expect(summary.deletedParagraphs).toBe(1)
    const zip = await JSZip.loadAsync(bytes)
    const xml = zip.file('word/document.xml')!.async('string')
    for (const [mode, expected] of [
      ['accept', ['The quick red fox jumps.', 'A brand new paragraph appears here.']],
      ['reject', ['The quick brown fox jumps.', 'This paragraph was removed entirely.']],
    ] as const) {
      const out = applyTrackedChangesXml(await xml, mode)
      expect(out.remaining).toBe(0)
      expect(out.applied).toBeGreaterThan(0)
      const text = [...out.xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join('')
      expect(text).toContain(expected[0])
      expect(text).toContain(expected[1])
      expect(text).not.toContain(mode === 'accept' ? 'brown' : 'red')
      expect(text).not.toContain(mode === 'accept' ? 'removed entirely' : 'brand new')
    }
    expect(countTrackedChangesXml(await xml)).toBeGreaterThan(0)
  })

  it('marks a deleted paragraph as a tracked deletion', async () => {
    const { xml, summary } = await compare(UNCHANGED_P + REMOVED_P, UNCHANGED_P)
    expect(xml).toContain(`<w:del w:id="`)
    expect(xml).toContain(`w:author="${AUTHOR}"`)
    expect(xml).toContain('This paragraph was removed entirely.')
    expect(xml).toContain('<w:delText')
    expect(summary.deletedParagraphs).toBe(1)
    expect(summary.insertedParagraphs).toBe(0)
  })

  it('diffs a changed paragraph word by word', async () => {
    const { xml, summary } = await compare(UNCHANGED_P + CHANGED_P, UNCHANGED_P + CHANGED_EDITED_P)
    // the untouched words stay plain, the edited word is deleted and re-inserted
    expect(xml).toContain('The quick ')
    expect(xml).toContain('brown</w:delText>')
    expect(xml).toContain('>red</w:t>')
    expect(xml).toMatch(/<w:del\b[^>]*>.*brown.*<\/w:del>/s)
    expect(xml).toMatch(/<w:ins\b[^>]*>.*red.*<\/w:ins>/s)
    expect(summary.changedParagraphs).toBe(1)
    expect(summary.insertions).toBe(1)
    expect(summary.deletions).toBe(1)
  })

  it('records the author and the date on every revision and keeps the ids unique', async () => {
    const { xml } = await compare(ORIGINAL_BODY, REVISED_BODY)
    const revisions = [...xml.matchAll(/<w:(ins|del)\b[^>]*>/g)].map((m) => m[0]!)
    expect(revisions.length).toBeGreaterThan(0)
    for (const tag of revisions) {
      expect(tag).toContain(`w:author="${AUTHOR}"`)
      expect(tag).toContain(`w:date="${DATE}"`)
    }
    const ids = revisionIds(xml)
    expect(ids.length).toBe(revisions.length)
    expect(new Set(ids).size).toBe(ids.length)
  })

  it('leaves the unchanged paragraphs byte-identical', async () => {
    const revised = await buildDocx({ bodyXml: REVISED_BODY })
    const { xml } = await compare(ORIGINAL_BODY, REVISED_BODY)
    expect(xml).toContain(UNCHANGED_P)
    const original = await parseDocx(revised)
    const untouched = original.blocks.find((b) => b.runs?.[0]?.text === 'Hello world')
    expect(untouched).toBeDefined()
  })

  it('ignores a case-only difference when asked to', async () => {
    const lower = '<w:p><w:r><w:t>the quick brown fox</w:t></w:r></w:p>'
    const upper = '<w:p><w:r><w:t>The quick brown fox</w:t></w:r></w:p>'
    const without = await compare(lower, upper)
    expect(without.summary.changedParagraphs).toBe(1)
    const with_ = await compare(lower, upper, { ignore: ['case'] })
    expect(with_.summary.changedParagraphs).toBe(0)
    expect(with_.xml).not.toContain('<w:ins')
    expect(with_.xml).not.toContain('<w:del')
  })

  it('ignores a whitespace-only difference when asked to', async () => {
    const spaced = '<w:p><w:r><w:t xml:space="preserve">the quick  fox</w:t></w:r></w:p>'
    const tight = '<w:p><w:r><w:t xml:space="preserve">the quick fox</w:t></w:r></w:p>'
    const without = await compare(spaced, tight)
    expect(without.summary.changedParagraphs).toBe(1)
    const with_ = await compare(spaced, tight, { ignore: ['whitespace'] })
    expect(with_.summary.changedParagraphs).toBe(0)
    expect(with_.xml).not.toContain('<w:ins')
    expect(with_.xml).not.toContain('<w:del')
  })

  it('reparses the redline into the revised text plus the revision marks', async () => {
    const original = await parseDocx(await buildDocx({ bodyXml: ORIGINAL_BODY }))
    const revised = await parseDocx(await buildDocx({ bodyXml: REVISED_BODY }))
    const result = await compareDocuments(original, revised, { author: AUTHOR, date: DATE })
    const again = await parseDocx(result.bytes)

    // the text a reader sees is the revised text
    const visible = again.blocks
      .filter((b) => b.runs)
      .map((b) =>
        b
          .runs!.filter((r) => !r.del)
          .map((r) => r.text)
          .join(''),
      )
    expect(visible).toContain('The quick red fox jumps.')
    expect(visible).toContain('A brand new paragraph appears here.')

    // the deleted words come back with their author
    const deleted = again.blocks.flatMap((b) => b.runs ?? []).filter((r) => r.del)
    expect(deleted.map((r) => r.text)).toContain('brown')
    expect(deleted[0]!.del).toMatchObject({ author: AUTHOR, date: DATE })
  })
})

const INLINE_PICTURE =
  `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
  `<wp:extent cx="914400" cy="914400"/><wp:docPr id="2" name="Picture 2"/>` +
  `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
  `<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Picture 2"/><pic:cNvPicPr/></pic:nvPicPr>` +
  `<pic:blipFill><a:blip r:embed="rId10"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
  `</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`

async function compareWithPictures(
  originalBody: string,
  revisedBody: string,
): Promise<{ xml: string; summary: Awaited<ReturnType<typeof compareDocuments>>['summary'] }> {
  const original = await parseDocx(await buildDocx({ bodyXml: originalBody, withImage: true }))
  const revised = await parseDocx(await buildDocx({ bodyXml: revisedBody, withImage: true }))
  const result = await compareDocuments(original, revised, { author: AUTHOR, date: DATE })
  return { xml: await documentXml(result.bytes), summary: result.summary }
}

describe('document comparison (pictures, pending changes, ids)', () => {
  it('keeps the inline picture of a paragraph whose text changed', async () => {
    const before = `<w:p><w:r><w:t xml:space="preserve">Figure one </w:t></w:r>${INLINE_PICTURE}</w:p>`
    const after = `<w:p><w:r><w:t xml:space="preserve">Figure two </w:t></w:r>${INLINE_PICTURE}</w:p>`
    const { xml, summary } = await compareWithPictures(before, after)
    expect(summary.changedParagraphs).toBe(1)
    expect(xml).toContain('r:embed="rId10"')
    expect(xml).not.toContain('￼')
    expect(xml).toMatch(/<w:del\b[^>]*>[\s\S]*?<w:delText[^>]*>one<\/w:delText>/)
  })

  it('names a deleted picture of the original instead of carrying its media reference', async () => {
    const pictureP = `<w:p>${INLINE_PICTURE}</w:p>`
    const { xml, summary } = await compareWithPictures(UNCHANGED_P + pictureP, UNCHANGED_P)
    expect(summary.deletedParagraphs).toBe(1)
    expect(xml).not.toContain('r:embed')
    expect(xml).toContain('[picture]')
  })

  it('marks a picture replaced by a different picture, and keeps an identical one', async () => {
    const second = INLINE_PICTURE.replace('rId10', 'rId11')
    const original = await parseDocx(
      await buildDocx({ bodyXml: UNCHANGED_P + `<w:p>${INLINE_PICTURE}</w:p>`, withImage: true }),
    )
    const revised = await parseDocx(
      await buildDocx({
        bodyXml: UNCHANGED_P + `<w:p>${second}</w:p>`,
        withImage: true,
        extraRels:
          '<Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/other.png"/>',
        binaryParts: [
          {
            path: 'word/media/other.png',
            base64:
              'iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAEklEQVR4nGP4z8DwHwyBNAMDACfkA/0W4kQ3AAAAAElFTkSuQmCC',
            extension: 'png',
            contentType: 'image/png',
          },
        ],
      }),
    )
    const replaced = await compareDocuments(original, revised, { author: AUTHOR, date: DATE })
    expect(replaced.summary.deletedParagraphs).toBe(1)
    expect(replaced.summary.insertedParagraphs).toBe(1)
    const same = await compareWithPictures(
      UNCHANGED_P + `<w:p>${INLINE_PICTURE}</w:p>`,
      UNCHANGED_P + `<w:p>${INLINE_PICTURE}</w:p>`,
    )
    expect(same.summary.unchangedParagraphs).toBe(2)
    expect(same.xml).not.toContain('<w:del ')
  })

  it('compares the documents with their pending changes accepted', async () => {
    const pending =
      '<w:p><w:r><w:t xml:space="preserve">The quick </w:t></w:r>' +
      `<w:del w:id="5" w:author="A" w:date="${DATE}"><w:r><w:delText xml:space="preserve">brown </w:delText></w:r></w:del>` +
      `<w:ins w:id="6" w:author="A" w:date="${DATE}"><w:r><w:t xml:space="preserve">red </w:t></w:r></w:ins>` +
      '<w:r><w:t>fox jumps.</w:t></w:r></w:p>'
    const { summary } = await compare(pending, CHANGED_EDITED_P)
    expect(summary.changedParagraphs).toBe(0)
    expect(summary.unchangedParagraphs).toBe(1)
  })

  it('keeps every revision id unique when a deletion spans several runs', async () => {
    const split =
      '<w:p><w:r><w:t xml:space="preserve">The quick </w:t></w:r>' +
      '<w:r><w:t>br</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">own </w:t></w:r>' +
      '<w:r><w:t>fox jumps.</w:t></w:r></w:p>'
    const { xml } = await compare(split + REMOVED_P, CHANGED_EDITED_P + ADDED_P)
    const ids = revisionIds(xml)
    expect(ids.length).toBeGreaterThan(2)
    expect(new Set(ids).size).toBe(ids.length)
  })

  it('passes an unchanged table through instead of marking it replaced', async () => {
    const table =
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    const { xml, summary } = await compare(UNCHANGED_P + table, UNCHANGED_P + table)
    expect(summary.unchangedParagraphs).toBe(2)
    expect(summary.deletedParagraphs).toBe(0)
    expect(summary.insertedParagraphs).toBe(0)
    expect(xml).not.toContain('<w:del ')
    expect(xml).not.toContain('<w:ins ')
  })

  it('keeps the content-control shell balanced around changed and deleted paragraphs', async () => {
    const sdt = (second: string) =>
      '<w:sdt><w:sdtPr><w:tag w:val="example:block"/><w:id w:val="7"/></w:sdtPr><w:sdtContent>' +
      '<w:p><w:r><w:t>Caption one</w:t></w:r></w:p>' +
      `<w:p><w:r><w:t>${second}</w:t></w:r></w:p>` +
      '</w:sdtContent></w:sdt>'
    const changed = await compare(
      UNCHANGED_P + sdt('Signature block'),
      UNCHANGED_P + sdt('Signature line'),
    )
    expect(changed.summary.changedParagraphs).toBe(1)
    expect((changed.xml.match(/<w:sdt>/g) ?? []).length).toBe(1)
    expect((changed.xml.match(/<\/w:sdt>/g) ?? []).length).toBe(1)
    expect(changed.xml.indexOf('<w:sdt>')).toBeLessThan(changed.xml.indexOf('<w:del '))
    const removed = await compare(UNCHANGED_P + sdt('Signature block'), UNCHANGED_P)
    expect((removed.xml.match(/<w:sdt>/g) ?? []).length).toBe(1)
    expect((removed.xml.match(/<\/w:sdt>/g) ?? []).length).toBe(1)
    expect(removed.xml).not.toMatch(/<w:del\b[^>]*><w:sdt>/)
    await expect(parseDocx(removed.bytes)).resolves.toBeTruthy()
  })

  it('anchors on unique paragraphs so repeated separators do not pair across sections', async () => {
    const sep = '<w:p><w:r><w:t>---</w:t></w:r></w:p>'
    const a = '<w:p><w:r><w:t>Section A text.</w:t></w:r></w:p>'
    const b = '<w:p><w:r><w:t>Section B text.</w:t></w:r></w:p>'
    const { summary } = await compare(a + sep + b + sep, sep + a + sep + b + sep)
    expect(summary.unchangedParagraphs).toBe(4)
    expect(summary.insertedParagraphs).toBe(1)
    expect(summary.deletedParagraphs).toBe(0)
  })
})
