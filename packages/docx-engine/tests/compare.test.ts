import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { compareDocuments, parseDocx, type CompareOptions } from '../src/index'
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
