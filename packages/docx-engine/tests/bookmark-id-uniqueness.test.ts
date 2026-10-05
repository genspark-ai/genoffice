/**
 * w:bookmarkStart/@w:id is unique within the part. The id used to be a 31-bit
 * hash of the bookmark name, which both overlaps the small ids Word itself hands
 * out and collides between two new names, so a rebuilt document could hold two
 * bookmarks under one id — after which Word mis-pairs them and a cross-reference
 * lands on the wrong target.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { generateParagraphXml, parseDocx, saveDocx, type GenerateContext } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const TEXT_PARA = '<w:p><w:r><w:t>body</w:t></w:r></w:p>'

const GEN_CTX: GenerateContext = {
  headingStyleIds: new Map([[1, 'Heading1']]),
  allocateHyperlinkRel: () => 'rId999',
}

/** a paragraph holding one already-present bookmark */
const existingBookmark = (id: string, name: string) =>
  `<w:p><w:bookmarkStart w:id=${id} w:name="${name}"/><w:bookmarkEnd w:id=${id}/>` +
  '<w:r><w:t>original</w:t></w:r></w:p>'

/** every bookmarkStart id in the saved document, in order */
async function bookmarkIdsOf(bytes: Uint8Array): Promise<number[]> {
  const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
  return [...xml.matchAll(/<w:bookmarkStart\s[^>]*?\bid="(\d+)"/g)].map((m) => Number(m[1]))
}

/** save a document whose single original paragraph is kept, followed by one generated paragraph per group */
async function saveAddingBookmarks(bodyXml: string, groups: string[][]): Promise<number[]> {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const saved = await saveDocx(parsed, [
    { kind: 'original', docxIndex: 0 },
    ...groups.map((bookmarks) => ({
      kind: 'generated' as const,
      block: { type: 'paragraph' as const, bookmarks, runs: [{ text: 'added' }] },
    })),
  ])
  return bookmarkIdsOf(saved)
}

describe('minted bookmark ids', () => {
  // the old hash of the name "a" is 97, so a document already holding id 97 got
  // a second bookmarkStart id 97 on the next save
  it('never mints an id the document already uses', async () => {
    const ids = await saveAddingBookmarks(existingBookmark('"97"', 'existing'), [['a']])
    expect(ids).toEqual([97, 98])
  })

  it('starts above the highest id in the part, however it is quoted', async () => {
    const ids = await saveAddingBookmarks(
      existingBookmark('"40000"', 'high') + existingBookmark("'9000'", 'mid'),
      [['a', 'b', 'c']],
    )
    expect(new Set(ids).size).toBe(ids.length)
    expect(Math.min(...ids.slice(2))).toBeGreaterThan(40000)
  })

  it('keeps ids unique among themselves at a realistic bookmark count', async () => {
    // 2000 short names: the old hash collapsed them onto 715 distinct ids, so a
    // duplicate would be all but guaranteed. Spread over 20 paragraphs, which
    // also catches an allocator that restarts per paragraph.
    const names: string[] = []
    for (let a = 0x20; a < 0x7f && names.length < 2000; a++) {
      for (let b = 0x20; b < 0x7f && names.length < 2000; b++) {
        names.push(`sec_${String.fromCharCode(a)}${String.fromCharCode(b)}`)
      }
    }
    const groups = Array.from({ length: 20 }, (_, i) => names.slice(i * 100, i * 100 + 100))
    const ids = await saveAddingBookmarks(TEXT_PARA, groups)
    expect(ids).toHaveLength(2000)
    expect(new Set(ids).size).toBe(2000)
  })

  it('pairs each start with its own end, and keeps ids unique without a document', () => {
    const xml = generateParagraphXml(
      { type: 'paragraph', bookmarks: ['章节A', '章节B'], runs: [{ text: 'x' }] },
      GEN_CTX,
    )
    const starts = [...xml.matchAll(/<w:bookmarkStart w:id="(\d+)" w:name="([^"]+)"\/>/g)]
    const ends = [...xml.matchAll(/<w:bookmarkEnd w:id="(\d+)"\/>/g)].map((m) => m[1])
    expect(starts.map((m) => m[2])).toEqual(['章节A', '章节B'])
    expect(ends).toEqual(starts.map((m) => m[1]))
    expect(new Set(ends).size).toBe(2)
  })
})
