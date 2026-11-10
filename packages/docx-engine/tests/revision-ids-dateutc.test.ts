import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { generateParagraphXml, parseDocx, saveDocx, type GenerateContext } from '../src/index'
import { buildDocx } from './helpers/build-docx'

/** genoffice#524: w16du:dateUtc round trip and document-wide unique revision ids. */

const W16DU_NS = 'http://schemas.microsoft.com/office/word/2023/wordml/word16du'
const UTC = 'w16du:dateUtc="2026-01-01T10:00:00Z"'
const A = `w:author="Author A" w:date="2026-01-01T10:00:00Z" ${UTC}`
const GEN_CTX: GenerateContext = {
  headingStyleIds: new Map(),
  allocateHyperlinkRel: () => 'rId999',
}

async function documentXmlOf(bytes: Uint8Array): Promise<string> {
  return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
}

const revisionIds = (xml: string): string[] =>
  [...xml.matchAll(/<w:(?:ins|del)\b[^>]*?w:id="(\d+)"/g)].map((m) => m[1])

describe('w16du:dateUtc', () => {
  it('is captured on run, paragraph-mark, pPrChange and rPrChange revisions', async () => {
    const parsed = await parseDocx(
      await buildDocx({
        docRootExtraAttrs: `xmlns:w16du="${W16DU_NS}"`,
        bodyXml:
          `<w:p><w:pPr><w:rPr><w:del w:id="1" ${A}/></w:rPr><w:pPrChange w:id="2" ${A}><w:pPr/></w:pPrChange></w:pPr>` +
          `<w:ins w:id="3" ${A}><w:r><w:t>a</w:t></w:r></w:ins>` +
          `<w:del w:id="4" ${A}><w:r><w:delText>b</w:delText></w:r></w:del>` +
          `<w:r><w:rPr><w:rPrChange w:id="5" ${A}><w:rPr/></w:rPrChange></w:rPr><w:t>c</w:t></w:r></w:p>`,
      }),
    )
    const block = parsed.blocks[0]
    const utc = '2026-01-01T10:00:00Z'
    expect(block.paraMarkDel?.dateUtc).toBe(utc)
    expect(block.pPrChangeInfo?.dateUtc).toBe(utc)
    expect(block.runs![0].ins?.dateUtc).toBe(utc)
    expect(block.runs![1].del?.dateUtc).toBe(utc)
    expect(block.runs![2].rPrChange?.dateUtc).toBe(utc)
  })

  it('is re-emitted by generateParagraphXml', () => {
    const xml = generateParagraphXml(
      {
        type: 'paragraph',
        runs: [
          {
            text: 'a',
            ins: {
              author: 'A',
              date: '2026-01-01T10:00:00Z',
              dateUtc: '2026-01-01T10:00:00Z',
              id: '3',
            },
          },
          { text: 'c', rPrChange: { author: 'A', dateUtc: '2026-01-01T10:00:00Z', id: '5' } },
        ],
      },
      GEN_CTX,
    )
    expect(xml).toContain(`<w:ins w:id="3" w:author="A" w:date="2026-01-01T10:00:00Z" ${UTC}>`)
    expect(xml).toContain(`<w:rPrChange w:id="5" w:author="A" ${UTC}>`)
  })

  it('declares the w16du namespace on a document that lacked it', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>a</w:t></w:r></w:p>' }),
    )
    const saved = await saveDocx(parsed, [
      {
        kind: 'generated',
        block: {
          type: 'paragraph',
          runs: [{ text: 'a', ins: { author: 'A', dateUtc: '2026-01-01T10:00:00Z' } }],
        },
        docxIndex: 0,
      },
    ])
    const xml = await documentXmlOf(saved)
    expect(xml).toContain(UTC)
    expect(xml).toMatch(new RegExp(`<w:document[^>]*xmlns:w16du="${W16DU_NS}"`))
  })
})

describe('revision w:id uniqueness', () => {
  it('two regenerated paragraphs never share minted ids, seeded above existing ids', async () => {
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml:
          `<w:p><w:ins w:id="9001" w:author="A"><w:r><w:t>kept</w:t></w:r></w:ins></w:p>` +
          '<w:p><w:r><w:t>b</w:t></w:r></w:p><w:p><w:r><w:t>c</w:t></w:r></w:p>',
      }),
    )
    const tracked = (text: string) => ({
      kind: 'generated' as const,
      block: {
        type: 'paragraph' as const,
        runs: [
          { text },
          { text: 'new', ins: { author: 'R' } },
          { text: 'old', del: { author: 'R' } },
        ],
      },
    })
    const saved = await saveDocx(parsed, [
      { kind: 'original', docxIndex: 0 },
      { ...tracked('b'), docxIndex: 1 },
      { ...tracked('c'), docxIndex: 2 },
    ])
    const ids = revisionIds(await documentXmlOf(saved))
    expect(ids).toEqual(['9001', '9002', '9003', '9004', '9005'])
  })

  it('a stored id split in two by a later edit is emitted only once', () => {
    const ins = { author: 'A', id: '21' }
    const xml = generateParagraphXml(
      {
        type: 'paragraph',
        runs: [
          { text: 'one ', ins },
          { text: 'cut', ins, del: { author: 'R' } },
          { text: ' two', ins },
        ],
      },
      GEN_CTX,
    )
    expect(revisionIds(xml)).toEqual(['21', '9001', '9002', '9003'])
  })

  it('a minted fallback id never repeats a stored id in the same range', () => {
    const xml = generateParagraphXml(
      {
        type: 'paragraph',
        runs: [
          { text: 'one', ins: { author: 'A', id: '9001' } },
          { text: 'two', ins: { author: 'B' } },
          { text: 'three', ins: { author: 'A', id: '9001' } },
        ],
      },
      GEN_CTX,
    )
    expect(new Set(revisionIds(xml)).size).toBe(revisionIds(xml).length)
  })
})
