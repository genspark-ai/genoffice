import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '../src/index'
import { nextFreeRelId } from '../src/patch'
import { TINY_PNG_BASE64, buildDocx } from './helpers/build-docx'

/**
 * Allocating the next relationship id in a .rels part used to test one
 * candidate id at a time against the whole part:
 *
 *     let n = 1
 *     while (relTagWithId(`rId${n}`).test(relsXml)) n++
 *
 * Every candidate costs a full scan of the part, so a part holding N
 * relationships costs N scans to find the one free id - quadratic. A part with
 * ~50k relationships (a few MB, inside a ZIP of only a few MB) then spends tens
 * of minutes of CPU on saving, for one id.
 *
 * The allocator now collects the ids the part already hands out in a single
 * pass. The ids it treats as taken are exactly the ones relTagWithId matched -
 * an id the watermark reclaim failed to free is still never reissued - so the
 * id handed out is unchanged; only the cost of finding it is.
 *
 * These tests pin the chosen id, so a future refactor cannot quietly change
 * which id a part gets (a different id is a different relationship as far as
 * every other part in the package is concerned).
 */

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const RELS_OPEN = `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
const RELS_CLOSE = '</Relationships>'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const IMAGE_REL_TYPE = `${REL}/image`

/** A .rels part holding exactly the given <Relationship> tags. */
const relsOf = (tags: string): string => RELS_OPEN + tags + RELS_CLOSE

const rel = (id: string, target = `media/${id}.png`): string =>
  `<Relationship Id="${id}" Type="${IMAGE_REL_TYPE}" Target="${target}"/>`

/** Every id in a rels part, in whichever quote style it is spelled. */
const relIds = (relsXml: string): string[] =>
  [...relsXml.matchAll(/\bId\s*=\s*(["'])(.*?)\1/g)].map((m) => m[2])

const HEADER_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
const XML_DECL_HDR_NS =
  `${XML_DECL}<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
  '<w:p><w:r><w:t>body text</w:t></w:r></w:p></w:hdr>'

/** A docx whose header part's own rels already hold `relsTags`. */
const docWithHeaderRels = (relsTags: string) =>
  buildDocx({
    bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
    extraRels: `<Relationship Id="rId88" Type="${REL}/header" Target="header1.xml"/>`,
    extraParts: [
      { path: 'word/header1.xml', xml: XML_DECL_HDR_NS, contentType: HEADER_TYPE },
      {
        path: 'word/_rels/header1.xml.rels',
        xml: relsOf(relsTags),
        contentType: 'application/vnd.openxmlformats-package.relationships+xml',
      },
    ],
    sectPrExtra: '<w:headerReference w:type="default" r:id="rId88"/>',
  })

const WATERMARK = {
  image: { base64: TINY_PNG_BASE64, mime: 'image/png' as const, widthPx: 96, heightPx: 48 },
}

const originalOrder = (doc: Awaited<ReturnType<typeof parseDocx>>) =>
  doc.blocks
    .filter((b) => !b.hidden && b.docxIndex !== null)
    .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! }))

const partText = async (bytes: Uint8Array, path: string): Promise<string> => {
  const zip = await JSZip.loadAsync(bytes)
  const file = zip.file(path)
  if (!file) throw new Error(`missing part ${path}`)
  return file.async('string')
}

/** A rels part of `count` contiguous relationships, rId1..rId{count}. */
const manyRels = (count: number): string => {
  const parts: string[] = [RELS_OPEN]
  for (let i = 1; i <= count; i++) parts.push(rel(`rId${i}`, `media/image${i}.png`))
  parts.push(RELS_CLOSE)
  return parts.join('')
}

describe('next relationship id in a .rels part', () => {
  it('starts at rId1 in a part that hands out nothing', () => {
    expect(nextFreeRelId(relsOf(''))).toBe('rId1')
  })

  it('takes the id after the last one in a normally numbered part', () => {
    expect(nextFreeRelId(relsOf(rel('rId1') + rel('rId2') + rel('rId3')))).toBe('rId4')
  })

  it('fills a gap rather than skipping past it', () => {
    expect(nextFreeRelId(relsOf(rel('rId1') + rel('rId2') + rel('rId4') + rel('rId5')))).toBe(
      'rId3',
    )
  })

  it('never hands out rId0, even when the part already owns it', () => {
    expect(nextFreeRelId(relsOf(rel('rId0') + rel('rId1')))).toBe('rId2')
  })

  it('leaves a part numbered from rId7 alone and starts from rId1', () => {
    // the allocator looks for the lowest free id, it does not continue the
    // part's own numbering
    expect(nextFreeRelId(relsOf(rel('rId7') + rel('rId8') + rel('rId9')))).toBe('rId1')
  })

  it('reads ids spelled with either quote style or padded around =', () => {
    const padded = relsOf(
      `<Relationship Id = 'rId1' Type = '${IMAGE_REL_TYPE}' Target = 'media/a.png'/>` +
        `<Relationship Id='rId2' Type='${IMAGE_REL_TYPE}' Target='media/b.png'/>`,
    )
    expect(nextFreeRelId(padded)).toBe('rId3')
  })

  it('keeps a zero-padded rId02 distinct from rId2', () => {
    // relTagWithId compares the id literally, so rId02 never blocks rId2
    expect(nextFreeRelId(relsOf(rel('rId1') + rel('rId02')))).toBe('rId2')
  })

  it('ignores an id carried by an element that is not a Relationship', () => {
    const withOverride =
      RELS_OPEN + `<Override Id="rId1" PartName="/x"/>` + rel('rId2') + RELS_CLOSE
    expect(nextFreeRelId(withOverride)).toBe('rId1')
  })

  it('ignores an id on a Relationship tag that does not self-close', () => {
    const unclosed = relsOf(
      `<Relationship Id="rId1" Type="${IMAGE_REL_TYPE}" Target="media/a.png"></Relationship>` +
        rel('rId2'),
    )
    expect(nextFreeRelId(unclosed)).toBe('rId1')
  })

  it('allocates the right id in a part with 50,000 relationships', () => {
    // Only a single pass can answer this in any sane time: the old allocator
    // needed one full scan of this ~7 MB part per candidate id.
    expect(nextFreeRelId(manyRels(50_000))).toBe('rId50001')
  }, 60_000)

  it('allocates the right id in a part with 200,000 relationships', () => {
    // The scale gate, and deliberately not a wall-clock budget: reaching the
    // right answer here is only possible in one pass over the part. The old
    // allocator needed one full scan of this ~28 MB part per candidate id -
    // some 20 hours of CPU - so this assertion cannot pass on it however long
    // a generous timeout allows, while the fixed allocator answers in well
    // under a second.
    expect(nextFreeRelId(manyRels(200_000))).toBe('rId200001')
  }, 60_000)
})

describe('a picture watermark landing in a header rels part', () => {
  it('takes the lowest free id in the part', async () => {
    const doc = await parseDocx(await docWithHeaderRels(rel('rId1') + rel('rId2')))
    const saved = await saveDocx(doc, originalOrder(doc), { watermark: WATERMARK })
    const relsXml = await partText(saved, 'word/_rels/header1.xml.rels')
    const ids = relIds(relsXml)
    expect(new Set(ids).size).toBe(ids.length)
    expect(ids).toContain('rId3')
    expect(relsXml).toContain(`<Relationship Id="rId3" Type="${IMAGE_REL_TYPE}"`)
  })

  it('fills a gap in the part rather than appending past it', async () => {
    const doc = await parseDocx(await docWithHeaderRels(rel('rId1') + rel('rId3')))
    const saved = await saveDocx(doc, originalOrder(doc), { watermark: WATERMARK })
    const relsXml = await partText(saved, 'word/_rels/header1.xml.rels')
    const ids = relIds(relsXml)
    expect(new Set(ids).size).toBe(ids.length)
    expect(ids).toContain('rId2')
    expect(relsXml).toContain(`<Relationship Id="rId2" Type="${IMAGE_REL_TYPE}"`)
  })

  it('keeps ids unique in a header rels part that already holds many', async () => {
    const many = Array.from({ length: 400 }, (_v, i) => rel(`rId${i + 1}`)).join('')
    const doc = await parseDocx(await docWithHeaderRels(many))
    const saved = await saveDocx(doc, originalOrder(doc), { watermark: WATERMARK })
    const relsXml = await partText(saved, 'word/_rels/header1.xml.rels')
    const ids = relIds(relsXml)
    expect(new Set(ids).size).toBe(ids.length)
    expect(ids).toContain('rId401')
    expect(ids.filter((id) => id === 'rId401')).toHaveLength(1)
  }, 30_000)
})
