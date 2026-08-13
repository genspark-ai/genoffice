import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  INK_NAME_PREFIX,
  injectInkRunsIntoParagraph,
  parseDocx,
  saveDocx,
  stripInkRuns,
  type NewInkImage,
  type SaveBlock,
} from '../src/index'
import { buildDocx, TINY_PNG_BASE64 } from './helpers/build-docx'

const TWO_PARAS = '<w:p><w:r><w:t>第一段</w:t></w:r></w:p><w:p><w:r><w:t>第二段</w:t></w:r></w:p>'

function asOriginal(doc: Awaited<ReturnType<typeof parseDocx>>): SaveBlock[] {
  return doc.blocks
    .filter((b) => !b.hidden)
    .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))
}

function makeInk(blockIndex: number, extra: Partial<NewInkImage> = {}): NewInkImage {
  return {
    blockIndex,
    base64: TINY_PNG_BASE64,
    widthPx: 200,
    heightPx: 80,
    offsetXPx: 40,
    offsetYPx: -10,
    payload: '{"strokes":[{"tool":"pen","color":"C00000"}]}',
    ...extra,
  }
}

async function docXmlOf(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  return zip.file('word/document.xml')!.async('string')
}

describe('ink save (wp:anchor floating picture)', () => {
  it('writes an anchored in-front-of-text picture into the anchor paragraph', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: TWO_PARAS }))
    const saved = await saveDocx(doc, asOriginal(doc), { inks: [makeInk(1)] })
    const xml = await docXmlOf(saved)

    // second paragraph carries the anchored run; the first is untouched
    expect(xml).toContain('<w:p><w:r><w:t>第一段</w:t></w:r></w:p>')
    const second = /<w:p><w:r><w:t>第二段[\s\S]*?<\/wp:anchor><\/w:drawing><\/w:r><\/w:p>/.exec(xml)
    expect(second).not.toBeNull()
    expect(second![0]).toContain('behindDoc="0"')
    expect(second![0]).toContain('<wp:wrapNone/>')
    expect(second![0]).toContain(
      '<wp:positionH relativeFrom="column"><wp:posOffset>381000</wp:posOffset>',
    )
    expect(second![0]).toContain(
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>-95250</wp:posOffset>',
    )
    expect(second![0]).toContain(`name="${INK_NAME_PREFIX} `)
    expect(second![0]).toContain('descr="{&quot;strokes&quot;')

    // media part + relationship + png content type all present
    const zip = await JSZip.loadAsync(saved)
    expect(zip.file('word/media/aidocsink1.png')).not.toBeNull()
    expect(await zip.file('word/_rels/document.xml.rels')!.async('string')).toContain(
      'Target="media/aidocsink1.png"',
    )
    expect(await zip.file('[Content_Types].xml')!.async('string')).toContain('Extension="png"')
  })

  it('anchors into a self-closing empty paragraph', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: '<w:p/>' }))
    const saved = await saveDocx(doc, asOriginal(doc), { inks: [makeInk(0)] })
    const xml = await docXmlOf(saved)
    expect(xml).toContain(`name="${INK_NAME_PREFIX} `)
    expect(xml).not.toContain('<w:p/><w:r>')
  })

  it('skips a non-paragraph anchor without leaving orphan media/relationships', async () => {
    const table =
      '<w:tbl><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    const doc = await parseDocx(await buildDocx({ bodyXml: table }))
    const saved = await saveDocx(doc, asOriginal(doc), { inks: [makeInk(0)] })
    const zip = await JSZip.loadAsync(saved)
    expect(zip.file('word/media/aidocsink1.png')).toBeNull()
    expect(await zip.file('word/_rels/document.xml.rels')!.async('string')).not.toContain(
      'aidocsink',
    )
    expect(await docXmlOf(saved)).not.toContain(INK_NAME_PREFIX)
  })

  it('inks undefined keeps the file byte-identical (no-op save)', async () => {
    const bytes = await buildDocx({ bodyXml: TWO_PARAS })
    const doc = await parseDocx(bytes)
    expect(await saveDocx(doc, asOriginal(doc))).toBe(bytes)
  })
})

describe('ink reopen (ParsedDoc.inks)', () => {
  async function savedWithInk(): Promise<Uint8Array> {
    const doc = await parseDocx(await buildDocx({ bodyXml: TWO_PARAS }))
    return saveDocx(doc, asOriginal(doc), { inks: [makeInk(1)] })
  }

  it('round-trips placement, payload and image bytes', async () => {
    const reopened = await parseDocx(await savedWithInk())
    expect(reopened.inks).toHaveLength(1)
    const ink = reopened.inks[0]
    expect(ink.anchorIndex).toBe(1)
    expect(ink.offsetXPx).toBeCloseTo(40, 1)
    expect(ink.offsetYPx).toBeCloseTo(-10, 1)
    expect(ink.widthPx).toBeCloseTo(200, 1)
    expect(ink.heightPx).toBeCloseTo(80, 1)
    expect(ink.payload).toBe('{"strokes":[{"tool":"pen","color":"C00000"}]}')
    expect(ink.dataUrl).toContain('data:image/png;base64,')
  })

  it('the annotated paragraph stays an editable text paragraph', async () => {
    const reopened = await parseDocx(await savedWithInk())
    const para = reopened.blocks[1]
    expect(para.type).toBe('paragraph')
    expect(para.runs?.map((r) => r.text).join('')).toBe('第二段')
  })

  it('saving with inks: [] removes the annotation, its media part and its relationship', async () => {
    const reopened = await parseDocx(await savedWithInk())
    const cleared = await saveDocx(reopened, asOriginal(reopened), { inks: [] })
    const xml = await docXmlOf(cleared)
    expect(xml).not.toContain(INK_NAME_PREFIX)
    expect(xml).toContain('第二段')
    expect((await parseDocx(cleared)).inks).toHaveLength(0)
    const zip = await JSZip.loadAsync(cleared)
    expect(zip.file('word/media/aidocsink1.png')).toBeNull()
    expect(await zip.file('word/_rels/document.xml.rels')!.async('string')).not.toContain(
      'aidocsink',
    )
  })

  it('re-saving does not accumulate media parts or relationships', async () => {
    const reopened = await parseDocx(await savedWithInk())
    const ink = reopened.inks[0]
    const resaved = await saveDocx(reopened, asOriginal(reopened), {
      inks: [
        {
          blockIndex: 1,
          base64: TINY_PNG_BASE64,
          widthPx: ink.widthPx,
          heightPx: ink.heightPx,
          offsetXPx: ink.offsetXPx,
          offsetYPx: ink.offsetYPx,
          payload: ink.payload ?? undefined,
        },
      ],
    })
    const zip = await JSZip.loadAsync(resaved)
    const inkMedia = Object.keys(zip.files).filter((n) => n.includes('aidocsink'))
    expect(inkMedia).toEqual(['word/media/aidocsink1.png'])
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
    expect(rels.match(/aidocsink/g)).toHaveLength(1)
    expect((await parseDocx(resaved)).inks).toHaveLength(1)
  })

  it('re-saving moves the annotation to a new anchor without duplicating runs', async () => {
    const reopened = await parseDocx(await savedWithInk())
    const ink = reopened.inks[0]
    const moved = await saveDocx(reopened, asOriginal(reopened), {
      inks: [
        {
          blockIndex: 0,
          base64: TINY_PNG_BASE64,
          widthPx: ink.widthPx,
          heightPx: ink.heightPx,
          offsetXPx: ink.offsetXPx,
          offsetYPx: ink.offsetYPx,
          payload: ink.payload ?? undefined,
        },
      ],
    })
    const again = await parseDocx(moved)
    expect(again.inks).toHaveLength(1)
    expect(again.inks[0].anchorIndex).toBe(0)
  })
})

describe('ink XML helpers', () => {
  it('stripInkRuns leaves foreign anchored drawings alone', () => {
    const foreign =
      '<w:p><w:r><w:drawing><wp:anchor behindDoc="1"><wp:docPr id="5" name="图片 5"/>' +
      '</wp:anchor></w:drawing></w:r><w:r><w:t>x</w:t></w:r></w:p>'
    expect(stripInkRuns(foreign)).toBe(foreign)
  })

  it('injectInkRunsIntoParagraph rejects non-paragraph roots', () => {
    expect(injectInkRunsIntoParagraph('<w:tbl></w:tbl>', '<w:r/>')).toBeNull()
    expect(injectInkRunsIntoParagraph('<w:p><w:r><w:t>a</w:t></w:r></w:p>', '<w:r/>')).toBe(
      '<w:p><w:r><w:t>a</w:t></w:r><w:r/></w:p>',
    )
  })
})
