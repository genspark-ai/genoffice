/**
 * A heading made in a document whose styles.xml defines no heading style for
 * its level must survive the save: the generator writes a heading as its
 * level's style reference, so without one the paragraph came back as body
 * text on reopen (the editor's autosave round trip showed it "reverting").
 * Word's built-in definition is created for each missing level.
 */
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import {
  buildBlankDocx,
  builtInHeadingStyleUpsert,
  missingHeadingLevels,
  parseDocx,
  saveDocx,
  type SaveBlock,
} from '../src/index'

const STYLE_RE = (id: string): RegExp =>
  new RegExp(`<w:style\\b[^>]*w:styleId="${id}"[^>]*>[\\s\\S]*?</w:style>`)

async function withoutHeadingStyles(bytes: Uint8Array): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes)
  const xml = await zip.file('word/styles.xml')!.async('string')
  zip.file(
    'word/styles.xml',
    xml.replace(/<w:style\b[^>]*w:styleId="Heading\d"[^>]*>[\s\S]*?<\/w:style>/g, ''),
  )
  return zip.generateAsync({ type: 'uint8array' })
}

async function stylesXmlOf(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  return zip.file('word/styles.xml')!.async('string')
}

async function documentXmlOf(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  return zip.file('word/document.xml')!.async('string')
}

const heading = (level: number, text: string, extra: Record<string, unknown> = {}): SaveBlock => ({
  kind: 'generated',
  block: { type: 'heading', level, runs: [{ text }], ...extra },
})

describe('missingHeadingLevels', () => {
  it('lists the generated heading levels the document has no style for, once each, in order', () => {
    const blocks: SaveBlock[] = [
      heading(3, 'c'),
      heading(1, 'a'),
      heading(3, 'c2'),
      { kind: 'generated', block: { type: 'paragraph', runs: [{ text: 'body' }] } },
      { kind: 'original', docxIndex: 0 },
    ]
    expect(missingHeadingLevels(blocks, new Map())).toEqual([1, 3])
    expect(missingHeadingLevels(blocks, new Map([[1, 'Heading1']]))).toEqual([3])
    expect(
      missingHeadingLevels(
        blocks,
        new Map([
          [1, 'Heading1'],
          [3, 'Heading3'],
        ]),
      ),
    ).toEqual([])
  })

  it('leaves outline-only headings and headings with their own style alone', () => {
    const blocks: SaveBlock[] = [
      heading(1, 'outline', { outlineOnly: true }),
      heading(2, 'custom', { styleId: 'MyTitle' }),
    ]
    expect(missingHeadingLevels(blocks, new Map())).toEqual([])
  })

  it('clamps the level like the generator does', () => {
    expect(missingHeadingLevels([heading(0, 'x'), heading(12, 'y')], new Map())).toEqual([1, 9])
  })
})

describe('builtInHeadingStyleUpsert', () => {
  it('is the blank template heading, based on the document default style', () => {
    expect(builtInHeadingStyleUpsert(2, 'a')).toEqual({
      styleId: 'Heading2',
      type: 'paragraph',
      name: 'heading 2',
      basedOn: 'a',
      next: 'a',
      quickFormat: true,
      builtIn: true,
      pPr: { keepNext: true, spaceBeforeTwips: 240, spaceAfterTwips: 120, outlineLevel: 2 },
      rPr: { bold: true, sizeHalfPoints: 28 },
    })
  })

  it('needs no default style, and deep levels get the smallest size and tighter spacing', () => {
    const up = builtInHeadingStyleUpsert(7, undefined)
    expect(up.basedOn).toBeUndefined()
    expect(up.next).toBeUndefined()
    expect(up.pPr?.spaceBeforeTwips).toBe(160)
    expect(up.pPr?.outlineLevel).toBe(7)
    expect(up.rPr?.sizeHalfPoints).toBe(22)
  })
})

describe('saveDocx creates the heading styles a document lacks', () => {
  it('a heading in a document without heading styles reopens as that heading', async () => {
    const doc = await parseDocx(await withoutHeadingStyles(await buildBlankDocx()))
    expect(doc.headingStyleIds.size).toBe(0)

    const saved = await saveDocx(doc, [
      heading(1, 'Chapter'),
      heading(3, 'Detail'),
      { kind: 'generated', block: { type: 'paragraph', runs: [{ text: 'body' }] } },
    ])
    const reparsed = await parseDocx(saved)
    const visible = reparsed.blocks.filter((b) => !b.hidden)
    expect(visible.map((b) => [b.type, b.level ?? null, b.styleId ?? null])).toEqual([
      ['heading', 1, 'Heading1'],
      ['heading', 3, 'Heading3'],
      ['paragraph', null, null],
    ])
    expect(reparsed.headingStyleIds.get(1)).toBe('Heading1')
    expect(reparsed.headingStyleIds.get(3)).toBe('Heading3')

    const styles = await stylesXmlOf(saved)
    const h1 = STYLE_RE('Heading1').exec(styles)?.[0]
    expect(h1).toBeDefined()
    expect(h1).toContain('<w:name w:val="heading 1"/>')
    expect(h1).toContain('<w:basedOn w:val="Normal"/>')
    expect(h1).toContain('<w:next w:val="Normal"/>')
    expect(h1).toContain('<w:qFormat/>')
    expect(h1).toContain('<w:keepNext/>')
    expect(h1).toContain('<w:outlineLvl w:val="0"/>')
    expect(h1).toContain('<w:b/>')
    expect(h1).toContain('<w:sz w:val="32"/>')
    // Word's built-in headings are not custom styles
    expect(h1).not.toContain('w:customStyle')
    expect(STYLE_RE('Heading3').exec(styles)?.[0]).toContain('<w:outlineLvl w:val="2"/>')
    // only the levels the body uses are created
    expect(styles).not.toMatch(STYLE_RE('Heading2'))

    const document = await documentXmlOf(saved)
    expect(document).toContain('<w:pStyle w:val="Heading1"/>')
    expect(document).toContain('<w:pStyle w:val="Heading3"/>')
  })

  it('a document that already defines the heading style keeps its definition', async () => {
    const doc = await parseDocx(await buildBlankDocx())
    const before = await stylesXmlOf(await buildBlankDocx())
    const saved = await saveDocx(doc, [heading(1, 'Chapter')])
    const styles = await stylesXmlOf(saved)
    expect(styles.match(/w:styleId="Heading1"/g)?.length).toBe(1)
    expect(STYLE_RE('Heading1').exec(styles)?.[0]).toBe(STYLE_RE('Heading1').exec(before)?.[0])
    expect(await documentXmlOf(saved)).toContain('<w:pStyle w:val="Heading1"/>')
  })

  it('a heading style under another id (localized) is used instead of inventing HeadingN', async () => {
    const zip = await JSZip.loadAsync(await withoutHeadingStyles(await buildBlankDocx()))
    const xml = await zip.file('word/styles.xml')!.async('string')
    zip.file(
      'word/styles.xml',
      xml.replace(
        '</w:styles>',
        '<w:style w:type="paragraph" w:styleId="berschrift1"><w:name w:val="Überschrift 1"/>' +
          '<w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style></w:styles>',
      ),
    )
    const doc = await parseDocx(await zip.generateAsync({ type: 'uint8array' }))
    expect(doc.headingStyleIds.get(1)).toBe('berschrift1')

    const saved = await saveDocx(doc, [heading(1, 'Kapitel')])
    expect(await documentXmlOf(saved)).toContain('<w:pStyle w:val="berschrift1"/>')
    expect(await stylesXmlOf(saved)).not.toMatch(STYLE_RE('Heading1'))
    const reparsed = await parseDocx(saved)
    expect(reparsed.blocks.filter((b) => !b.hidden)[0]).toMatchObject({
      type: 'heading',
      level: 1,
      styleId: 'berschrift1',
    })
  })

  it('an outline-only heading keeps writing its level and creates no style', async () => {
    const doc = await parseDocx(await withoutHeadingStyles(await buildBlankDocx()))
    const saved = await saveDocx(doc, [heading(2, 'Outline', { outlineOnly: true })])
    expect(await documentXmlOf(saved)).toContain('<w:outlineLvl w:val="1"/>')
    expect(await stylesXmlOf(saved)).not.toMatch(STYLE_RE('Heading2'))
    const reparsed = await parseDocx(saved)
    expect(reparsed.blocks.filter((b) => !b.hidden)[0]).toMatchObject({
      type: 'heading',
      level: 2,
      outlineOnly: true,
    })
  })
})
