import { describe, expect, it } from 'vitest'
import {
  addElement,
  addPicture,
  addSlideComment,
  addMedia,
  createBlankPptx,
  deleteSlide,
  duplicateSlide,
  openPptx,
  setSlideNotes,
  type OpenedPptx,
} from '../src/index'
import { cNvPrIdsInXml, elementSpid } from '../src/animation'
import { hasContentTypeOverride, maxRelationshipIdNumber } from '../src/xml-utils'
import { relsPathFor } from '../src/zip'

const OFF = { x: 914400, y: 914400, cx: 3657600, cy: 2057400 }
const MP4 = new Uint8Array([0, 0, 0, 24, 0x66, 0x74, 0x79, 0x70])
const DEFAULT_TAGS = (ct: string, ext: string) =>
  [...ct.matchAll(/<Default\b[^>]*\/?>/g)]
    .map((m) => m[0])
    .filter((tag) => new RegExp(`\\bExtension\\s*=\\s*["']${ext}["']`).test(tag))

/** Rewrite every Id="rIdN" in a rels part with single quotes. */
const singleQuoteIds = (xml: string): string => xml.replace(/="(rId\d+)"/g, "='$1'")

/** Rewrite every PartName="/x" in [Content_Types].xml with single quotes. */
const singleQuotePartNames = (xml: string): string => xml.replace(/="(\/[^"]+)"/g, "='$1'")

const relsIds = (xml: string): string[] =>
  [...xml.matchAll(/\bId\s*=\s*(["'])(rId\d+)\1/g)].map((m) => m[2]!)

const partNames = (xml: string): string[] =>
  [...xml.matchAll(/\bPartName\s*=\s*(["'])(.*?)\1/g)].map((m) => m[2]!)

const expectNoDuplicates = (values: string[]) => expect(new Set(values).size).toBe(values.length)

const withSingleQuotedPackage = async (): Promise<OpenedPptx> => {
  const opened = await openPptx(await createBlankPptx())
  const slide = opened.deck.slides[0]!
  const relsPath = relsPathFor(slide.path)
  const rels = opened.archive.readText(relsPath)
  opened.archive.entries.set(relsPath, Buffer.from(singleQuoteIds(rels!), 'utf8'))
  const ct = opened.archive.readText('[Content_Types].xml')!
  opened.archive.entries.set('[Content_Types].xml', Buffer.from(singleQuotePartNames(ct), 'utf8'))
  return opened
}

describe('quote-agnostic relationship and content-type lookup', () => {
  it('counts double- and single-quoted relationship ids alike', () => {
    expect(
      maxRelationshipIdNumber(
        `<Relationships><Relationship Id="rId1"/><Relationship Id='rId7'/></Relationships>`,
      ),
    ).toBe(7)
    expect(
      maxRelationshipIdNumber(`<Relationships><Relationship Id='rId4'/></Relationships>`),
    ).toBe(4)
    expect(
      maxRelationshipIdNumber(`<Relationships><Relationship Id="rId2"/></Relationships>`),
    ).toBe(2)
  })

  it('ignores ids that are not rIdN and handles empty rels', () => {
    expect(
      maxRelationshipIdNumber(`<Relationships><Relationship Id="slide1"/></Relationships>`),
    ).toBe(0)
    expect(
      maxRelationshipIdNumber(
        `<Relationships><Relationship Id='slide1' Id="rId3"/></Relationships>`,
      ),
    ).toBe(3)
    expect(maxRelationshipIdNumber('')).toBe(0)
  })

  it('matches PartName in either quoting style', () => {
    const dq = '<Types><Override PartName="/ppt/slides/slide1.xml"/></Types>'
    const sq = "<Types><Override PartName='/ppt/slides/slide1.xml'/></Types>"
    expect(hasContentTypeOverride(dq, 'ppt/slides/slide1.xml')).toBe(true)
    expect(hasContentTypeOverride(sq, 'ppt/slides/slide1.xml')).toBe(true)
    expect(hasContentTypeOverride(sq, 'ppt/slides/slide2.xml')).toBe(false)
    expect(hasContentTypeOverride(dq, 'ppt/slides/slide1.xml.bak')).toBe(false)
  })

  it('matches a Default extension whatever the attribute order', async () => {
    const opened = await openPptx(await createBlankPptx())
    const ct = opened.archive.readText('[Content_Types].xml')!
    // ContentType before Extension, the order OPC producers are free to write
    const existing = '<Default ContentType="video/mp4" Extension="mp4"/>'
    opened.archive.entries.set(
      '[Content_Types].xml',
      Buffer.from(
        ct.replace('</Types>', () => `${existing}</Types>`),
        'utf8',
      ),
    )

    expect(addMedia(opened, 0, { kind: 'video', bytes: MP4, ext: 'mp4', offset: OFF })).toBeTruthy()

    const mp4Defaults = [
      ...opened.archive.readText('[Content_Types].xml')!.matchAll(/<Default\b[^>]*\/?>/g),
    ]
      .map((m) => m[0])
      .filter((tag) => /\bExtension\s*=\s*["']mp4["']/.test(tag))
    // A second Default for mp4 is what OPC forbids
    expect(mp4Defaults).toEqual([existing])
  })

  it('addPicture matches a Default extension whatever the attribute order', async () => {
    const opened = await openPptx(await createBlankPptx())
    const ctPath = '[Content_Types].xml'
    const existing = '<Default ContentType="image/gif" Extension="gif"/>'
    const ct = opened.archive.readText(ctPath)!
    expect(DEFAULT_TAGS(ct, 'gif')).toEqual([])
    opened.archive.entries.set(
      ctPath,
      Buffer.from(
        ct.replace('</Types>', () => `${existing}</Types>`),
        'utf8',
      ),
    )
    const gif = new Uint8Array([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 1, 0, 1, 0])
    expect(
      addPicture(opened, opened.deck.slides[0]!, { bytes: gif, ext: 'gif', offset: OFF }),
    ).toBeTruthy()
    expect(DEFAULT_TAGS(opened.archive.readText(ctPath)!, 'gif')).toEqual([existing])
  })
})

describe('writers allocate unique ids and overrides on a single-quoted package', () => {
  it('adds notes without duplicating a relationship id or override', async () => {
    const opened = await withSingleQuotedPackage()
    const slide = opened.deck.slides[0]!

    expect(setSlideNotes(opened, 0, 'speaker notes')).toBe(true)

    const ids = relsIds(opened.archive.readText(relsPathFor(slide.path))!)
    expect(ids.length).toBeGreaterThan(0)
    expectNoDuplicates(ids)

    const overrides = partNames(opened.archive.readText('[Content_Types].xml')!)
    expectNoDuplicates(overrides)
    expect(overrides).toContain('/ppt/notesSlides/notesSlide1.xml')
  })

  it('adds media without duplicating a relationship id', async () => {
    const opened = await withSingleQuotedPackage()
    const slide = opened.deck.slides[0]!

    expect(addMedia(opened, 0, { kind: 'video', bytes: MP4, ext: 'mp4', offset: OFF })).toBeTruthy()

    const ids = relsIds(opened.archive.readText(relsPathFor(slide.path))!)
    expectNoDuplicates(ids)
  })

  it('adds a comment without duplicating a relationship id or override', async () => {
    const opened = await withSingleQuotedPackage()
    const slide = opened.deck.slides[0]!

    expect(addSlideComment(opened, 0, { author: 'Ada', text: 'hi' })).toBeTruthy()

    const ids = relsIds(opened.archive.readText(relsPathFor(slide.path))!)
    expectNoDuplicates(ids)
    expectNoDuplicates(partNames(opened.archive.readText('[Content_Types].xml')!))
  })

  it('adds notes to a second slide whose rels only use single quotes', async () => {
    const opened = await withSingleQuotedPackage()
    const first = opened.deck.slides[0]!
    const duplicated = (await import('../src/index')).duplicateSlide(opened, 0)!
    const relsPath = relsPathFor(duplicated.path)
    opened.archive.entries.set(
      relsPath,
      Buffer.from(singleQuoteIds(opened.archive.readText(relsPath)!), 'utf8'),
    )

    expect(setSlideNotes(opened, 1, 'second')).toBe(true)

    const ids = relsIds(opened.archive.readText(relsPathFor(duplicated.path))!)
    expectNoDuplicates(ids)
    expect(ids.length).toBeGreaterThan(1)
    expect(opened.archive.readText(relsPathFor(first.path))).toBeTruthy()
  })
})

describe('quote-agnostic shape id and presentation scans', () => {
  it('collects cNvPr ids written with either quote style', () => {
    const xml =
      `<p:grpSp><p:nvGrpSpPr><p:cNvPr id='7' name="g"/></p:nvGrpSpPr>` +
      `<p:sp><p:nvSpPr><p:cNvPr id="8" name='a'/></p:nvSpPr></p:sp>` +
      `<p:sp><p:nvSpPr><p:cNvPr name='b' id='9'/></p:nvSpPr></p:sp></p:grpSp>`
    expect([...cNvPrIdsInXml(xml)]).toEqual([7, 8, 9])
    expect(elementSpid({ anchor: { originalXml: xml } } as any)).toBe(7)
  })

  it('mints a fresh shape id above single-quoted ones', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    addElement(slide, { kind: 'rect', offset: OFF })
    for (const el of slide.elements)
      el.anchor.originalXml = el.anchor.originalXml.replace(/\bid="(\d+)"/g, "id='$1'")
    const taken = new Set<number>()
    for (const el of slide.elements)
      for (const id of cNvPrIdsInXml(el.anchor.originalXml)) taken.add(id)

    const added = addElement(slide, { kind: 'rect', offset: OFF })
    const minted = elementSpid(added)!
    expect(taken.size).toBeGreaterThan(0)
    expect(taken.has(minted)).toBe(false)
  })

  it('deletes a slide from a single-quoted presentation.xml and rels', async () => {
    const opened = await openPptx(await createBlankPptx())
    duplicateSlide(opened, 0)
    for (const path of ['ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels']) {
      const xml = opened.archive.readText(path)!
      opened.archive.entries.set(path, Buffer.from(singleQuoteIds(xml), 'utf8'))
    }

    expect(deleteSlide(opened, 1)).toBe(true)
    expect(opened.deck.slides).toHaveLength(1)
    const pres = opened.archive.readText('ppt/presentation.xml')!
    expect(pres.match(/<p:sldId\b/g)).toHaveLength(1)
    const rels = opened.archive.readText('ppt/_rels/presentation.xml.rels')!
    expect(rels).not.toContain('slides/slide2.xml')
  })
})
