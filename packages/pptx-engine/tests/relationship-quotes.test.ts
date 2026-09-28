import { describe, expect, it } from 'vitest'
import {
  addSlideComment,
  addMedia,
  createBlankPptx,
  openPptx,
  setSlideNotes,
  type OpenedPptx,
} from '../src/index'
import { hasContentTypeOverride, maxRelationshipIdNumber } from '../src/xml-utils'
import { relsPathFor } from '../src/zip'

const OFF = { x: 914400, y: 914400, cx: 3657600, cy: 2057400 }
const MP4 = new Uint8Array([0, 0, 0, 24, 0x66, 0x74, 0x79, 0x70])

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
