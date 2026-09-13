import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import { PackageArchive } from '../src/zip'
import {
  eotToSfnt,
  listEmbeddedFonts,
  sfntCmapLookup,
  staleEmbeddedTypefaces,
  stripEmbeddedFonts,
} from '../src/embedded-fonts'
import type { TextElement } from '../src/types'
import { openPptx, savePptx, setElementFont, duplicateSlide } from '../src/index'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

/** Minimal valid-enough sfnt payload: TrueType magic + filler. */
function fakeSfnt(seed = 1): Uint8Array {
  const b = new Uint8Array(64).fill(seed)
  b.set([0x00, 0x01, 0x00, 0x00], 0)
  return b
}

/** Wrap a payload in a 16-byte-min EOT header (variable header fields collapsed to none). */
function eotWrap(payload: Uint8Array, flags: number, headerPad = 32): Uint8Array {
  const out = new Uint8Array(16 + headerPad + payload.length)
  const dv = new DataView(out.buffer)
  dv.setUint32(0, out.length, true) // EOTSize
  dv.setUint32(4, payload.length, true) // FontDataSize
  dv.setUint32(8, 0x00020002, true) // Version
  dv.setUint32(12, flags, true)
  out.set(payload, 16 + headerPad)
  return out
}

describe('eotToSfnt', () => {
  it('unwraps an uncompressed EOT to its trailing sfnt payload', () => {
    const sfnt = fakeSfnt()
    expect(eotToSfnt(eotWrap(sfnt, 0))).toEqual(sfnt)
  })

  it('decodes the XOR-obfuscated payload variant', () => {
    const sfnt = fakeSfnt()
    const xored = sfnt.map((b) => b ^ 0x50)
    expect(eotToSfnt(eotWrap(xored, 0x10000000))).toEqual(sfnt)
  })

  it('rejects MicroType-Express-compressed EOTs (flag 0x4)', () => {
    expect(eotToSfnt(eotWrap(fakeSfnt(), 0x00000005))).toBeNull()
  })

  it('rejects payloads without an sfnt magic and truncated headers', () => {
    expect(eotToSfnt(eotWrap(new Uint8Array(32).fill(9), 0))).toBeNull()
    expect(eotToSfnt(new Uint8Array(8))).toBeNull()
  })

  it('passes through a bare sfnt without an EOT wrapper (LibreOffice-style fntdata)', () => {
    const sfnt = fakeSfnt()
    expect(eotToSfnt(sfnt)).toBe(sfnt)
  })
})

describe('listEmbeddedFonts', () => {
  it('reads embeddedFontLst faces via presentation rels, skipping compressed parts', async () => {
    const zip = new JSZip()
    zip.file(
      'ppt/presentation.xml',
      '<?xml version="1.0"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<p:embeddedFontLst>' +
        '<p:embeddedFont><p:font typeface="League Spartan"/><p:regular r:id="rId7"/><p:bold r:id="rId8"/></p:embeddedFont>' +
        '<p:embeddedFont><p:font typeface="Poppins"/><p:regular r:id="rId9"/></p:embeddedFont>' +
        '</p:embeddedFontLst>' +
        '<p:sldSz cx="12192000" cy="6858000"/></p:presentation>',
    )
    zip.file(
      'ppt/_rels/presentation.xml.rels',
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId7" Type="f" Target="fonts/font1.fntdata"/>' +
        '<Relationship Id="rId8" Type="f" Target="fonts/font2.fntdata"/>' +
        '<Relationship Id="rId9" Type="f" Target="fonts/font3.fntdata"/>' +
        '</Relationships>',
    )
    const reg = fakeSfnt(1)
    const bold = fakeSfnt(2)
    zip.file('ppt/fonts/font1.fntdata', eotWrap(reg, 0))
    zip.file('ppt/fonts/font2.fntdata', eotWrap(bold, 0))
    zip.file('ppt/fonts/font3.fntdata', eotWrap(fakeSfnt(3), 0x5)) // MTX-compressed -> skipped
    const archive = await PackageArchive.open(await zip.generateAsync({ type: 'uint8array' }))
    const faces = listEmbeddedFonts(archive)
    expect(faces.map((f) => [f.typeface, f.style])).toEqual([
      ['League Spartan', 'regular'],
      ['League Spartan', 'bold'],
    ])
    expect(faces[0]!.sfnt).toEqual(reg)
    expect(faces[1]!.sfnt).toEqual(bold)
  })

  it('returns [] when no embedded fonts are declared', async () => {
    const zip = new JSZip()
    zip.file(
      'ppt/presentation.xml',
      '<?xml version="1.0"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="12192000" cy="6858000"/></p:presentation>',
    )
    const archive = await PackageArchive.open(await zip.generateAsync({ type: 'uint8array' }))
    expect(listEmbeddedFonts(archive)).toEqual([])
  })
})

/** Minimal sfnt whose only table is a format-4 cmap mapping exactly `chars` to non-zero glyphs. */
function sfntCovering(chars: string): Uint8Array {
  const cps = [...new Set([...chars].map((c) => c.codePointAt(0)!))].sort((a, b) => a - b)
  const segCount = cps.length + 1
  const subLen = 16 + segCount * 8
  const cmapLen = 12 + subLen
  const out = new Uint8Array(12 + 16 + cmapLen)
  const dv = new DataView(out.buffer)
  dv.setUint32(0, 0x00010000)
  dv.setUint16(4, 1)
  dv.setUint32(12, 0x636d6170) // 'cmap'
  dv.setUint32(20, 28)
  dv.setUint32(24, cmapLen)
  const cmap = 28
  dv.setUint16(cmap + 2, 1)
  dv.setUint16(cmap + 4, 3)
  dv.setUint16(cmap + 6, 1)
  dv.setUint32(cmap + 8, 12)
  const sub = cmap + 12
  dv.setUint16(sub, 4)
  dv.setUint16(sub + 2, subLen)
  dv.setUint16(sub + 6, segCount * 2)
  const ends = sub + 14
  const starts = ends + segCount * 2 + 2
  const deltas = starts + segCount * 2
  cps.forEach((cp, i) => {
    dv.setUint16(ends + i * 2, cp)
    dv.setUint16(starts + i * 2, cp)
    dv.setUint16(deltas + i * 2, (i + 1 - cp) & 0xffff)
  })
  dv.setUint16(ends + cps.length * 2, 0xffff)
  dv.setUint16(starts + cps.length * 2, 0xffff)
  dv.setUint16(deltas + cps.length * 2, 1)
  return out
}

const FONT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font'

describe('sfntCmapLookup', () => {
  it('reports exactly the mapped code points', () => {
    const lookup = sfntCmapLookup(sfntCovering('Hello'))!
    expect(lookup('H'.codePointAt(0)!)).toBe(true)
    expect(lookup('o'.codePointAt(0)!)).toBe(true)
    expect(lookup('x'.codePointAt(0)!)).toBe(false)
    expect(lookup(0x4e2d)).toBe(false)
  })

  it('returns null without a readable cmap', () => {
    expect(sfntCmapLookup(fakeSfnt())).toBeNull()
  })
})

describe('stripEmbeddedFonts', () => {
  async function archiveWithFonts(): Promise<PackageArchive> {
    const zip = new JSZip()
    zip.file(
      'ppt/presentation.xml',
      '<?xml version="1.0"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" embedTrueTypeFonts="1" saveSubsetFonts="1">' +
        '<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>' +
        '<p:embeddedFontLst>' +
        '<p:embeddedFont><p:font typeface="MiSans"/><p:regular r:id="rId7"/><p:bold r:id="rId8"/></p:embeddedFont>' +
        '<p:embeddedFont><p:font typeface="Montserrat"/><p:regular r:id="rId9"/></p:embeddedFont>' +
        '</p:embeddedFontLst>' +
        '<p:sldSz cx="12192000" cy="6858000"/></p:presentation>',
    )
    zip.file(
      'ppt/_rels/presentation.xml.rels',
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>' +
        `<Relationship Id="rId7" Type="${FONT_REL}" Target="fonts/font1.fntdata"/>` +
        `<Relationship Id="rId8" Type="${FONT_REL}" Target="fonts/font2.fntdata"/>` +
        `<Relationship Id="rId9" Type="${FONT_REL}" Target="fonts/font3.fntdata"/>` +
        '</Relationships>',
    )
    zip.file(
      '[Content_Types].xml',
      '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="fntdata" ContentType="application/x-fontdata"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/ppt/fonts/font3.fntdata" ContentType="application/x-fontdata"/>' +
        '</Types>',
    )
    zip.file('ppt/fonts/font1.fntdata', eotWrap(sfntCovering('Hello'), 0))
    zip.file('ppt/fonts/font2.fntdata', eotWrap(sfntCovering('Hello'), 0))
    zip.file('ppt/fonts/font3.fntdata', eotWrap(sfntCovering('abc'), 0))
    return PackageArchive.open(await zip.generateAsync({ type: 'uint8array' }))
  }

  it('removes every face, the list, its attributes, rels, parts and content types', async () => {
    const archive = await archiveWithFonts()
    expect(stripEmbeddedFonts(archive)).toBe(true)

    const pres = archive.readText('ppt/presentation.xml')!
    expect(pres).not.toContain('embeddedFont')
    expect(pres).not.toContain('embedTrueTypeFonts')
    expect(pres).not.toContain('saveSubsetFonts')
    expect(pres).toContain('<p:sldIdLst>')

    const rels = archive.readText('ppt/_rels/presentation.xml.rels')!
    expect(rels).not.toContain('/font"')
    expect(rels).toContain('slides/slide1.xml')

    const ct = archive.readText('[Content_Types].xml')!
    expect(ct).not.toContain('fntdata')
    expect(ct).toContain('Extension="xml"')

    expect(archive.has('ppt/fonts/font1.fntdata')).toBe(false)
    expect(archive.has('ppt/fonts/font3.fntdata')).toBe(false)
  })

  it('removes only the named face and keeps the list for the rest', async () => {
    const archive = await archiveWithFonts()
    expect(stripEmbeddedFonts(archive, new Set(['misans']))).toBe(true)

    const pres = archive.readText('ppt/presentation.xml')!
    expect(pres).not.toContain('MiSans')
    expect(pres).toContain('<p:embeddedFontLst><p:embeddedFont><p:font typeface="Montserrat"/>')
    expect(pres).toContain('embedTrueTypeFonts="1"')

    const rels = archive.readText('ppt/_rels/presentation.xml.rels')!
    expect(rels).not.toContain('rId7')
    expect(rels).not.toContain('rId8')
    expect(rels).toContain('rId9')
    expect(archive.has('ppt/fonts/font1.fntdata')).toBe(false)
    expect(archive.has('ppt/fonts/font2.fntdata')).toBe(false)
    expect(archive.has('ppt/fonts/font3.fntdata')).toBe(true)
    expect(archive.readText('[Content_Types].xml')).toContain('Extension="fntdata"')
  })

  it('is a no-op when no embedded fonts are declared', async () => {
    const archive = await archiveWithFonts()
    stripEmbeddedFonts(archive)
    expect(stripEmbeddedFonts(archive)).toBe(false)
  })

  const cps = (s: string) => new Set([...s].map((c) => c.codePointAt(0)!))
  const demand = (key: string, s: string) => new Map([[key, cps(s)]])

  it('flags only faces whose subset misses demanded glyphs', async () => {
    const archive = await archiveWithFonts()
    expect(staleEmbeddedTypefaces(archive, demand('misans|regular', 'Hello'))).toEqual([])
    expect(staleEmbeddedTypefaces(archive, demand('misans|regular', 'Help'))).toEqual(['MiSans'])
    expect(staleEmbeddedTypefaces(archive, demand('montserrat|regular', 'abc'))).toEqual([])
    expect(staleEmbeddedTypefaces(archive, demand('*|regular', 'x'))).toEqual([
      'MiSans',
      'Montserrat',
    ])
    expect(staleEmbeddedTypefaces(archive, demand('calibri|regular', 'xyz'))).toEqual([])
  })

  it('checks each style against its own part', async () => {
    const archive = await archiveWithFonts()
    // Bold part covers "Bold" only; regular covers "Hello" only.
    archive.entries.set('ppt/fonts/font2.fntdata', eotWrap(sfntCovering('Bold'), 0))
    expect(staleEmbeddedTypefaces(archive, demand('misans|bold', 'Bold'))).toEqual([])
    expect(staleEmbeddedTypefaces(archive, demand('misans|regular', 'Hello'))).toEqual([])
    expect(staleEmbeddedTypefaces(archive, demand('misans|bold', 'Hello'))).toEqual(['MiSans'])
    // No italic part: italic runs render from the regular subset.
    expect(staleEmbeddedTypefaces(archive, demand('misans|italic', 'Hello'))).toEqual([])
    expect(staleEmbeddedTypefaces(archive, demand('misans|italic', 'Bold'))).toEqual(['MiSans'])
  })
})

describe('save strips embedded faces the edits outgrew', () => {
  /** Fixture deck with two embedded faces injected: MiSans covers "Hello", Montserrat covers "abc". */
  async function fixtureWithEmbeddedFonts(): Promise<Uint8Array> {
    const zip = await JSZip.loadAsync(fx('01_standard_business.pptx'))
    const pres = await zip.file('ppt/presentation.xml')!.async('string')
    zip.file(
      'ppt/presentation.xml',
      pres
        .replace('<p:presentation ', '<p:presentation embedTrueTypeFonts="1" saveSubsetFonts="1" ')
        .replace(
          '</p:presentation>',
          '<p:embeddedFontLst>' +
            '<p:embeddedFont><p:font typeface="MiSans"/><p:regular r:id="rId900"/></p:embeddedFont>' +
            '<p:embeddedFont><p:font typeface="Montserrat"/><p:regular r:id="rId901"/></p:embeddedFont>' +
            '</p:embeddedFontLst></p:presentation>',
        ),
    )
    const rels = await zip.file('ppt/_rels/presentation.xml.rels')!.async('string')
    zip.file(
      'ppt/_rels/presentation.xml.rels',
      rels.replace(
        '</Relationships>',
        `<Relationship Id="rId900" Type="${FONT_REL}" Target="fonts/font1.fntdata"/>` +
          `<Relationship Id="rId901" Type="${FONT_REL}" Target="fonts/font2.fntdata"/></Relationships>`,
      ),
    )
    const ct = await zip.file('[Content_Types].xml')!.async('string')
    zip.file(
      '[Content_Types].xml',
      ct.replace(
        /<Types[^>]*>/,
        '$&<Default Extension="fntdata" ContentType="application/x-fontdata"/>',
      ),
    )
    zip.file('ppt/fonts/font1.fntdata', eotWrap(sfntCovering('Hello'), 0))
    zip.file('ppt/fonts/font2.fntdata', eotWrap(sfntCovering('abc'), 0))
    return zip.generateAsync({ type: 'uint8array' })
  }

  function firstTextElement(opened: Awaited<ReturnType<typeof openPptx>>) {
    const slide = opened.deck.slides.find((s) =>
      s.elements.some(
        (e) => (e.type === 'text' || e.type === 'shape') && (e as any).text?.paragraphs?.length,
      ),
    )!
    const el = slide.elements.find(
      (e) => (e.type === 'text' || e.type === 'shape') && (e as any).text?.paragraphs?.length,
    )! as TextElement
    return { slide, el }
  }

  async function savedParts(opened: Awaited<ReturnType<typeof openPptx>>) {
    const zip = await JSZip.loadAsync(await savePptx(opened))
    return {
      pres: await zip.file('ppt/presentation.xml')!.async('string'),
      rels: await zip.file('ppt/_rels/presentation.xml.rels')!.async('string'),
      font1: zip.file('ppt/fonts/font1.fntdata'),
      font2: zip.file('ppt/fonts/font2.fntdata'),
    }
  }

  it('keeps embedded fonts on a no-edit save', async () => {
    const opened = await openPptx(await fixtureWithEmbeddedFonts())
    const out = await savedParts(opened)
    expect(out.pres).toContain('embeddedFontLst')
    expect(out.font1).not.toBeNull()
    expect(out.font2).not.toBeNull()
  })

  it('keeps embedded fonts when the edited text uses other faces', async () => {
    const opened = await openPptx(await fixtureWithEmbeddedFonts())
    const { slide, el } = firstTextElement(opened)
    expect(setElementFont(slide, el.id, { color: '#123456' })).toBe(true)
    const out = await savedParts(opened)
    expect(out.pres).toContain('embeddedFontLst')
    expect(out.font1).not.toBeNull()
  })

  it('keeps a face whose bold part covers a bold run the regular part does not', async () => {
    const zip = await JSZip.loadAsync(await fixtureWithEmbeddedFonts())
    const pres = await zip.file('ppt/presentation.xml')!.async('string')
    zip.file(
      'ppt/presentation.xml',
      pres.replace(
        '<p:regular r:id="rId900"/>',
        '<p:regular r:id="rId900"/><p:bold r:id="rId902"/>',
      ),
    )
    const rels = await zip.file('ppt/_rels/presentation.xml.rels')!.async('string')
    zip.file(
      'ppt/_rels/presentation.xml.rels',
      rels.replace(
        '</Relationships>',
        `<Relationship Id="rId902" Type="${FONT_REL}" Target="fonts/font3.fntdata"/></Relationships>`,
      ),
    )
    zip.file('ppt/fonts/font3.fntdata', eotWrap(sfntCovering('Bold'), 0))
    const opened = await openPptx(await zip.generateAsync({ type: 'uint8array' }))
    const { el } = firstTextElement(opened)
    el.text!.paragraphs = [{ runs: [{ text: 'Bold', fontFamily: 'MiSans', bold: true }] }]
    el.dirty = true
    const out = await savedParts(opened)
    expect(out.pres).toContain('typeface="MiSans"')
    expect(out.font1).not.toBeNull()
  })

  it('keeps a face whose subset still covers the edited text', async () => {
    const opened = await openPptx(await fixtureWithEmbeddedFonts())
    const { el } = firstTextElement(opened)
    el.text!.paragraphs = [{ runs: [{ text: 'Hello', fontFamily: 'MiSans' }] }]
    el.dirty = true
    const out = await savedParts(opened)
    expect(out.pres).toContain('typeface="MiSans"')
    expect(out.font1).not.toBeNull()
  })

  it('strips only the face the edited text outgrew', async () => {
    const opened = await openPptx(await fixtureWithEmbeddedFonts())
    const { el } = firstTextElement(opened)
    el.text!.paragraphs = [{ runs: [{ text: 'Hello World', fontFamily: 'MiSans' }] }]
    el.dirty = true
    const out = await savedParts(opened)
    expect(out.pres).not.toContain('MiSans')
    expect(out.pres).toContain('typeface="Montserrat"')
    expect(out.pres).toContain('embedTrueTypeFonts')
    expect(out.rels).not.toContain('rId900')
    expect(out.rels).toContain('rId901')
    expect(out.font1).toBeNull()
    expect(out.font2).not.toBeNull()
  })

  it('drops the whole list once every face is stripped', async () => {
    const opened = await openPptx(await fixtureWithEmbeddedFonts())
    const { el } = firstTextElement(opened)
    el.text!.paragraphs = [
      { runs: [{ text: 'Help', fontFamily: 'MiSans' }] },
      { runs: [{ text: 'xyz', fontFamily: 'Montserrat' }] },
    ]
    el.dirty = true
    const out = await savedParts(opened)
    expect(out.pres).not.toContain('embeddedFont')
    expect(out.pres).not.toContain('embedTrueTypeFonts')
    expect(out.rels).not.toContain('/font"')
    expect(out.font1).toBeNull()
    expect(out.font2).toBeNull()
  })

  it('keeps embedded fonts when a slide is duplicated (no new glyphs)', async () => {
    const opened = await openPptx(await fixtureWithEmbeddedFonts())
    expect(duplicateSlide(opened, 0)).not.toBeNull()
    const out = await savedParts(opened)
    expect(out.pres).toContain('embeddedFontLst')
    expect(out.font1).not.toBeNull()
  })
})
