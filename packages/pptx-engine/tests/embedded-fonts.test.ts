import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import { PackageArchive } from '../src/zip'
import { eotToSfnt, listEmbeddedFonts, stripEmbeddedFonts } from '../src/embedded-fonts'
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

describe('stripEmbeddedFonts', () => {
  async function archiveWithFonts(): Promise<PackageArchive> {
    const zip = new JSZip()
    zip.file(
      'ppt/presentation.xml',
      '<?xml version="1.0"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" embedTrueTypeFonts="1" saveSubsetFonts="1">' +
        '<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>' +
        '<p:embeddedFontLst>' +
        '<p:embeddedFont><p:font typeface="MiSans"/><p:regular r:id="rId7"/></p:embeddedFont>' +
        '</p:embeddedFontLst>' +
        '<p:sldSz cx="12192000" cy="6858000"/></p:presentation>',
    )
    zip.file(
      'ppt/_rels/presentation.xml.rels',
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>' +
        '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font1.fntdata"/>' +
        '</Relationships>',
    )
    zip.file(
      '[Content_Types].xml',
      '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="fntdata" ContentType="application/x-fontdata"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '</Types>',
    )
    zip.file('ppt/fonts/font1.fntdata', eotWrap(fakeSfnt(), 0x5))
    return PackageArchive.open(await zip.generateAsync({ type: 'uint8array' }))
  }

  it('removes embeddedFontLst, font rels, fntdata parts and the content-type declaration', async () => {
    const archive = await archiveWithFonts()
    expect(stripEmbeddedFonts(archive)).toBe(true)

    const pres = archive.readText('ppt/presentation.xml')!
    expect(pres).not.toContain('embeddedFont')
    expect(pres).not.toContain('embedTrueTypeFonts')
    expect(pres).not.toContain('saveSubsetFonts')
    expect(pres).toContain('<p:sldIdLst>')

    const rels = archive.readText('ppt/_rels/presentation.xml.rels')!
    expect(rels).not.toContain('/font"')
    expect(rels).toContain('slides/slide1.xml') // non-font rels survive

    const ct = archive.readText('[Content_Types].xml')!
    expect(ct).not.toContain('fntdata')
    expect(ct).toContain('Extension="xml"')

    expect(archive.has('ppt/fonts/font1.fntdata')).toBe(false)
  })

  it('is a no-op when no embedded fonts are declared', async () => {
    const archive = await archiveWithFonts()
    stripEmbeddedFonts(archive)
    expect(stripEmbeddedFonts(archive)).toBe(false)
  })
})

describe('save strips stale embedded-font subsets after edits', () => {
  /** Fixture deck with an embeddedFontLst + fntdata part injected. */
  async function fixtureWithEmbeddedFonts(): Promise<Uint8Array> {
    const zip = await JSZip.loadAsync(fx('01_standard_business.pptx'))
    const pres = await zip.file('ppt/presentation.xml')!.async('string')
    zip.file(
      'ppt/presentation.xml',
      pres
        .replace('<p:presentation ', '<p:presentation embedTrueTypeFonts="1" saveSubsetFonts="1" ')
        .replace(
          '</p:presentation>',
          '<p:embeddedFontLst><p:embeddedFont><p:font typeface="MiSans"/><p:regular r:id="rId900"/></p:embeddedFont></p:embeddedFontLst></p:presentation>',
        ),
    )
    const rels = await zip.file('ppt/_rels/presentation.xml.rels')!.async('string')
    zip.file(
      'ppt/_rels/presentation.xml.rels',
      rels.replace(
        '</Relationships>',
        '<Relationship Id="rId900" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font1.fntdata"/></Relationships>',
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
    zip.file('ppt/fonts/font1.fntdata', eotWrap(fakeSfnt(), 0x5))
    return zip.generateAsync({ type: 'uint8array' })
  }

  async function savedPresentationXml(bytes: Uint8Array): Promise<JSZip> {
    return JSZip.loadAsync(bytes)
  }

  it('keeps embedded fonts on a no-edit save', async () => {
    const opened = await openPptx(await fixtureWithEmbeddedFonts())
    const zip = await savedPresentationXml(await savePptx(opened))
    const pres = await zip.file('ppt/presentation.xml')!.async('string')
    expect(pres).toContain('embeddedFontLst')
    expect(zip.file('ppt/fonts/font1.fntdata')).not.toBeNull()
  })

  it('strips embedded fonts once a run is edited', async () => {
    const opened = await openPptx(await fixtureWithEmbeddedFonts())
    const slide = opened.deck.slides.find((s) =>
      s.elements.some(
        (e) => (e.type === 'text' || e.type === 'shape') && (e as any).text?.paragraphs?.length,
      ),
    )!
    const el = slide.elements.find(
      (e) => (e.type === 'text' || e.type === 'shape') && (e as any).text?.paragraphs?.length,
    )!
    expect(setElementFont(slide, el.id, { color: '#123456' })).toBe(true)

    const zip = await savedPresentationXml(await savePptx(opened))
    const pres = await zip.file('ppt/presentation.xml')!.async('string')
    expect(pres).not.toContain('embeddedFont')
    expect(pres).not.toContain('embedTrueTypeFonts')
    const rels = await zip.file('ppt/_rels/presentation.xml.rels')!.async('string')
    expect(rels).not.toContain('/font"')
    expect(zip.file('ppt/fonts/font1.fntdata')).toBeNull()
  })

  it('strips embedded fonts when a slide is duplicated (new text may not be in the subsets)', async () => {
    const opened = await openPptx(await fixtureWithEmbeddedFonts())
    expect(duplicateSlide(opened, 0)).not.toBeNull()

    const zip = await savedPresentationXml(await savePptx(opened))
    const pres = await zip.file('ppt/presentation.xml')!.async('string')
    expect(pres).not.toContain('embeddedFont')
    expect(zip.file('ppt/fonts/font1.fntdata')).toBeNull()
  })
})
