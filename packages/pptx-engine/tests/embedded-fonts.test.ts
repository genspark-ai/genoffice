import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { PackageArchive } from '../src/zip'
import { eotToSfnt, listEmbeddedFonts } from '../src/embedded-fonts'

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
