import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { PackageArchive } from '../src/index'
import { canonicalizeNsPrefixes } from '../src/xml-utils'

const P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'

// What Python's ElementTree writes back after editing presentation.xml.
const presentationNs0 =
  `<?xml version='1.0' encoding='utf-8'?>\n` +
  `<ns0:presentation xmlns:ns0="${P}" xmlns:ns1="${R}" saveSubsetFonts="1">` +
  `<ns0:sldIdLst><ns0:sldId id="256" ns1:id="rId7" /><ns0:sldId id="257" ns1:id="rId8" /></ns0:sldIdLst>` +
  `<ns0:sldSz cx="12192000" cy="6858000" type="screen4x3" /></ns0:presentation>`
const relsNs0 =
  `<?xml version='1.0' encoding='utf-8'?>\n` +
  `<ns0:Relationships xmlns:ns0="${PKG}">` +
  `<ns0:Relationship Id="rId7" Type="${R}/slide" Target="slides/slide1.xml" />` +
  `<ns0:Relationship Id="rId8" Type="${R}/slide" Target="slides/slide2.xml" />` +
  `</ns0:Relationships>`

describe('canonicalizeNsPrefixes', () => {
  it('renames ElementTree prefixes to the PowerPoint ones', () => {
    const out = canonicalizeNsPrefixes(presentationNs0)
    expect(out).toContain(`<p:presentation xmlns:p="${P}" xmlns:r="${R}"`)
    expect(out).toContain('<p:sldId id="256" r:id="rId7" />')
    expect(out).toContain('</p:presentation>')
    expect(out).not.toContain('ns0')
  })

  it('turns a prefixed package-relationships namespace back into the default one', () => {
    const out = canonicalizeNsPrefixes(relsNs0)
    expect(out).toContain(`<Relationships xmlns="${PKG}">`)
    expect(out).toContain('<Relationship Id="rId7"')
    expect(out).toContain('</Relationships>')
  })

  it('leaves canonical markup untouched and refuses a colliding rename', () => {
    const canonical = `<p:sld xmlns:p="${P}" xmlns:a="x"><p:cSld/></p:sld>`
    expect(canonicalizeNsPrefixes(canonical)).toBe(canonical)
    const colliding = `<ns0:sld xmlns:ns0="${P}" xmlns:p="urn:other"><ns0:cSld/></ns0:sld>`
    expect(canonicalizeNsPrefixes(colliding)).toBe(colliding)
  })
})

describe('PackageArchive with non-canonical prefixes', () => {
  it('reads the slide list from an ElementTree-serialized presentation', async () => {
    const zip = new JSZip()
    zip.file('ppt/presentation.xml', presentationNs0)
    zip.file('ppt/_rels/presentation.xml.rels', relsNs0)
    zip.file('ppt/slides/slide1.xml', `<p:sld xmlns:p="${P}"/>`)
    zip.file('ppt/slides/slide2.xml', `<p:sld xmlns:p="${P}"/>`)
    const archive = await PackageArchive.open(await zip.generateAsync({ type: 'uint8array' }))
    const { size, slidePaths } = archive.readPresentation()
    expect(size).toEqual({ cx: 12192000, cy: 6858000 })
    expect(slidePaths).toEqual(['ppt/slides/slide1.xml', 'ppt/slides/slide2.xml'])
    expect(archive.readText('ppt/slides/slide1.xml')).toBe(`<p:sld xmlns:p="${P}"/>`)
  })
})
