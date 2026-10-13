/**
 * A zip entry is inflated on the main-process heap with nothing to stop it, so
 * the size the central directory declares is checked before the inflate. The
 * app reports the refusal as an unreadable file rather than dying with the
 * heap full.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  convertOdpToPptx,
  convertOdtToDocx,
  MAX_ODF_PART_BYTES,
  OdfConvertError,
} from '../src/index'

const NAMESPACES =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"' +
  ' xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"' +
  ' xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"' +
  ' xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"' +
  ' xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"' +
  ' xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"' +
  ' xmlns:xlink="http://www.w3.org/1999/xlink"'

const PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

const PARAGRAPH = '<text:p>hello</text:p>'
const IMAGE_FRAME =
  '<text:p><draw:frame svg:x="0cm" svg:y="0cm" svg:width="2cm" svg:height="2cm">' +
  '<draw:image xlink:href="Pictures/pic.png"/></draw:frame></text:p>'
const SLIDE_FRAME =
  '<draw:page draw:name="page1">' +
  '<draw:frame draw:layer="layout" svg:x="1cm" svg:y="1cm" svg:width="2cm" svg:height="2cm">' +
  '<draw:image xlink:href="Pictures/pic.png"/></draw:frame>' +
  '</draw:page>'

/**
 * Rewrite the uncompressed size one entry declares, in the local header and in
 * the central directory, the way an archive that inflates to more than it says
 * would. Building a real multi-hundred-megabyte entry would cost more than the
 * assertion is worth.
 */
function forgeDeclaredSize(archive: Uint8Array, entry: string, size: number): Uint8Array {
  const bytes = archive.slice()
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const name = new TextEncoder().encode(entry)
  const at = (offset: number) => offset >= 0 && offset + name.length <= bytes.length
  const matches = (offset: number) => name.every((byte, i) => bytes[offset + i] === byte)
  for (let i = 0; i + 30 + name.length <= bytes.length; i += 1) {
    // local file header: uncompressed size at 22, name length at 26, name at 30
    if (
      view.getUint32(i, true) === 0x04034b50 &&
      view.getUint16(i + 26, true) === name.length &&
      matches(i + 30)
    ) {
      view.setUint32(i + 22, size, true)
    }
    // central directory: uncompressed size at 24, name length at 28, name at 46
    if (
      view.getUint32(i, true) === 0x02014b50 &&
      view.getUint16(i + 28, true) === name.length &&
      matches(i + 46) &&
      at(i + 46)
    ) {
      view.setUint32(i + 24, size, true)
    }
  }
  return bytes
}

const pastTheCap = (archive: Uint8Array, entry: string): Uint8Array =>
  forgeDeclaredSize(archive, entry, MAX_ODF_PART_BYTES + 1)

async function buildOdt(oversized: 'content' | 'image' | null): Promise<Uint8Array> {
  const body = PARAGRAPH + (oversized === 'image' ? IMAGE_FRAME : '')
  const zip = new JSZip()
  const content = `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NAMESPACES} office:version="1.2"><office:body><office:text>${body}</office:text></office:body></office:document-content>`
  zip.file('content.xml', content)
  if (oversized === 'image') zip.file('Pictures/pic.png', PNG_BASE64, { base64: true })
  const archive = await zip.generateAsync({ type: 'uint8array' })
  if (oversized === 'content') return pastTheCap(archive, 'content.xml')
  if (oversized === 'image') return pastTheCap(archive, 'Pictures/pic.png')
  return archive
}

async function buildOdp(oversized: boolean): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file(
    'content.xml',
    `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NAMESPACES} office:version="1.2"><office:body><office:presentation>${SLIDE_FRAME}</office:presentation></office:body></office:document-content>`,
  )
  zip.file('Pictures/pic.png', PNG_BASE64, { base64: true })
  const archive = await zip.generateAsync({ type: 'uint8array' })
  return oversized ? pastTheCap(archive, 'Pictures/pic.png') : archive
}

describe('oversized zip entries', () => {
  it('still converts a package whose parts are within the cap', async () => {
    // the gate reads the declared size, so a normal document must be untouched
    const result = await convertOdtToDocx(await buildOdt(null))
    expect(result.dropped).toEqual([])
  })

  it('refuses an oversized content part instead of inflating it', async () => {
    await expect(convertOdtToDocx(await buildOdt('content'))).rejects.toThrow(OdfConvertError)
  })

  it('refuses an oversized image instead of inflating it', async () => {
    // an image that fails to load is otherwise counted as dropped, which would
    // hand the user a document with its pictures missing and no explanation
    await expect(convertOdtToDocx(await buildOdt('image'))).rejects.toThrow(
      new OdfConvertError('unreadable'),
    )
  })

  it('refuses the same archive on the presentation route', async () => {
    await expect(convertOdpToPptx(await buildOdp(true))).rejects.toThrow(OdfConvertError)
  })
})
