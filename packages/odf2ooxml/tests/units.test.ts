/**
 * A length is offered to the OOXML writers as a plain integer count of twips
 * or EMU. JavaScript switches to exponent notation well before the largest
 * float, and `1.13e+21` is not a valid ST_TwipsMeasure, so a length far past
 * anything a document can hold must be refused rather than written out.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { convertOdtToDocx } from '../src/index'
import { odfLengthToEmu } from '../src/units'

const NAMESPACES =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"' +
  ' xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"' +
  ' xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"' +
  ' xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"'

const HUGE_PAGE =
  '<style:page-layout style:name="PM1"><style:page-layout-properties' +
  ' fo:page-width="2000000000000000000cm" fo:page-height="29.7cm"' +
  ' fo:margin-top="2cm" fo:margin-bottom="2cm" fo:margin-left="2cm" fo:margin-right="2cm"/>' +
  '</style:page-layout>'

function contentXml(body: string): string {
  return (
    `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NAMESPACES} office:version="1.2">` +
    `<office:body><office:text>${body}</office:text></office:body></office:document-content>`
  )
}

function stylesXml(automatic: string): string {
  return (
    `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NAMESPACES} office:version="1.2">` +
    `<office:styles/><office:automatic-styles>${automatic}</office:automatic-styles>` +
    '<office:master-styles><style:master-page style:name="Standard" style:page-layout-name="PM1"/></office:master-styles>' +
    '</office:document-styles>'
  )
}

async function buildOdt(content: string, styles?: string): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file('mimetype', 'application/vnd.oasis.opendocument.text')
  zip.file('content.xml', content)
  if (styles) zip.file('styles.xml', styles)
  return zip.generateAsync({ type: 'uint8array' })
}

async function readDocxEntry(docx: Uint8Array, path: string): Promise<string | null> {
  const zip = await JSZip.loadAsync(docx)
  const entry = zip.file(path)
  return entry ? entry.async('string') : null
}

describe('odfLengthToEmu', () => {
  it('keeps ordinary lengths, signs and the bare zero unchanged', () => {
    expect(odfLengthToEmu('2.5cm')).toBe(900_000)
    expect(odfLengthToEmu('-1cm')).toBe(-360_000)
    expect(odfLengthToEmu('0')).toBe(0)
    expect(odfLengthToEmu('12pt')).toBe(152_400)
    expect(odfLengthToEmu('2pc')).toBe(304_800)
  })

  it('refuses a length past the document-scale cap', () => {
    // the cap is ~27.8 m: a 27.77 m column is still a number, 27.78 m is not paper
    expect(odfLengthToEmu('2777cm')).toBe(999_720_000)
    expect(odfLengthToEmu('2778cm')).toBeNull()
    expect(odfLengthToEmu('2000000000000000000cm')).toBeNull()
    expect(odfLengthToEmu('-2000000000000000000cm')).toBeNull()
  })
})

describe('page geometry from a hostile style', () => {
  it('never writes an exponent into the page size', async () => {
    const result = await convertOdtToDocx(
      await buildOdt(contentXml('<text:p>page</text:p>'), stylesXml(HUGE_PAGE)),
    )
    const document = (await readDocxEntry(result.docx, 'word/document.xml')) ?? ''
    expect(document).not.toContain('e+')
    // a refused size falls back to the blank document's A4 section
    expect(document).toMatch(/<w:pgSz w:w="\d+" w:h="\d+"/)
  })
})
