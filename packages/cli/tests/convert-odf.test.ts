import { existsSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { xlsxSidecarPath } from '../src/resources'
import { run, tempDir } from './helpers'

const ODF_NAMESPACES =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"' +
  ' xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"' +
  ' xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"' +
  ' xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"' +
  ' xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"' +
  ' xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"' +
  ' xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"' +
  ' xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"' +
  ' xmlns:xlink="http://www.w3.org/1999/xlink"'

async function buildZip(files: Record<string, string>): Promise<Uint8Array> {
  const zip = new JSZip()
  for (const [path, body] of Object.entries(files)) zip.file(path, body)
  return zip.generateAsync({ type: 'uint8array' })
}

async function entryText(path: string, entry: string): Promise<string> {
  const zip = await JSZip.loadAsync(
    await import('node:fs/promises').then((fs) => fs.readFile(path)),
  )
  const file = zip.file(entry)
  return file ? file.async('string') : ''
}

describe('genoffice convert ODF import filters', () => {
  it('converts .odt to .docx headlessly', async () => {
    const dir = tempDir()
    const content =
      `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${ODF_NAMESPACES} office:version="1.2">` +
      '<office:body><office:text>' +
      '<text:h text:outline-level="1">Imported heading</text:h>' +
      '<text:p>plain <text:span text:style-name="T1">bold</text:span></text:p>' +
      '</office:text></office:body></office:document-content>'
    const styles =
      `<?xml version="1.0"?><office:document-styles ${ODF_NAMESPACES}>` +
      '<office:styles><style:style style:name="T1" style:family="text">' +
      '<style:text-properties fo:font-weight="bold"/></style:style></office:styles>' +
      '</office:document-styles>'
    const odt = join(dir, 'report.odt')
    const { writeFile } = await import('node:fs/promises')
    await writeFile(odt, await buildZip({ 'content.xml': content, 'styles.xml': styles }))

    const out = join(dir, 'report.docx')
    const r = await run(['convert', odt, '--to', 'docx', '--out', out, '--json'])
    expect(r.code).toBe(0)
    expect(existsSync(out)).toBe(true)
    const document = await entryText(out, 'word/document.xml')
    expect(document).toContain('Imported heading')
    expect(document).toContain('<w:b/>')
    expect(r.json().detail.dropped).toEqual([])
  })

  it('converts .odp to .pptx headlessly', async () => {
    const dir = tempDir()
    const content =
      `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${ODF_NAMESPACES} office:version="1.2">` +
      '<office:body><office:presentation>' +
      '<draw:page draw:name="p1">' +
      '<draw:frame svg:x="2cm" svg:y="1cm" svg:width="20cm" svg:height="3cm">' +
      '<draw:text-box><text:p>Slide one words</text:p></draw:text-box>' +
      '</draw:frame></draw:page>' +
      '</office:presentation></office:body></office:document-content>'
    const odp = join(dir, 'deck.odp')
    const { writeFile } = await import('node:fs/promises')
    await writeFile(odp, await buildZip({ 'content.xml': content }))

    const out = join(dir, 'deck.pptx')
    const r = await run(['convert', odp, '--to', 'pptx', '--out', out, '--json'])
    expect(r.code).toBe(0)
    expect(existsSync(out)).toBe(true)
    const slide1 = await entryText(out, 'ppt/slides/slide1.xml')
    expect(slide1).toContain('Slide one words')
    expect(r.json().detail).toMatchObject({ slides: 1 })
  })

  it('converts .ott/.otp templates through the same filters as .odt/.odp', async () => {
    const dir = tempDir()
    const { writeFile } = await import('node:fs/promises')
    const textContent =
      `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${ODF_NAMESPACES} office:version="1.2">` +
      '<office:body><office:text><text:h text:outline-level="1">Templated</text:h></office:text></office:body></office:document-content>'
    const ott = join(dir, 'letter.ott')
    await writeFile(ott, await buildZip({ 'content.xml': textContent }))
    const ottOut = join(dir, 'letter.docx')
    expect((await run(['convert', ott, '--to', 'docx', '--out', ottOut, '--json'])).code).toBe(0)
    expect(await entryText(ottOut, 'word/document.xml')).toContain('Templated')

    const presContent =
      `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${ODF_NAMESPACES} office:version="1.2">` +
      '<office:body><office:presentation><draw:page draw:name="p1">' +
      '<draw:frame svg:x="1cm" svg:y="1cm" svg:width="10cm" svg:height="2cm"><draw:text-box><text:p>Deck template</text:p></draw:text-box></draw:frame>' +
      '</draw:page></office:presentation></office:body></office:document-content>'
    const otp = join(dir, 'deck.otp')
    await writeFile(otp, await buildZip({ 'content.xml': presContent }))
    const otpOut = join(dir, 'deck.pptx')
    expect((await run(['convert', otp, '--to', 'pptx', '--out', otpOut, '--json'])).code).toBe(0)
    expect(await entryText(otpOut, 'ppt/slides/slide1.xml')).toContain('Deck template')
  })

  it.skipIf(!xlsxSidecarPath())('converts .ots through the workbook sidecar route', async () => {
    const dir = tempDir()
    const { writeFile } = await import('node:fs/promises')
    const sheetContent =
      `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${ODF_NAMESPACES} office:version="1.2">` +
      '<office:body><office:spreadsheet><table:table table:name="Probe"><table:table-row>' +
      '<table:table-cell office:value-type="float" office:value="10"><text:p>10</text:p></table:table-cell>' +
      '</table:table-row></table:table></office:spreadsheet></office:body></office:document-content>'
    const ots = join(dir, 'grid.ots')
    await writeFile(
      ots,
      await buildZip({
        mimetype: 'application/vnd.oasis.opendocument.spreadsheet-template',
        'META-INF/manifest.xml':
          '<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet-template"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>',
        'content.xml': sheetContent,
      }),
    )
    const otsOut = join(dir, 'grid.xlsx')
    const r = await run(['convert', ots, '--to', 'xlsx', '--out', otsOut, '--json'])
    expect(r.code).toBe(0)
    expect(r.json().detail).toMatchObject({ sheets: 1 })
  })

  it('refuses ODF routes that do not exist', async () => {
    const dir = tempDir()
    const odt = join(dir, 'report.odt')
    const { writeFile } = await import('node:fs/promises')
    await writeFile(odt, await buildZip({ 'content.xml': '<?xml version="1.0"?><x/>' }))
    const r = await run(['convert', odt, '--to', 'pptx', '--json'])
    expect(r.code).toBe(1)
    expect(r.json().message).toContain('cannot convert')
  })
})
