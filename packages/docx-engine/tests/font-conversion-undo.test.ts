/**
 * Convert Permanently at the save, and what a later save does once the
 * conversion was undone: the fonts come back in every part the save changed.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { writeFontConversions } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const fonts = (name: string) => `<w:rPr><w:rFonts w:ascii="${name}" w:hAnsi="${name}"/></w:rPr>`
const run = (name: string, text: string) => `<w:r>${fonts(name)}<w:t>${text}</w:t></w:r>`

const PARTS = [
  'word/document.xml',
  'word/styles.xml',
  'word/header1.xml',
  'word/footer1.xml',
  'word/footnotes.xml',
  'word/numbering.xml',
  'word/theme/theme1.xml',
]

/** a document that names Old Sans in every kind of part, and Arial beside it in each */
const source = () =>
  buildDocx({
    bodyXml: `<w:p>${run('Old Sans', 'body')}${run('Arial', 'was Arial')}</w:p>`,
    extraStylesXml:
      `<w:style w:type="paragraph" w:styleId="Fancy"><w:name w:val="Fancy"/>${fonts('Old Sans')}</w:style>` +
      `<w:style w:type="paragraph" w:styleId="Plain"><w:name w:val="Plain"/>${fonts('Arial')}</w:style>`,
    headerXml: `<w:p>${run('Old Sans', 'head')}${run('Arial', 'head too')}</w:p>`,
    footerXml: `<w:p>${run('Old Sans', 'foot')}${run('Arial', 'foot too')}</w:p>`,
    numberingXml:
      `<w:numbering ${W}><w:abstractNum w:abstractNumId="0">` +
      `<w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="o"/>${fonts('Old Sans')}</w:lvl>` +
      `<w:lvl w:ilvl="1"><w:numFmt w:val="bullet"/><w:lvlText w:val="o"/>${fonts('Arial')}</w:lvl>` +
      `</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
    extraParts: [
      {
        path: 'word/footnotes.xml',
        xml: `<w:footnotes ${W}><w:footnote w:id="1"><w:p>${run('Old Sans', 'note')}${run('Arial', 'note too')}</w:p></w:footnote></w:footnotes>`,
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
      },
      {
        path: 'word/theme/theme1.xml',
        xml:
          '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="x">' +
          '<a:majorFont><a:latin typeface="Old Sans"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont>' +
          '</a:fontScheme></a:themeElements></a:theme>',
        contentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
      },
    ],
  })

const partsOf = async (bytes: Uint8Array): Promise<Record<string, string>> => {
  const zip = await JSZip.loadAsync(bytes)
  const out: Record<string, string> = {}
  for (const name of PARTS) out[name] = (await zip.file(name)?.async('string')) ?? ''
  return out
}

const withPart = async (bytes: Uint8Array, name: string, edit: (xml: string) => string) => {
  const zip = await JSZip.loadAsync(bytes)
  zip.file(name, edit(await zip.file(name)!.async('string')))
  return zip.generateAsync({ type: 'uint8array' })
}

const CONVERT = { 'Old Sans': 'Arial' }

describe('a font conversion written at the save', () => {
  it('names every part it changed, and each of them loses the missing font', async () => {
    const original = await source()
    const before = await partsOf(original)
    for (const name of PARTS) expect(before[name], name).toContain('Old Sans')
    const { bytes, written } = await writeFontConversions(original, CONVERT)
    const after = await partsOf(bytes)
    for (const name of PARTS) expect(after[name], name).not.toContain('Old Sans')
    expect(written).toHaveLength(1)
    expect(Object.keys(written[0].parts).sort()).toEqual([...PARTS].sort())
    expect(written[0].kept).toBe(written[0].fonts)
  })

  it('is taken back out of every part by the save after its undo, leaving the fonts that were the stand-in all along', async () => {
    const original = await source()
    const before = await partsOf(original)
    const converted = await writeFontConversions(original, CONVERT)
    const undone = converted.written.map((w) => ({ ...w, kept: {} }))
    const { bytes, written } = await writeFontConversions(converted.bytes, {}, undone)
    expect(await partsOf(bytes)).toEqual(before)
    // nothing of the conversion is left to undo
    expect(written).toEqual([])
  })

  it('leaves a save with nothing undone alone', async () => {
    const converted = await writeFontConversions(await source(), CONVERT)
    const again = await writeFontConversions(converted.bytes, {}, converted.written)
    expect(again.bytes).toBe(converted.bytes)
    expect(again.written).toEqual(converted.written)
  })

  it('keeps a part converted once it changed after the conversion, and restores the others', async () => {
    const original = await source()
    const before = await partsOf(original)
    const converted = await writeFontConversions(original, CONVERT)
    const edited = await withPart(converted.bytes, 'word/header1.xml', (xml) =>
      xml.replace('head too', 'head, edited'),
    )
    const undone = converted.written.map((w) => ({ ...w, kept: {} }))
    const after = await partsOf((await writeFontConversions(edited, {}, undone)).bytes)
    expect(after['word/header1.xml']).not.toContain('Old Sans')
    expect(after['word/header1.xml']).toContain('head, edited')
    for (const name of PARTS.filter((part) => part !== 'word/header1.xml'))
      expect(after[name], name).toBe(before[name])
  })

  it('takes back one of two fonts converted together and keeps the other', async () => {
    const original = await withPart(await source(), 'word/styles.xml', (xml) =>
      xml.replace(
        '</w:styles>',
        `<w:style w:type="paragraph" w:styleId="Gone"><w:name w:val="Gone"/>${fonts('Lost Serif')}</w:style></w:styles>`,
      ),
    )
    const both = { 'Old Sans': 'Arial', 'Lost Serif': 'Georgia' }
    const converted = await writeFontConversions(original, both)
    const undone = converted.written.map((w) => ({ ...w, kept: { 'Old Sans': 'Arial' } }))
    const { bytes, written } = await writeFontConversions(converted.bytes, {}, undone)
    const styles = (await partsOf(bytes))['word/styles.xml']
    expect(styles).toContain('Lost Serif')
    expect(styles).not.toContain('Georgia')
    expect(styles).not.toContain('Old Sans')
    expect(written).toHaveLength(1)
    expect(written[0].fonts).toEqual({ 'Old Sans': 'Arial' })
    expect(written[0].kept).toBe(written[0].fonts)

    // and the one that was kept can still be undone after that save
    const rest = written.map((w) => ({ ...w, kept: {} }))
    expect(await partsOf((await writeFontConversions(bytes, {}, rest)).bytes)).toEqual(
      await partsOf(original),
    )
  })

  it('undoes two conversions of two saves in one save, the later first', async () => {
    const original = await withPart(await source(), 'word/styles.xml', (xml) =>
      xml.replace(
        '</w:styles>',
        `<w:style w:type="paragraph" w:styleId="Gone"><w:name w:val="Gone"/>${fonts('Lost Serif')}</w:style></w:styles>`,
      ),
    )
    const first = await writeFontConversions(original, CONVERT)
    const second = await writeFontConversions(first.bytes, { 'Lost Serif': 'Arial' }, first.written)
    expect(second.written).toHaveLength(2)
    const undone = second.written.map((w) => ({ ...w, kept: {} }))
    const { bytes, written } = await writeFontConversions(second.bytes, {}, undone)
    expect(await partsOf(bytes)).toEqual(await partsOf(original))
    expect(written).toEqual([])
  })

  it('converts again on a redo after the undo was saved', async () => {
    const original = await source()
    const converted = await writeFontConversions(original, CONVERT)
    const undone = await writeFontConversions(
      converted.bytes,
      {},
      converted.written.map((w) => ({ ...w, kept: {} })),
    )
    const redone = await writeFontConversions(undone.bytes, CONVERT, undone.written)
    expect(await partsOf(redone.bytes)).toEqual(await partsOf(converted.bytes))
    expect(redone.written).toHaveLength(1)
  })
})
