import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '@genoffice/docx-engine'
import {
  convertLegacyDocToDocx,
  openLegacyDoc,
  parseFontEntries,
  parseFontTable,
  readLegacyDocBody,
  readResolvedLegacyDocModel,
} from '../src/index'
import {
  SYNTHETIC_TEXT_OFFSET,
  buildSyntheticDoc,
  encodeChpxFkp,
  encodeSprm,
  encodeSttbfFfn,
  type SyntheticStyle,
} from './helpers/synthetic-doc'

const FONTS = ['Times New Roman', 'SimSun', 'Arial']

function ftcSprms(ascii?: number, eastAsia?: number, bi?: number): Uint8Array {
  const parts: Uint8Array[] = []
  if (ascii !== undefined)
    parts.push(encodeSprm(0x4a4f, [ascii, 0]), encodeSprm(0x4a51, [ascii, 0]))
  if (eastAsia !== undefined) parts.push(encodeSprm(0x4a50, [eastAsia, 0]))
  if (bi !== undefined) parts.push(encodeSprm(0x4a5e, [bi, 0]))
  return Uint8Array.from(parts.flatMap((part) => [...part]))
}

function fontDoc(options: {
  normalChpx?: Uint8Array
  runSprms?: Uint8Array
  defaults?: [number, number, number]
  fonts?: string[]
  sttbfFfn?: Uint8Array
  lcbSttbfFfn?: number
}): Uint8Array {
  const text = 'Hello\r'
  const styles: SyntheticStyle[] = [
    { sti: 0, name: 'Normal', ...(options.normalChpx ? { chpxSprms: options.normalChpx } : {}) },
  ]
  return buildSyntheticDoc({
    pieces: [{ text }],
    stshStyles: styles,
    ...(options.defaults ? { stshDefaultFonts: options.defaults } : {}),
    ...(options.sttbfFfn ? { sttbfFfn: options.sttbfFfn } : { fonts: options.fonts ?? FONTS }),
    ...(options.lcbSttbfFfn !== undefined ? { lcbSttbfFfn: options.lcbSttbfFfn } : {}),
    ...(options.runSprms
      ? {
          chpxFkpPages: [
            encodeChpxFkp([
              {
                fcStart: SYNTHETIC_TEXT_OFFSET,
                fcEnd: SYNTHETIC_TEXT_OFFSET + text.length * 2,
                grpprl: options.runSprms,
              },
            ]),
          ],
        }
      : {}),
  })
}

async function documentXml(docx: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(docx)
  return zip.file('word/document.xml')!.async('string')
}

describe('font table (SttbfFfn)', () => {
  it('reads the font names by ftc', () => {
    const doc = openLegacyDoc(fontDoc({}))
    const warnings: string[] = []
    expect(parseFontTable(doc.fib, doc.table, warnings)).toEqual(FONTS)
    expect(warnings).toEqual([])
  })

  it('resolves the Normal style fonts and a run override into w:rFonts', async () => {
    const bytes = fontDoc({ normalChpx: ftcSprms(0, 1), runSprms: ftcSprms(2) })
    const model = readResolvedLegacyDocModel(bytes)
    expect(model.paragraphs[0]!.runs[0]!.formatting.fontNames).toEqual({
      ascii: 'Arial',
      eastAsia: 'SimSun',
    })
    const xml = await documentXml((await convertLegacyDocToDocx(bytes)).docx)
    expect(xml).toContain('<w:rFonts w:ascii="Arial" w:eastAsia="SimSun" w:hAnsi="Arial"/>')
    const body = readLegacyDocBody(bytes)
    const block = body.blocks[0]!
    expect(block.type === 'paragraph' && block.paragraph.runs[0]!.formatting.fontNames).toEqual({
      ascii: 'Arial',
      eastAsia: 'SimSun',
    })
  })

  it('falls back to the stylesheet default fonts when no style or run names one', async () => {
    const bytes = fontDoc({ defaults: [2, 1, 2] })
    const model = readResolvedLegacyDocModel(bytes)
    expect(model.paragraphs[0]!.runs[0]!.formatting.fontNames).toEqual({
      ascii: 'Arial',
      eastAsia: 'SimSun',
    })
    const xml = await documentXml((await convertLegacyDocToDocx(bytes)).docx)
    expect(xml).toContain('<w:rFonts w:ascii="Arial" w:eastAsia="SimSun" w:hAnsi="Arial"/>')
  })

  it('writes the complex-script font from sprmCFtcBi', async () => {
    const bytes = fontDoc({ normalChpx: ftcSprms(0, 1, 2) })
    const xml = await documentXml((await convertLegacyDocToDocx(bytes)).docx)
    expect(xml).toContain('w:cs="Arial"')
  })

  it('reads the FFN hints and writes them as word/fontTable.xml', async () => {
    // ffid: prq fixed (1) + ff modern (3 << 4), chs 0x80, ixchSzAlt 3 ("MS " + name), PANOSE, FontSignature
    const name = 'Foo'
    const alt = 'Bar'
    const ffn = [
      0,
      0x31,
      0,
      0,
      0x80,
      name.length + 1,
      ...[2, 2, 6, 9, 4, 2, 5, 8, 3, 4],
      ...[0xff, 0x02, 0, 0xe0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0x9f, 0, 2, 0, 0, 0, 0, 0],
    ]
    for (const char of name + '\0' + alt + '\0') ffn.push(char.charCodeAt(0), 0)
    ffn[0] = ffn.length - 1
    const plain = [0, 0x22, 0, 0, 0, 0, ...new Array(34).fill(0)]
    for (const char of 'Arial\0') plain.push(char.charCodeAt(0), 0)
    plain[0] = plain.length - 1
    const sttbfFfn = Uint8Array.from([2, 0, 0, 0, ...ffn, ...plain])
    const doc = openLegacyDoc(fontDoc({ sttbfFfn }))
    const warnings: string[] = []
    const entries = parseFontEntries(doc.fib, doc.table, warnings)
    expect(warnings).toEqual([])
    expect(entries.map((entry) => entry.name)).toEqual(['Foo', 'Arial'])
    expect(entries[0]).toMatchObject({
      altName: 'Bar',
      family: 'modern',
      pitch: 'fixed',
      charset: 0x80,
      signature: [0xe00002ff, 0, 0, 0, 0x2009f, 0],
    })
    expect(entries[1]).toMatchObject({ family: 'swiss', pitch: 'variable', charset: 0 })
    expect(entries[1]!.altName).toBeUndefined()
    const { docx } = await convertLegacyDocToDocx(fontDoc({ sttbfFfn }))
    const zip = await JSZip.loadAsync(docx)
    const fontTable = await zip.file('word/fontTable.xml')!.async('string')
    expect(fontTable).toContain(
      '<w:font w:name="Foo"><w:altName w:val="Bar"/><w:panose1 w:val="02020609040205080304"/>' +
        '<w:charset w:val="80"/><w:family w:val="modern"/><w:pitch w:val="fixed"/>' +
        '<w:sig w:usb0="E00002FF" w:usb1="00000000" w:usb2="00000000" w:usb3="00000000" w:csb0="0002009F" w:csb1="00000000"/></w:font>',
    )
    expect(fontTable).toContain('<w:font w:name="Arial"><w:panose1 w:val="00000000000000000000"/>')
    expect(await zip.file('[Content_Types].xml')!.async('string')).toContain(
      '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>',
    )
    expect(await zip.file('word/_rels/document.xml.rels')!.async('string')).toMatch(
      /<Relationship Id="rId\d+" Type="http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships\/fontTable" Target="fontTable.xml"\/>/,
    )
    const parsed = await parseDocx(docx)
    expect(parsed.fontTable?.find((entry) => entry.name === 'Foo')).toMatchObject({
      altName: 'Bar',
      charset: '80',
    })
  })

  it('keeps the names read before a truncated font table and warns', () => {
    const full = encodeSttbfFfn(FONTS)
    const doc = openLegacyDoc(fontDoc({ sttbfFfn: full, lcbSttbfFfn: full.length - 10 }))
    const warnings: string[] = []
    expect(parseFontTable(doc.fib, doc.table, warnings)).toEqual(FONTS.slice(0, 2))
    expect(warnings.some((warning) => warning.includes('font 2 overruns'))).toBe(true)
    const model = readResolvedLegacyDocModel(fontDoc({ sttbfFfn: full, lcbSttbfFfn: 2 }))
    expect(model.warnings.some((warning) => warning.includes('font table is invalid'))).toBe(true)
  })
})
