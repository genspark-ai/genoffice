import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  convertLegacyDocToDocx,
  openLegacyDoc,
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
