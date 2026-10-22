/**
 * Run-level RTL (w:rtl) + complex-script font slot (rFonts w:cs) — added for
 * pdf2docx P2. The round-trip rule: untouched runs keep their original bytes;
 * runs without cs info generate exactly what they did before.
 */
import { describe, expect, it } from 'vitest'
import { generateParagraphXml, parseDocx, saveDocx, type GenerateContext } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const GEN_CTX: GenerateContext = {
  headingStyleIds: new Map([[1, 'Heading1']]),
  allocateHyperlinkRel: () => 'rId999',
}

const AR_RPR =
  '<w:rPr>' +
  '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Traditional Arabic"/>' +
  '<w:sz w:val="24"/><w:szCs w:val="24"/><w:rtl/>' +
  '</w:rPr>'
const AR_BODY =
  `<w:p><w:pPr><w:bidi/></w:pPr><w:r>${AR_RPR}<w:t>مرحبا</w:t></w:r></w:p>` +
  '<w:p><w:r><w:rPr><w:rtl w:val="0"/></w:rPr><w:t>ltr</w:t></w:r></w:p>'

describe('run-level w:rtl + rFonts w:cs slot', () => {
  it('parses w:rtl and the literal w:cs font into the run model', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: AR_BODY }))
    const run = doc.blocks[0].runs![0]
    expect(run.rtl).toBe(true)
    expect(run.fontCs).toBe('Traditional Arabic')
    expect(run.fontAscii).toBe('Arial')
    // negated rtl parses as explicit false
    expect(doc.blocks[1].runs![0].rtl).toBe(false)
  })

  it('keeps untouched RTL runs byte-identical through save', async () => {
    const bytes = await buildDocx({ bodyXml: AR_BODY })
    const doc = await parseDocx(bytes)
    const blocks = doc.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! }))
    expect(await saveDocx(doc, blocks)).toEqual(bytes)
  })

  it('re-generating a parsed RTL run keeps its original rPr bytes (group equality)', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: AR_BODY }))
    const run = doc.blocks[0].runs![0]
    const xml = generateParagraphXml({ type: 'paragraph', runs: [{ ...run }] }, GEN_CTX)
    expect(xml).toContain(AR_RPR)
  })

  it('generates w:rtl and the cs slot for model-only runs, in schema order', () => {
    const xml = generateParagraphXml(
      {
        type: 'paragraph',
        format: { bidi: true },
        runs: [
          {
            text: 'שלום',
            rtl: true,
            fontAscii: 'David',
            fontCs: 'David',
            sizeHalfPoints: 24,
            vertAlign: 'superscript',
          },
        ],
      },
      GEN_CTX,
    )
    expect(xml).toContain('<w:rFonts w:ascii="David" w:hAnsi="David" w:cs="David"/>')
    // schema order: vertAlign before rtl
    expect(xml).toMatch(/<w:vertAlign w:val="superscript"\/><w:rtl\/>/)
    expect(xml).toContain('<w:bidi/>')
  })

  it('a Latin-only font keeps the complex-script font inherited', () => {
    const xml = generateParagraphXml(
      { type: 'paragraph', runs: [{ text: 'plain', fontAscii: 'Georgia' }] },
      GEN_CTX,
    )
    expect(xml).toContain('<w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/>')
    expect(xml).not.toContain('<w:rtl/>')
  })

  it('editing the cs font rebuilds only that slot inside the original rFonts', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: AR_BODY }))
    const run = doc.blocks[0].runs![0]
    const xml = generateParagraphXml(
      { type: 'paragraph', runs: [{ ...run, fontCs: 'Simplified Arabic' }] },
      GEN_CTX,
    )
    expect(xml).toContain('<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Simplified Arabic"/>')
    expect(xml).toContain('<w:rtl/>')
  })
})

describe('table cell paragraphs carry w:bidi (visual↔logical jc swap)', () => {
  it('emits bidi + swapped jc for RTL cell paragraphs', async () => {
    const { generateTableModelXml } = await import('../src/index')
    const xml = generateTableModelXml({
      rows: [
        [
          {
            paras: ['مرحبا'],
            richParas: [{ bidi: true, align: 'left', runs: [{ text: 'مرحبا', rtl: true }] }],
          },
          { paras: ['plain'], richParas: [{ align: 'left', runs: [{ text: 'plain' }] }] },
        ],
      ],
    })
    // visual left maps back to logical right inside a bidi cell paragraph
    expect(xml).toContain('<w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr>')
    // non-bidi cell paragraphs are untouched
    expect(xml).toContain('<w:pPr><w:jc w:val="left"/></w:pPr>')
    expect(xml).toContain('<w:rtl/>')
  })
})

// Word reads an rtl run's size/weight from the Cs twins all the way up the chain
// (prod corpus 20261021 sample 007: a Heading1 `<w:b/><w:sz 32/>` run carrying only
// `<w:rtl/>` is 11pt regular; sample 023: a `<w:sz 36/><w:rtl/>` paragraph mark under
// a szCs=24 style lays a 12pt empty line)
describe('rtl runs inherit the Cs twins, not the base set', () => {
  const STYLES =
    '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>' +
    '<w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="BodyText"><w:name w:val="Body Text"/>' +
    '<w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>' +
    '<w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/>' +
    '<w:rPr><w:bCs/><w:szCs w:val="28"/></w:rPr></w:style>'
  const wrap = (styles: string) =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${styles}</w:styles>`

  it('a bare rtl run under a bold 16pt heading style is 11pt regular', async () => {
    const body =
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:bidi/></w:pPr>' +
      '<w:r><w:rPr><w:rtl/></w:rPr><w:t>שלום</w:t></w:r>' +
      '<w:r><w:rPr><w:rStyle w:val="Strong"/><w:rtl/></w:rPr><w:t>עולם</w:t></w:r>' +
      '<w:r><w:t>Latin</w:t></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: body, stylesXml: wrap(STYLES) }))
    const [plain, strong, latin] = doc.blocks[0].runs!
    expect(plain.cs).toBe(true)
    expect(plain.sizeHalfPoints).toBeUndefined()
    const cs = (r: typeof plain) => ({
      sizeHalfPoints: r.effective!.sizeHalfPoints,
      bold: r.effective!.bold,
      italic: r.effective!.italic,
    })
    expect(plain.effective!.cs).toBe(true)
    expect(cs(plain)).toEqual({ sizeHalfPoints: 22, bold: false, italic: false })
    expect(cs(strong)).toEqual({ sizeHalfPoints: 28, bold: true, italic: false })
    expect(latin.cs).toBeUndefined()
    expect(latin.effective!.cs).toBe(false)
  })

  it('an rtl paragraph mark sizes the empty line by szCs, falling back to the style chain', async () => {
    const body =
      '<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:rPr><w:b/><w:sz w:val="36"/><w:rtl/></w:rPr></w:pPr></w:p>' +
      '<w:p><w:pPr><w:rPr><w:sz w:val="36"/><w:szCs w:val="40"/><w:rtl/></w:rPr></w:pPr></w:p>' +
      '<w:p><w:pPr><w:rPr><w:sz w:val="36"/><w:szCs w:val="40"/></w:rPr></w:pPr></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: body, stylesXml: wrap(STYLES) }))
    const sizes = doc.blocks.slice(0, 3).map((b) => b.format?.emptyRunSizeHalfPoints)
    expect(sizes).toEqual([24, 40, 36])
  })
})
