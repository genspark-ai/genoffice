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

  it('runs without cs info generate exactly the legacy rFonts shape', () => {
    const xml = generateParagraphXml(
      { type: 'paragraph', runs: [{ text: 'plain', fontAscii: 'Georgia' }] },
      GEN_CTX,
    )
    expect(xml).toContain('<w:rFonts w:ascii="Georgia" w:hAnsi="Georgia" w:cs="Georgia"/>')
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
