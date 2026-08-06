/**
 * Dual-slot run fonts (bug: editing the font flattened all four rFonts slots,
 * losing e.g. Times New Roman when a mixed CJK/Latin run was switched to KaiTi).
 */
import { describe, expect, it } from 'vitest'
import { generateParagraphXml, parseDocx, type GenerateContext } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const GEN_CTX: GenerateContext = {
  headingStyleIds: new Map([[1, 'Heading1']]),
  allocateHyperlinkRel: () => 'rId999',
}

const MIXED_RPR =
  '<w:rPr>' +
  '<w:rFonts w:ascii="Times New Roman" w:eastAsia="SimSun" w:hAnsi="Times New Roman" w:cs="Arial" w:hint="eastAsia"/>' +
  '</w:rPr>'
const MIXED_PARA = `<w:p><w:r>${MIXED_RPR}<w:t>合同 Contract</w:t></w:r></w:p>`

async function mixedRun() {
  const doc = await parseDocx(await buildDocx({ bodyXml: MIXED_PARA }))
  return doc.blocks[0].runs![0]
}

describe('rFonts dual-slot model', () => {
  it('parses eastAsia as primary font and ascii as the Latin slot', async () => {
    const run = await mixedRun()
    expect(run.font).toBe('SimSun')
    expect(run.fontAscii).toBe('Times New Roman')
  })

  it('untouched round-trip keeps the original rFonts bytes', async () => {
    const run = await mixedRun()
    const xml = generateParagraphXml({ type: 'paragraph', runs: [{ ...run }] }, GEN_CTX)
    expect(xml).toContain(
      '<w:rFonts w:ascii="Times New Roman" w:eastAsia="SimSun" w:hAnsi="Times New Roman" w:cs="Arial" w:hint="eastAsia"/>',
    )
  })

  it('switching the CJK font keeps the Latin and cs slots (bug 54)', async () => {
    const run = await mixedRun()
    const xml = generateParagraphXml(
      { type: 'paragraph', runs: [{ ...run, font: 'KaiTi' }] },
      GEN_CTX,
    )
    expect(xml).toContain('w:eastAsia="KaiTi"')
    expect(xml).toContain('w:ascii="Times New Roman"')
    expect(xml).toContain('w:hAnsi="Times New Roman"')
    expect(xml).toContain('w:cs="Arial"')
    expect(xml).toContain('w:hint="eastAsia"')
  })

  it('switching the Latin font keeps the eastAsia and cs slots', async () => {
    const run = await mixedRun()
    const xml = generateParagraphXml(
      { type: 'paragraph', runs: [{ ...run, fontAscii: 'Georgia' }] },
      GEN_CTX,
    )
    expect(xml).toContain('w:ascii="Georgia"')
    expect(xml).toContain('w:hAnsi="Georgia"')
    expect(xml).toContain('w:eastAsia="SimSun"')
    expect(xml).toContain('w:cs="Arial"')
  })

  it('Latin edit on an ascii-only run does not invent an eastAsia slot', async () => {
    const para =
      '<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>' +
      '<w:t>latin only</w:t></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    const run = doc.blocks[0].runs![0]
    expect(run.font).toBe('Arial') // parse-side derivation, not a real eastAsia slot
    const xml = generateParagraphXml(
      { type: 'paragraph', runs: [{ ...run, fontAscii: 'Georgia' }] },
      GEN_CTX,
    )
    expect(xml).toContain('w:ascii="Georgia"')
    expect(xml).toContain('w:hAnsi="Georgia"')
    expect(xml).not.toContain('eastAsia')
  })

  it('CJK edit on an ascii-only run adds eastAsia and keeps the Latin slot', async () => {
    const para =
      '<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>' +
      '<w:t>latin only</w:t></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    const run = doc.blocks[0].runs![0]
    const xml = generateParagraphXml(
      { type: 'paragraph', runs: [{ ...run, font: 'KaiTi' }] },
      GEN_CTX,
    )
    expect(xml).toContain('w:eastAsia="KaiTi"')
    expect(xml).toContain('w:ascii="Arial"')
    expect(xml).toContain('w:hAnsi="Arial"')
  })

  it('an explicit Latin font wins over a leftover theme attribute', async () => {
    const para =
      '<w:p><w:r><w:rPr>' +
      '<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsia="SimSun" w:hAnsiTheme="minorHAnsi"/>' +
      '</w:rPr><w:t>合同 Contract</w:t></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: para }))
    const run = doc.blocks[0].runs![0]
    const xml = generateParagraphXml(
      { type: 'paragraph', runs: [{ ...run, fontAscii: 'Arial' }] },
      GEN_CTX,
    )
    expect(xml).toContain('w:ascii="Arial"')
    expect(xml).toContain('w:hAnsi="Arial"')
    expect(xml).not.toContain('asciiTheme')
    expect(xml).not.toContain('hAnsiTheme')
    expect(xml).toContain('w:eastAsia="SimSun"')
  })

  it('setting only the primary font on a run without rFonts fills every slot (legacy)', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>plain text</w:t></w:r></w:p>' }),
    )
    const run = doc.blocks[0].runs![0]
    const xml = generateParagraphXml(
      { type: 'paragraph', runs: [{ ...run, font: 'KaiTi' }] },
      GEN_CTX,
    )
    expect(xml).toContain(
      '<w:rFonts w:ascii="KaiTi" w:eastAsia="KaiTi" w:hAnsi="KaiTi" w:cs="KaiTi"/>',
    )
  })

  it('setting only the Latin font on a run without rFonts leaves eastAsia empty', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>plain text</w:t></w:r></w:p>' }),
    )
    const run = doc.blocks[0].runs![0]
    const xml = generateParagraphXml(
      { type: 'paragraph', runs: [{ ...run, fontAscii: 'Arial' }] },
      GEN_CTX,
    )
    expect(xml).toContain('<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>')
    expect(xml).not.toContain('eastAsia')
  })
})
