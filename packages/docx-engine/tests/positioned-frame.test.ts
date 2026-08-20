import { describe, expect, it } from 'vitest'
import {
  generateParagraphXml,
  generateTableModelXml,
  mergePPrFormat,
  parseDocx,
  saveDocx,
  type GenerateContext,
  type ParaFormat,
  type TableModel,
} from '../src/index'
import { buildDocx } from './helpers/build-docx'

const GEN_CTX: GenerateContext = {
  headingStyleIds: new Map([[1, 'Heading1']]),
  allocateHyperlinkRel: () => 'rId999',
}

describe('positioned paragraph frames (w:framePr)', () => {
  it('emits a page-anchored frame with position, size and wrap', () => {
    const format: ParaFormat = {
      frame: { wTwips: 4000, xTwips: 1200, yTwips: 800 },
      lineRule: 'exact',
      lineRawTwips: 300,
      spaceAfter: 0,
    }
    const xml = generateParagraphXml(
      { type: 'paragraph', runs: [{ text: 'slide title' }], format },
      GEN_CTX,
    )
    expect(xml).toContain('<w:framePr')
    expect(xml).toContain('w:w="4000"')
    expect(xml).toContain('w:x="1200"')
    expect(xml).toContain('w:y="800"')
    expect(xml).toContain('w:hAnchor="page"')
    expect(xml).toContain('w:vAnchor="page"')
    expect(xml).toContain('w:wrap="none"')
    // no height attrs unless requested (auto height must not clip substituted fonts)
    expect(xml).not.toContain('w:h=')
    expect(xml).not.toContain('w:hRule=')
    // schema order: framePr before spacing inside pPr
    expect(xml.indexOf('<w:framePr')).toBeLessThan(xml.indexOf('<w:spacing'))
  })

  it('emits height, hRule, custom anchors and wrap when set', () => {
    const format: ParaFormat = {
      frame: {
        wTwips: 2000,
        hTwips: 600,
        hRule: 'exact',
        xTwips: 0,
        yTwips: 0,
        hAnchor: 'margin',
        vAnchor: 'margin',
        wrap: 'around',
      },
    }
    const xml = generateParagraphXml({ type: 'paragraph', runs: [{ text: 'x' }], format }, GEN_CTX)
    expect(xml).toContain('w:h="600"')
    expect(xml).toContain('w:hRule="exact"')
    expect(xml).toContain('w:hAnchor="margin"')
    expect(xml).toContain('w:vAnchor="margin"')
    expect(xml).toContain('w:wrap="around"')
    expect(xml).toContain('w:x="0"')
    expect(xml).toContain('w:y="0"')
  })

  it('keeps CT_PPr child order when built from the model alone', () => {
    const merged = mergePPrFormat('<w:pPr/>', {
      frame: { wTwips: 1000, xTwips: 10, yTwips: 20 },
      pageBreakBefore: true,
      align: 'center',
      lineRule: 'exact',
      lineRawTwips: 240,
    })
    // pageBreakBefore (rank 3) < framePr (rank 4) < spacing < jc
    expect(merged.indexOf('<w:pageBreakBefore')).toBeLessThan(merged.indexOf('<w:framePr'))
    expect(merged.indexOf('<w:framePr')).toBeLessThan(merged.indexOf('<w:spacing'))
    expect(merged.indexOf('<w:spacing')).toBeLessThan(merged.indexOf('<w:jc'))
  })

  it('frame wins over dropCap when both are set (single w:framePr)', () => {
    const xml = generateParagraphXml(
      {
        type: 'paragraph',
        runs: [{ text: 'x' }],
        format: {
          frame: { wTwips: 1000, xTwips: 10, yTwips: 20 },
          dropCap: { type: 'drop', lines: 3 },
        },
      },
      GEN_CTX,
    )
    expect(xml.match(/<w:framePr/g)).toHaveLength(1)
    expect(xml).toContain('w:x="10"')
    expect(xml).not.toContain('w:dropCap')
  })

  it('leaves an existing non-dropCap framePr untouched when the format has no frame', async () => {
    const rawPPr =
      '<w:pPr><w:framePr w:w="5000" w:hAnchor="page" w:vAnchor="page" w:x="100" w:y="200" w:wrap="none"/></w:pPr>'
    // no frame/dropCap in the format → w:framePr is unmanaged → raw bytes survive
    expect(mergePPrFormat(rawPPr, { align: 'center' })).toContain(
      '<w:framePr w:w="5000" w:hAnchor="page" w:vAnchor="page" w:x="100" w:y="200" w:wrap="none"/>',
    )
  })

  it('round-trips a framed paragraph through saveDocx without corrupting the body', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>seed</w:t></w:r></w:p>' }),
    )
    const saved = await saveDocx(doc, [
      {
        kind: 'generated',
        block: {
          type: 'paragraph',
          runs: [{ text: 'framed' }],
          format: { frame: { wTwips: 3000, xTwips: 500, yTwips: 700 } },
        },
      },
    ])
    const reparsed = await parseDocx(saved)
    const texts = reparsed.blocks.map((b) => (b.runs ?? []).map((r) => r.text).join(''))
    expect(texts).toContain('framed')
  })
})

describe('floating table position (w:tblpPr)', () => {
  const model: TableModel = {
    rows: [[{ paras: ['a'] }, { paras: ['b'] }]],
    colWidthsTwips: [1000, 1000],
    floatPos: { xTwips: 2400, yTwips: 1800 },
  }

  it('emits page-anchored tblpPr with overlap before tblW', () => {
    const xml = generateTableModelXml(model)
    expect(xml).toContain('<w:tblpPr')
    expect(xml).toContain('w:tblpX="2400"')
    expect(xml).toContain('w:tblpY="1800"')
    expect(xml).toContain('w:horzAnchor="page"')
    expect(xml).toContain('w:vertAnchor="page"')
    expect(xml).toContain('<w:tblOverlap w:val="overlap"/>')
    expect(xml.indexOf('<w:tblpPr')).toBeLessThan(xml.indexOf('<w:tblW'))
    expect(xml.indexOf('<w:tblpPr')).toBeLessThan(xml.indexOf('<w:tblOverlap'))
  })

  it('omits tblpPr without floatPos (existing callers unchanged)', () => {
    const xml = generateTableModelXml({ ...model, floatPos: undefined })
    expect(xml).not.toContain('tblpPr')
    expect(xml).not.toContain('tblOverlap')
  })
})
