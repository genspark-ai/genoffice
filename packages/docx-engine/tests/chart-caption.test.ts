/** an inline chart sharing its paragraph with a "Figure N:" caption (SEQ field) */
import { describe, expect, it } from 'vitest'
import { parseDocx, type ChartDisplay, type FieldDisplay } from '../src/index'
import { buildDocx, CHART_PARAGRAPH_XML, CHART_PART_XML, CHART_RELS } from './helpers/build-docx'

const CAPTION_RUNS =
  '<w:r><w:t xml:space="preserve">Figure </w:t></w:r>' +
  '<w:fldSimple w:instr=" SEQ Figure \\* ARABIC "><w:r><w:t>1</w:t></w:r></w:fldSimple>' +
  '<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">: Thermal conductivity</w:t></w:r>'

const build = (bodyXml: string) =>
  buildDocx({
    bodyXml,
    extraRels: CHART_RELS,
    extraParts: [
      {
        path: 'word/charts/chart1.xml',
        xml: CHART_PART_XML,
        contentType: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
      },
    ],
  })

type ChartBlock = {
  type: string
  label: string
  chartDisplay?: ChartDisplay
  fieldDisplay?: FieldDisplay
}

describe('chart paragraph with caption field', () => {
  it('keeps the chart and carries the caption as a text field display', async () => {
    const body = CHART_PARAGRAPH_XML.replace('</w:r></w:p>', `</w:r>${CAPTION_RUNS}</w:p>`).replace(
      '<w:p>',
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>',
    )
    const doc = await parseDocx(await build(body))
    const block = doc.blocks[0] as ChartBlock
    expect(block.type).toBe('passthrough')
    expect(block.label).toBe('Chart')
    expect(block.chartDisplay?.series.length).toBeGreaterThan(0)
    expect(block.fieldDisplay?.kind).toBe('text')
    expect(block.fieldDisplay?.left).toBe('Figure 1: Thermal conductivity')
    expect(block.fieldDisplay?.align).toBe('center')
    expect(block.fieldDisplay?.runs?.map((r) => r.text).join('')).toBe(
      'Figure 1: Thermal conductivity',
    )
    expect(block.fieldDisplay?.runs?.at(-1)?.italic).toBe(true)
  })

  it('falls back to the caption field paragraph when the chart part is missing', async () => {
    const body = CHART_PARAGRAPH_XML.replace('</w:r></w:p>', `</w:r>${CAPTION_RUNS}</w:p>`)
    const doc = await parseDocx(await buildDocx({ bodyXml: body, extraRels: CHART_RELS }))
    const block = doc.blocks[0] as ChartBlock
    expect(block.type).toBe('passthrough')
    expect(block.chartDisplay).toBeUndefined()
    expect(block.fieldDisplay?.kind).toBe('text')
    expect(block.fieldDisplay?.left).toBe('Figure 1: Thermal conductivity')
  })

  it('a chart alone carries no field display', async () => {
    const doc = await parseDocx(await build(CHART_PARAGRAPH_XML))
    const block = doc.blocks[0] as ChartBlock
    expect(block.label).toBe('Chart')
    expect(block.fieldDisplay).toBeUndefined()
  })
})
