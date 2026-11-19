import { describe, expect, it } from 'vitest'
import { inlinePictureDrawingXml } from '@genoffice/docx-engine'
import { inlineToRuns, type PmNode } from '../src/renderer/editor/convert'

/** a header-cell / table-cell picture dragged to a new size saves that size into its drawing */
describe('inline picture size on save', () => {
  const xml = inlinePictureDrawingXml({ rId: 'rId1', docPrId: 1, widthPx: 158.74, heightPx: 39.44 })
  const node = (widthPx: number, heightPx: number): PmNode => ({
    type: 'docInlineImage',
    attrs: { dataUrl: 'data:image/png;base64,AA==', widthPx, heightPx, xml },
  })

  it('keeps the original drawing when the size is untouched', () => {
    expect(inlineToRuns([node(158.74, 39.44)])[0].image?.xml).toBe(xml)
  })

  it('writes the resized extent', () => {
    const run = inlineToRuns([node(240, 60)])[0]
    expect(run.image?.widthPx).toBe(240)
    expect(run.image?.xml).toContain('<wp:extent cx="2286000" cy="571500"/>')
  })
})

import { tableFormattingSignature } from '../src/renderer/editor/convert'

/** a body table cell saves through the table patch: a resized picture must count as a change */
describe('inline picture size in the table signature', () => {
  const table = (widthPx: number): PmNode => ({
    type: 'docTable',
    content: [
      {
        type: 'docRow',
        content: [
          {
            type: 'docCell',
            content: [
              {
                type: 'docParagraph',
                content: [
                  {
                    type: 'docInlineImage',
                    attrs: {
                      dataUrl: 'data:image/png;base64,AA==',
                      widthPx,
                      heightPx: 50,
                      xml: '<w:drawing/>',
                    },
                  },
                ],
              },
            ],
          },
        ],
      },
    ],
  })

  it('changes when the picture is resized', () => {
    expect(tableFormattingSignature(table(100))).not.toBe(tableFormattingSignature(table(135)))
    expect(tableFormattingSignature(table(100))).toBe(tableFormattingSignature(table(100)))
  })
})
