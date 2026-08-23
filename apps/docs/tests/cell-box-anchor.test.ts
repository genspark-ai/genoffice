/**
 * Cell-anchored boxes: positionV relativeFrom="paragraph" resolves from the
 * anchor paragraph, so the display strut must sit before that paragraph, and
 * exact line spacing inside a fixed-height box must not inherit the page's
 * line height (the single line would fall outside the clip window).
 */
import { describe, expect, it } from 'vitest'
import type { TableModel, TextboxDisplay } from '@genoffice/docx-engine'
import { pmTableToModel, tableModelToPmNode } from '../src/renderer/editor/convert'
import { renderTextboxSpec } from '../src/renderer/editor/protected-render'

const box = (text: string): TextboxDisplay => ({
  paras: [{ runs: [{ text }] }],
  widthPx: 100,
  heightPx: 16,
  offsetXEmu: 0,
  offsetYEmu: 9525,
})

const boxTexts = (node: { attrs?: Record<string, unknown> }): string[] =>
  (node.attrs?.boxes as TextboxDisplay[]).map((b) => b.paras[0].runs[0].text)

describe('cell-anchored box placement', () => {
  const model: TableModel = {
    rows: [
      [
        {
          paras: ['first', 'second', 'third'],
          anchoredBoxes: [box('a'), box('b'), box('c')],
          anchoredBoxAnchors: [1, 1, 2],
        },
      ],
    ],
  }

  it('splices box struts before their anchor paragraphs, grouped per anchor', () => {
    const cell = tableModelToPmNode(model).content![0].content![0]
    expect(cell.content!.map((n) => n.type)).toEqual([
      'docParagraph',
      'docCellBoxes',
      'docParagraph',
      'docCellBoxes',
      'docParagraph',
    ])
    expect(boxTexts(cell.content![1])).toEqual(['a', 'b'])
    expect(boxTexts(cell.content![3])).toEqual(['c'])
  })

  it('defaults to the cell top when anchor indices are absent', () => {
    const legacy: TableModel = { rows: [[{ paras: ['only'], anchoredBoxes: [box('a')] }]] }
    const cell = tableModelToPmNode(legacy).content![0].content![0]
    expect(cell.content!.map((n) => n.type)).toEqual(['docCellBoxes', 'docParagraph'])
  })

  it('round-trips anchors through pmTableToModel', () => {
    const cell = pmTableToModel(tableModelToPmNode(model)).rows[0][0]
    expect(cell.anchoredBoxes?.map((b) => b.paras[0].runs[0].text)).toEqual(['a', 'b', 'c'])
    expect(cell.anchoredBoxAnchors).toEqual([1, 1, 2])
  })
})

describe('renderTextboxSpec line rules', () => {
  // read the raw spec style (jsdom's CSS parser re-serializes and drops max())
  const paraStyle = (spec: unknown): string => {
    const para = (spec as unknown[])[2] as [string, Record<string, string>]
    return para[1].style ?? ''
  }

  it('exact fixes the line box at the declared height', () => {
    const style = paraStyle(
      renderTextboxSpec({
        heightPx: 16,
        paras: [{ runs: [{ text: 'x' }], lineRule: 'exact', lineRawTwips: 240 }],
      }),
    )
    expect(style).toContain('line-height:12.0pt')
  })

  it('atLeast only floors the line box (grid-snapped single as the floor)', () => {
    const style = paraStyle(
      renderTextboxSpec({
        paras: [{ runs: [{ text: 'x' }], lineRule: 'atLeast', lineRawTwips: 400 }],
      }),
    )
    expect(style).toContain(
      'line-height:max(20.0pt, var(--doc-line-grid, calc(var(--doc-line-factor,1.2) * 1em)))',
    )
  })

  it('auto single spacing resolves the shared grid/factor base', () => {
    const style = paraStyle(
      renderTextboxSpec({ paras: [{ runs: [{ text: 'x' }], lineSpacing: 1 }] }),
    )
    expect(style).toContain('line-height:var(--doc-line-grid,var(--doc-line-factor,1.2))')
  })

  it('snapToGrid=0 opts the paragraph out of docGrid snapping', () => {
    const spec = renderTextboxSpec({
      paras: [{ runs: [{ text: 'x' }], snapToGrid: false }],
    }) as unknown[]
    const para = spec[2] as [string, Record<string, string>]
    expect(para[1].class).toContain('doc-nosnap')
    expect(para[1].style).toContain('--doc-grid-pitch:0.0001px')
  })

  it('CJK run fonts drive the paragraph line factor', () => {
    const style = paraStyle(
      renderTextboxSpec({
        paras: [{ runs: [{ text: '患者', font: 'ＭＳ Ｐゴシック', sizeHalfPoints: 24 }] }],
      }),
    )
    expect(style).toContain('--doc-line-factor:1.3029')
    expect(style).toContain('--doc-strut:12pt')
  })
})
