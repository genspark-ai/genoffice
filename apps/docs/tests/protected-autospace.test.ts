/**
 * Static protected DOM (read-only tables / textboxes) gets the same
 * .doc-autospace-pad spans the editor adds via decorations.
 */
import { describe, expect, it } from 'vitest'
import { DOMSerializer } from '@tiptap/pm/model'
import type { TableModel, TextboxDisplay } from '@genoffice/docx-engine'
import { renderTableSpec, renderTextboxSpec } from '../src/renderer/editor/protected-render'

const render = (spec: unknown): HTMLElement =>
  DOMSerializer.renderSpec(document, spec as never).dom as HTMLElement

const pads = (dom: HTMLElement) => dom.querySelectorAll('.doc-autospace-pad')

describe('renderTableSpec hangul spaces', () => {
  const H = '\ud55c\uae00'
  it('wraps spaces with a hangul neighbour across rich-cell runs, not Latin ones', () => {
    const model: TableModel = {
      rows: [
        [
          {
            paras: [`${H} A B`],
            richParas: [{ runs: [{ text: `${H} ` }, { text: 'A B', bold: true }] }],
          },
        ],
      ],
    }
    const dom = render(renderTableSpec(model))
    const wraps = dom.querySelectorAll('.doc-hangul-space')
    expect(wraps).toHaveLength(1)
    expect(dom.querySelector('td')!.textContent).toBe(`${H} A B`)
  })

  it('an inline picture between the runs breaks adjacency, an empty run does not', () => {
    const image = { dataUrl: 'data:image/png;base64,', xml: '', widthPx: 4, heightPx: 4 }
    const model: TableModel = {
      rows: [
        [
          {
            paras: [`${H} A ${H}B ${H}`],
            richParas: [
              {
                runs: [
                  { text: H },
                  { text: '', image },
                  { text: ' A' },
                  { text: '' },
                  { text: ` ${H}` },
                  { text: 'B' },
                  // the picture sits after this run's text: its trailing space is not beside the next hangul
                  { text: ' ', image },
                  { text: H },
                ],
              },
            ],
          },
        ],
      ],
    }
    const dom = render(renderTableSpec(model))
    const wraps = Array.from(dom.querySelectorAll('.doc-hangul-space'))
    expect(wraps).toHaveLength(1)
    expect(wraps[0].parentElement!.textContent).toBe(` ${H}`)
  })
})

describe('renderTableSpec autospace pads', () => {
  it('pads boundaries inside plain cell paragraphs without changing the text', () => {
    const dom = render(renderTableSpec({ rows: [[{ paras: ['ペン12'] }]] }))
    expect(pads(dom)).toHaveLength(1)
    expect(dom.querySelector('td')!.textContent).toBe('ペン12')
  })

  it('pads inside and between rich-cell runs', () => {
    const model: TableModel = {
      rows: [
        [
          {
            paras: ['ペン12', 'A漢'],
            richParas: [
              { runs: [{ text: 'ペン' }, { text: '12', bold: true }] },
              { runs: [{ text: 'A漢' }] },
            ],
          },
        ],
      ],
    }
    const dom = render(renderTableSpec(model))
    expect(pads(dom)).toHaveLength(2)
    expect(dom.querySelector('td')!.textContent).toBe('ペン12A漢')
  })

  it('adds none for space-separated seams', () => {
    const dom = render(renderTableSpec({ rows: [[{ paras: ['ペン 12'] }]] }))
    expect(pads(dom)).toHaveLength(0)
  })

  it('skips pads and disables native autospace when the paragraph turns autoSpace off', () => {
    const model: TableModel = {
      rows: [
        [
          {
            paras: ['ペン12A漢'],
            richParas: [{ runs: [{ text: 'ペン12' }, { text: 'A漢' }], autoSpace: false }],
          },
        ],
      ],
    }
    const dom = render(renderTableSpec(model))
    expect(pads(dom)).toHaveLength(0)
    const spans = dom.querySelectorAll('td span')
    expect(spans.length).toBe(2)
    spans.forEach((s) => {
      expect(s.getAttribute('style')).toMatch(/text-autospace:\s*no-autospace/)
    })
    expect(dom.querySelector('td')!.textContent).toBe('ペン12A漢')
  })

  it('gates pads per paragraph: only the autoSpace-off one loses them', () => {
    const model: TableModel = {
      rows: [
        [
          {
            paras: ['ペン12', 'ペン12'],
            richParas: [
              { runs: [{ text: 'ペン12' }] },
              { runs: [{ text: 'ペン12' }], autoSpace: false },
            ],
          },
        ],
      ],
    }
    const dom = render(renderTableSpec(model))
    expect(pads(dom)).toHaveLength(1)
  })
})

describe('renderTextboxSpec autospace pads', () => {
  it('pads across run boundaries and inside run text', () => {
    const box: TextboxDisplay = {
      paras: [{ runs: [{ text: 'テスト' }, { text: '17.0km', bold: true }] }],
    }
    const dom = render(renderTextboxSpec(box))
    expect(pads(dom)).toHaveLength(1)
    expect(dom.textContent).toBe('テスト17.0km')
  })

  it('autoSpace=false paragraph gets zero pads and native autospace off', () => {
    const box: TextboxDisplay = {
      paras: [{ runs: [{ text: 'テスト' }, { text: '17.0km', bold: true }], autoSpace: false }],
    }
    const dom = render(renderTextboxSpec(box))
    expect(pads(dom)).toHaveLength(0)
    dom.querySelectorAll('span').forEach((s) => {
      expect(s.getAttribute('style')).toMatch(/text-autospace:\s*no-autospace/)
    })
    expect(dom.textContent).toBe('テスト17.0km')
  })
})
