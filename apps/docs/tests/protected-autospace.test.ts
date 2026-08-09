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
