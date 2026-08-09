import { describe, expect, it } from 'vitest'
import { DOMSerializer } from '@tiptap/pm/model'
import type { TableModel } from '@genoffice/docx-engine'
import { renderTableSpec } from '../src/renderer/editor/protected-render'

function renderTable(model: TableModel): HTMLElement {
  const { dom } = DOMSerializer.renderSpec(document, renderTableSpec(model) as never)
  return dom as HTMLElement
}

/** innerText equivalent (<br> → \n); jsdom does not implement innerText */
function cellText(td: Element): string {
  let out = ''
  td.childNodes.forEach((n) => {
    out += n.nodeName === 'BR' ? '\n' : (n.textContent ?? '')
  })
  return out
}

const cell = (paras: string[], richParas?: TableModel['rows'][0][0]['richParas']) => ({
  paras,
  ...(richParas ? { richParas } : {}),
})

describe('renderTableSpec rich cell content', () => {
  it('renders run-level sz/color/bold from richParas', () => {
    const model: TableModel = {
      rows: [
        [
          cell(
            ['Big red', 'plain'],
            [
              {
                runs: [
                  { text: 'Big ', sizeHalfPoints: 15, color: 'FF0000', bold: true },
                  { text: 'red', sizeHalfPoints: 15, italic: true },
                ],
              },
              { runs: [{ text: 'plain' }] },
            ],
          ),
        ],
      ],
    }
    const td = renderTable(model).querySelector('td')!
    const spans = td.querySelectorAll('span')
    expect(spans).toHaveLength(3)
    // renderSpec assigns style via cssText, so jsdom normalizes (spaces, rgb());
    // assert on the attribute string, not CSSOM property reads
    const first = spans[0].getAttribute('style')!
    expect(first).toMatch(/font-size:\s*7\.5pt/)
    expect(first).toMatch(/color:\s*(#FF0000|rgb\(255,\s*0,\s*0\))/i)
    expect(first).toMatch(/font-weight:\s*700/)
    const second = spans[1].getAttribute('style')!
    expect(second).toMatch(/font-size:\s*7\.5pt/)
    expect(second).toMatch(/font-style:\s*italic/)
    expect(spans[2].getAttribute('style')).toBeNull()
    expect(cellText(td)).toBe('Big red\nplain')
  })

  it('keeps innerText line semantics identical to the paras fallback', () => {
    const paras = ['a', '', 'b']
    const rich = renderTable({
      rows: [[cell(paras, [{ runs: [{ text: 'a' }] }, { runs: [] }, { runs: [{ text: 'b' }] }])]],
    }).querySelector('td')!
    const plain = renderTable({ rows: [[cell(paras)]] }).querySelector('td')!
    expect(cellText(rich)).toBe('a\n\nb')
    expect(cellText(rich)).toBe(cellText(plain))
    expect(rich.querySelectorAll('br')).toHaveLength(plain.querySelectorAll('br').length)
  })

  it('falls back to plain paras when richParas is absent', () => {
    const td = renderTable({ rows: [[cell(['x', 'y'])]] }).querySelector('td')!
    expect(td.querySelectorAll('span')).toHaveLength(0)
    expect(cellText(td)).toBe('x\ny')
  })

  it('keeps the nbsp placeholder for cells whose richParas hold no text', () => {
    const td = renderTable({ rows: [[cell([''], [{ runs: [] }])]] }).querySelector('td')!
    expect(td.textContent).toBe('\u00a0')
  })

  it('leads with the cs font chain for complex-script run text', () => {
    const fonts = { font: 'Calibri', fontAscii: 'Calibri', csFont: 'Traditional Arabic' }
    const model: TableModel = {
      rows: [
        [
          cell(
            ['مرحبا Hi'],
            [
              {
                runs: [
                  { text: 'مرحبا', ...fonts },
                  { text: 'Hi', ...fonts },
                ],
              },
            ],
          ),
        ],
      ],
    }
    const spans = renderTable(model).querySelectorAll('td span')
    const arabic = spans[0].getAttribute('style')!
    expect(arabic).toMatch(/font-family:\s*['"]Traditional Arabic['"]/)
    expect(arabic).toContain('Calibri')
    // csFont present but text is Latin-only: keep the plain Latin chain
    const latin = spans[1].getAttribute('style')!
    expect(latin).toMatch(/font-family:\s*['"]Calibri['"]/)
    expect(latin).not.toContain('Traditional Arabic')
  })

  it('cell-level color/bold stay as td-level fallback', () => {
    const model: TableModel = {
      rows: [
        [
          {
            ...cell(['t'], [{ runs: [{ text: 't' }] }]),
            color: '112233',
            bold: true,
          },
        ],
      ],
    }
    const td = renderTable(model).querySelector('td')!
    const tdStyle = td.getAttribute('style')!
    expect(tdStyle).toMatch(/color:\s*(#112233|rgb\(17,\s*34,\s*51\))/i)
    expect(tdStyle).toMatch(/font-weight:\s*600/)
    expect(td.querySelector('span')!.getAttribute('style')).toBeNull()
  })
})
