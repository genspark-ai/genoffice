import { describe, expect, it } from 'vitest'
import { DOMSerializer } from '@tiptap/pm/model'
import type { TableModel } from '@genoffice/docx-engine'
import { renderTableSpec } from '../src/renderer/editor/protected-render'

function renderTable(model: TableModel): HTMLElement {
  const { dom } = DOMSerializer.renderSpec(document, renderTableSpec(model) as never)
  return dom as HTMLElement
}

/** innerText equivalent (paragraph divs → one line each); jsdom does not implement innerText */
function cellText(td: Element): string {
  return Array.from(td.children)
    .filter((c) => c.tagName === 'DIV')
    .map((c) => c.textContent ?? '')
    .join('\n')
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

  it('renders one empty line box for cells whose richParas hold no text', () => {
    const td = renderTable({ rows: [[cell([''], [{ runs: [] }])]] }).querySelector('td')!
    const paras = td.querySelectorAll(':scope > div')
    expect(paras).toHaveLength(1)
    expect(paras[0].classList.contains('doc-p-empty')).toBe(true)
    expect(paras[0].querySelector('br')).not.toBeNull()
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

// jsdom's CSSOM drops min()/round() values on cssText assignment, so the
// paragraph line-box styles are asserted on the raw spec
type Spec = [string, Record<string, string>, ...unknown[]]
const isSpec = (x: unknown): x is Spec => Array.isArray(x) && typeof x[0] === 'string'
function findSpecs(spec: unknown, tag: string): Spec[] {
  if (!isSpec(spec)) return []
  const rest = spec.slice(2).flatMap((c) => findSpecs(c, tag))
  return spec[0] === tag ? [spec, ...rest] : rest
}
const paraStyles = (model: TableModel): string[] =>
  findSpecs(renderTableSpec(model), 'div').map((d) => d[1].style ?? '')

const SINGLE_LH =
  'line-height:round(up, calc(var(--doc-line-factor,1.2) * 1em), var(--doc-grid-pitch,0.0001px))'

describe('renderTableSpec paragraph line box', () => {
  it('gives each paragraph a run-size strut and its own single-spacing line height', () => {
    const [style] = paraStyles({
      rows: [
        [
          cell(
            ['tiny'],
            [
              {
                runs: [
                  { text: 'ti', sizeHalfPoints: 14 },
                  { text: 'ny', sizeHalfPoints: 12 },
                ],
                lineRule: 'auto',
                lineRawTwips: 240,
              },
            ],
          ),
        ],
      ],
    })
    expect(style).toContain('--doc-strut:7pt')
    expect(style).toContain('font-size:min(var(--doc-strut), 1em)')
    expect(style).toContain('--doc-line-factor:var(--doc-line-factor-latin,1.2)')
    // explicit w:line=240 (single): re-evaluated at the paragraph's own strut size
    expect(style).toContain(SINGLE_LH)
  })

  it('keeps the inherited strut when any run omits its size', () => {
    const [style] = paraStyles({
      rows: [[cell(['ab'], [{ runs: [{ text: 'a', sizeHalfPoints: 14 }, { text: 'b' }] }])]],
    })
    expect(style).not.toContain('--doc-strut')
    expect(style).toContain(SINGLE_LH)
  })

  it('restores exact line rule and explicit paragraph spacing', () => {
    const [style] = paraStyles({
      rows: [
        [
          cell(
            ['x'],
            [
              {
                runs: [{ text: 'x' }],
                lineRule: 'exact',
                lineRawTwips: 200,
                spaceBefore: 80,
                spaceAfter: 20,
              },
            ],
          ),
        ],
      ],
    })
    expect(style).toContain('line-height:10.0pt')
    expect(style).toContain('margin-top:round(down, 4.0pt,')
    expect(style).toContain('margin-bottom:round(down, 1.0pt,')
  })

  it('declared CJK run fonts drive the paragraph line factor', () => {
    const [style] = paraStyles({
      rows: [[cell(['宋体字'], [{ runs: [{ text: '宋体字', font: 'SimSun' }] }])]],
    })
    expect(style).toContain('--doc-line-factor:1.7')
  })

  it('plain-paras cells get script-based line factors per paragraph', () => {
    const styles = paraStyles({ rows: [[cell(['中文', 'latin'])]] })
    expect(styles).toHaveLength(2)
    expect(styles[0]).toContain('--doc-line-factor:var(--doc-line-factor-cjk,1.7)')
    expect(styles[1]).toContain('--doc-line-factor:var(--doc-line-factor-latin,1.2)')
    expect(styles[1]).toContain(SINGLE_LH)
  })
})
