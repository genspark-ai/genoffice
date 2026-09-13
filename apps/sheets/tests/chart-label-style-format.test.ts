import { describe, expect, it } from 'vitest'
import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'

import { BarChart } from '../src/renderer/WorkbookVisuals'

// A Numbers export: source-linked currency labels (the sidecar resolved the
// cells' format) drawn white inside a stacked column, per c:dLbls/c:txPr.
const seriesList = [
  { name: 'Loan A', categories: ['Interest', 'Fees'], values: [3968, 120], numberFormat: '#,##0' },
  { name: 'Loan B', categories: ['Interest', 'Fees'], values: [2510, 90], numberFormat: '#,##0' },
]
const labels = (markup: string): string[] =>
  [...markup.matchAll(/<text [^>]*class="data-label"[^>]*>([^<]*)</g)].map(
    (match) => match[1] ?? '',
  )
const labelTags = (markup: string): string[] =>
  [...markup.matchAll(/<text [^>]*class="data-label"[^>]*>/g)].map((match) => match[0])

describe('bar data labels: source-linked format and txPr font', () => {
  const props = {
    seriesList,
    isHorizontal: false,
    grouping: 'stacked' as const,
    dataLabels: 'value' as const,
    dataLabelPosition: 'inside-end' as const,
    dataLabelFormat: '[$$-409]#,##0',
    valueAxis: { hidden: true },
  }

  it('formats label values through the full format code', () => {
    const texts = labels(renderToStaticMarkup(createElement(BarChart, props)))
    expect(texts).toEqual(expect.arrayContaining(['$3,968', '$2,510', '$120', '$90']))
  })

  it('applies the dLbls txPr color, size and weight as inline style', () => {
    const styled = labelTags(
      renderToStaticMarkup(
        createElement(BarChart, {
          ...props,
          dataLabelStyle: { color: '#FFFFFF', size: 12, bold: true },
        }),
      ),
    ).filter((tag) => tag.includes('style='))
    expect(styled).toHaveLength(4)
    for (const tag of styled) {
      expect(tag).toContain('fill:#FFFFFF')
      expect(tag).toContain('font-size:calc(12 * var(--chart-pt, 1px))')
      expect(tag).toContain('font-weight:700')
    }
  })

  it('leaves labels unstyled without a txPr', () => {
    const tags = labelTags(renderToStaticMarkup(createElement(BarChart, props)))
    expect(tags.some((tag) => tag.includes('style='))).toBe(false)
  })
})
