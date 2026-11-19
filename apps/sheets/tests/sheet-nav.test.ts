import { describe, expect, it } from 'vitest'
import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { Markdown } from '@genoffice/ui'

import { SHEET_NAV_SCHEME, parseSheetNavHref } from '../src/renderer/ai/sheet-nav'
import basePrompt from '../src/renderer/ai/prompts/base.md?raw'
import { WORKBOOK_TOOLS } from '../src/renderer/ai/tools'

describe('parseSheetNavHref', () => {
  it('reads a bare cell and a sheet-qualified range', () => {
    expect(parseSheetNavHref('sheetnav://B12')).toBe('B12')
    expect(parseSheetNavHref('sheetnav://Summary!B2:D9')).toBe('Summary!B2:D9')
  })

  it('decodes and quotes a sheet name the markdown link syntax forced the model to escape', () => {
    // Univer's range parser needs Excel quoting once the space is back
    expect(parseSheetNavHref('sheetnav://My%20Summary!B2')).toBe("'My Summary'!B2")
    expect(parseSheetNavHref("sheetnav://'My%20Summary'!B2")).toBe("'My Summary'!B2")
    expect(parseSheetNavHref('sheetnav://Summary!B2')).toBe('Summary!B2')
  })

  it('quotes parenthesised sheet names, escaped or not', () => {
    expect(parseSheetNavHref('sheetnav://Sheet1%20%282%29!B2')).toBe("'Sheet1 (2)'!B2")
    expect(parseSheetNavHref('sheetnav://Data(2)!B2:D9')).toBe("'Data(2)'!B2:D9")
  })

  it('keeps a malformed escape rather than dropping the citation', () => {
    expect(parseSheetNavHref('sheetnav://100%!B2')).toBe('100%!B2')
  })

  it('rejects other schemes and an empty reference', () => {
    expect(parseSheetNavHref('https://example.com')).toBeNull()
    expect(parseSheetNavHref('docnav://block/3')).toBeNull()
    expect(parseSheetNavHref(SHEET_NAV_SCHEME)).toBeNull()
    expect(parseSheetNavHref('sheetnav://   ')).toBeNull()
  })
})

describe('citation rendering', () => {
  it('renders sheetnav links as anchors and keeps external links literal', () => {
    const html = renderToStaticMarkup(
      createElement(Markdown, {
        text: 'The outlier is in [C42](sheetnav://C42), see [site](https://example.com).',
        nav: { scheme: SHEET_NAV_SCHEME, onNavigate: () => {} },
      }),
    )
    expect(html).toContain('class="ai-md-nav"')
    expect(html).toContain('href="sheetnav://C42"')
    expect(html).toContain('[site](https://example.com)')
  })

  it('keeps a citation clickable when the href carries balanced parens', () => {
    const html = renderToStaticMarkup(
      createElement(Markdown, {
        text: 'Totals live on [Sheet1 (2)!B2](sheetnav://Sheet1%20(2)!B2), see [x](sheetnav://Data%20%282%29!A1).',
        nav: { scheme: SHEET_NAV_SCHEME, onNavigate: () => {} },
      }),
    )
    expect(html).toContain('href="sheetnav://Sheet1%20(2)!B2"')
    expect(html).toContain('href="sheetnav://Data%20%282%29!A1"')
  })
})

describe('the prompt teaches citations over hijacking the selection', () => {
  it('gives the citation syntax in the base prompt', () => {
    expect(basePrompt).toContain('sheetnav://C42')
    expect(basePrompt).toContain('sheetnav://Summary!B2:D9')
  })

  it('reserves select_range for an explicit request to be moved', () => {
    const selectRange = WORKBOOK_TOOLS.find((tool) => tool.name === 'select_range')
    expect(selectRange?.description).toContain('sheetnav://C42')
    expect(selectRange?.description).toContain('only when they asked to be moved')
  })

  it('tells the model the selection is a send-time snapshot', () => {
    expect(basePrompt).toContain('captured at the moment the user hit send')
  })
})
