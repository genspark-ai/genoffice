import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { describe, expect, it } from 'vitest'
import { Markdown } from '@genoffice/ui'

const render = (text: string): string => renderToStaticMarkup(createElement(Markdown, { text }))

describe('AI bubble markdown: tables', () => {
  it('renders a pipe table with a header row and aligned columns', () => {
    const html = render(
      [
        '| Region | Q1 | Q2 |',
        '|:--|--:|:-:|',
        '| North | 120 | 130 |',
        '| South | 80 | 95 |',
      ].join('\n'),
    )
    expect(html).toContain('<table class="ai-md-table">')
    expect(html.match(/<th[ >]/g)).toHaveLength(3)
    expect(html.match(/<tr>/g)).toHaveLength(3)
    expect(html).toContain('<th style="text-align:left">Region</th>')
    expect(html).toContain('<th style="text-align:right">Q1</th>')
    expect(html).toContain('<td style="text-align:center">130</td>')
    expect(html).not.toContain('|')
  })

  it('keeps inline formatting inside cells and tolerates ragged rows', () => {
    const html = render(
      ['A | B', '--- | ---', '**x** | `a|b`', 'only-one |', 'p | q | extra'].join('\n'),
    )
    expect(html).toContain('<td><strong>x</strong></td>')
    expect(html).toContain('<td><code>a|b</code></td>')
    expect(html).toContain('<td>only-one</td><td></td>')
    expect(html).not.toContain('extra')
  })

  it('unescapes \\| inside a cell', () => {
    const html = render(['a | b', '- | -', 'x \\| y | z'].join('\n'))
    expect(html).toContain('<td>x | y</td>')
  })

  it('a header line without a delimiter row (still streaming) stays a paragraph', () => {
    const html = render('| Region | Q1 |')
    expect(html).not.toContain('<table')
    expect(html).toContain('<p>| Region | Q1 |</p>')
  })

  it('ends the table at a blank line or a line without pipes', () => {
    const html = render(
      ['a | b', '- | -', '1 | 2', 'Summary follows.', '', 'Next paragraph.'].join('\n'),
    )
    expect(html.match(/<tr>/g)).toHaveLength(2)
    expect(html).toContain('<p>Summary follows.</p>')
    expect(html).toContain('<p>Next paragraph.</p>')
  })

  it('does not treat a lone pipe in prose as a table', () => {
    const html = render('either A | B works\nand this line too')
    expect(html).not.toContain('<table')
  })
})

describe('AI bubble markdown: fenced code', () => {
  it('renders a fenced block verbatim without inline parsing', () => {
    const html = render(
      ['Use:', '```ts', 'const x = **not bold**', '  indented | pipe', '```', 'Done.'].join('\n'),
    )
    expect(html).toContain(
      '<pre class="ai-md-pre"><code>const x = **not bold**\n  indented | pipe</code></pre>',
    )
    expect(html).not.toContain('<strong>')
    expect(html).toContain('<p>Done.</p>')
  })

  it('an unterminated fence (still streaming) renders as code once, not twice', () => {
    const html = render('```\nline one\nline two')
    expect(html.match(/<pre/g)).toHaveLength(1)
    expect(html).toContain('line one\nline two')
  })
})
