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

describe('AI bubble markdown: images', () => {
  const withImages = (text: string, resolve?: (href: string) => string | undefined): string =>
    renderToStaticMarkup(createElement(Markdown, { text, images: { resolve } }))

  it('renders a standalone image line through the resolver', () => {
    const html = withImages('![the ribbon](help:ribbon)', (h) => `blob:${h}`)
    expect(html).toContain('<img class="ai-md-img" src="blob:help:ribbon" alt="the ribbon"')
  })

  it('renders an image sharing a line with prose', () => {
    // the case that was dropped: parseBlocks only ever made a *whole* line an
    // image block, so anything with text around it fell through to the inline
    // pass, which had no image case and printed the source
    const html = withImages(
      'Click the tab, then ![the ribbon](help:ribbon) appears.',
      (h) => `blob:${h}`,
    )
    expect(html).toContain('src="blob:help:ribbon"')
    expect(html).toContain('Click the tab, then')
    expect(html).toContain('appears.')
  })

  it('renders an image inside a list step', () => {
    const html = withImages('- open ![settings](help:s)\- open the dialog', (h) => `blob:${h}`)
    expect(html).toContain('src="blob:help:s')
  })

  it('renders an image in a table cell', () => {
    const html = withImages(
      ['| UI | Where |', '|:--|--:|', '| ![a](help:a) | x |'].join('\n'),
      (h) => `blob:${h}`,
    )
    expect(html).toContain('src="blob:help:a')
  })

  it('leaves an inline image as literal text when no resolver is passed', () => {
    // every chat panel passes no resolver, and this is what they rendered
    // before: the source text, untouched
    const html = render('Click then ![a](help:a) appears.')
    expect(html).not.toContain('<img')
    expect(html).toContain('![a](help:a)')
  })

  it('leaves an inline image as literal text when the resolver declines the href', () => {
    const html = withImages('Click then ![a](help:a) appears.', () => undefined)
    expect(html).not.toContain('<img')
    expect(html).toContain('![a](help:a)')
  })

  it('does not leave a stray bang in front of an image it rendered', () => {
    // ![a](b) also matches the link branch from its '['; the image
    // alternative has to win, or the output reads "!<img>"
    const html = withImages('see ![a](help:z) here', (h) => `blob:${h}`)
    expect(html).not.toContain('!<img')
    expect(html).not.toContain('![a]')
  })

  it('keeps links rendering as before', () => {
    const html = withImages('see [docs](help:d) now', (h) => `blob:${h}`)
    expect(html).toContain('[docs](help:d)')
  })

  it('drops a standalone image line nothing can resolve', () => {
    // the other path, and it is older than the inline one: a lone figure line
    // with no resolver behind it is noise, not prose to show the reader
    expect(render('![only](help:x)')).not.toContain('![only]')
  })
})
