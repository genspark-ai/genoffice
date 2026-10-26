import { describe, expect, it } from 'vitest'
import {
  copyMarkdownRich,
  markdownToHtml,
  markdownToPlainText,
} from '../src/markdown-copy'

describe('markdownToPlainText', () => {
  it('strips emphasis markers and keeps readable text', () => {
    expect(markdownToPlainText('Use **bold** and *ital* and `code` here.')).toBe(
      'Use bold and ital and code here.',
    )
  })

  it('turns a pipe table into tab-separated rows without the delimiter row', () => {
    const md = '| Name | Age |\n| --- | ---: |\n| Ann | 30 |\n| Bob | 41 |'
    expect(markdownToPlainText(md)).toBe('Name\tAge\nAnn\t30\nBob\t41')
  })

  it('keeps code content without the fences', () => {
    const md = '```ts\nconst a = 1\nconst b = 2\n```'
    expect(markdownToPlainText(md)).toBe('const a = 1\nconst b = 2')
  })

  it('numbers ordered lists and dashes unordered lists', () => {
    const md = '1. First\n2. Second\n\n- A\n- B'
    expect(markdownToPlainText(md)).toBe('1. First\n2. Second\n\n- A\n- B')
  })

  it('keeps link labels, not the hrefs', () => {
    expect(markdownToPlainText('See [the docs](https://x.test/a) now.')).toBe(
      'See the docs now.',
    )
  })

  it('keeps image alt text only', () => {
    expect(markdownToPlainText('![a red circle](img://1)\n\nAfter.')).toBe('a red circle\n\nAfter.')
  })
})

describe('markdownToHtml', () => {
  it('emits a real table for pipe tables', () => {
    const md = '| Name | Age |\n| --- | --- |\n| Ann | 30 |'
    const html = markdownToHtml(md)
    expect(html).toContain('<table>')
    expect(html).toContain('<th>Name</th>')
    expect(html).toContain('<td>Ann</td>')
    expect(html).not.toContain('| --- |')
  })

  it('emits pre/code for fences and strong/em/code inline', () => {
    expect(markdownToHtml('```\nlet x\n```')).toBe('<pre><code>let x</code></pre>')
    expect(markdownToHtml('**b** *i* `c`')).toBe(
      '<p><strong>b</strong> <em>i</em> <code>c</code></p>',
    )
  })

  it('escapes raw HTML in the source', () => {
    expect(markdownToHtml('a <b> & "c"')).toBe('<p>a &lt;b&gt; &amp; &quot;c&quot;</p>')
  })

  it('renders nav-free links as anchors with their href', () => {
    expect(markdownToHtml('[t](https://x.test)')).toBe('<p><a href="https://x.test">t</a></p>')
  })
})

describe('copyMarkdownRich', () => {
  it('falls back to plain text when ClipboardItem is unavailable', async () => {
    const written: string[] = []
    const nav = navigator as unknown as Record<string, unknown>
    const origWrite = nav.clipboard
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: { writeText: async (s: string) => written.push(s) },
    })
    try {
      await copyMarkdownRich('**hi**')
    } finally {
      Object.defineProperty(navigator, 'clipboard', { configurable: true, value: origWrite })
    }
    expect(written).toEqual(['hi'])
  })
})
