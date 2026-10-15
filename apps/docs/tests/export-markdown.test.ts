import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { parseHtmlFragment, replaceBlockRange } from '../src/renderer/ai/protocol'
import { documentToMarkdown, restrictedHtmlToMarkdown } from '../src/renderer/export-markdown'

describe('restrictedHtmlToMarkdown', () => {
  it('maps the restricted HTML subset to GFM', () => {
    const md = restrictedHtmlToMarkdown(
      '<h2>Title</h2>' +
        '<p>Plain <strong>bold </strong><em>it</em> <u>u</u> <s>gone</s> <a href="https://x.y/a b">link</a></p>' +
        '<ul><li>one</li><li>two<br>more</li></ul>' +
        '<ol><li>first</li><li>second</li></ol>' +
        '<pre>a `tick`\nb</pre>' +
        '<blockquote>quoted<br>twice</blockquote>' +
        '<table><tr><th>H1</th><th>H|2</th></tr><tr><td>c1</td><td>l1<br>l2</td></tr></table>' +
        '<formula>E = mc^2</formula>' +
        '<p>Inline <formula>x_1</formula> math</p>',
    )
    expect(md).toBe(
      [
        '## Title',
        '',
        'Plain **bold** *it* <u>u</u> ~~gone~~ [link](https://x.y/a%20b)',
        '',
        '- one',
        '- two\n  more',
        '',
        '1. first',
        '2. second',
        '',
        '```\na `tick`\nb\n```',
        '',
        '> quoted',
        '> twice',
        '',
        '| H1 | H\\|2 |',
        '| --- | --- |',
        '| c1 | l1<br>l2 |',
        '',
        '$$\nE = mc^2\n$$',
        '',
        'Inline $x_1$ math',
        '',
      ].join('\n'),
    )
  })

  it('escapes text that would otherwise become syntax', () => {
    expect(
      restrictedHtmlToMarkdown(
        '<p># not a heading</p><p>a*b_c [d] $5 and $10</p><p>1. not a list</p>' +
          '<p>line<br>- after break</p><blockquote>q<br># in quote</blockquote><ul><li>x<br>> y</li></ul>',
      ),
    ).toBe(
      [
        '\\# not a heading',
        'a\\*b\\_c \\[d\\] \\$5 and \\$10',
        '\\1. not a list',
        'line\n\\- after break',
        '> q\n> \\# in quote',
        '- x\n  \\> y',
      ].join('\n\n') + '\n',
    )
  })
})

describe('documentToMarkdown', () => {
  it('serializes a real editor document', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: '<p></p>',
    })
    try {
      const nodes = parseHtmlFragment('<h1>Doc</h1><p>Hello <strong>world</strong></p>', {
        bullet: null,
        ordered: null,
      })
      replaceBlockRange(editor, 0, editor.state.doc.childCount - 1, nodes)
      expect(documentToMarkdown(editor)).toEqual({
        markdown: '# Doc\n\nHello **world**\n',
        images: [],
      })
    } finally {
      editor.destroy()
    }
  })

  it('collects pictures and references them under the media folder token', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: '<p></p>',
    })
    const png = 'data:image/png;base64,iVBORw0KGgo='
    const lazy = 'genoffice-docx-media://abc/word/media/photo%201.jpeg'
    try {
      editor.commands.setContent({
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              { type: 'text', text: 'see ' },
              { type: 'docInlineImage', attrs: { dataUrl: png, widthPx: 10, heightPx: 10 } },
              { type: 'docInlineImage', attrs: { dataUrl: '', rule: { thicknessPx: 1 } } },
            ],
          },
          { type: 'docProtected', attrs: { blockType: 'image', imageDataUrl: lazy } },
          { type: 'docProtected', attrs: { blockType: 'image', imageDataUrl: png } },
          {
            type: 'docProtected',
            attrs: { blockType: 'image', imageDataUrl: 'data:text/plain;base64,QQ==' },
          },
          {
            type: 'docTable',
            content: [
              {
                type: 'docTableRow',
                content: [
                  {
                    type: 'docTableCell',
                    attrs: { rowspan: 2 },
                    content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'tall' }] }],
                  },
                  {
                    type: 'docTableCell',
                    content: [
                      {
                        type: 'docParagraph',
                        content: [
                          { type: 'text', text: 'cell ' },
                          { type: 'docInlineImage', attrs: { dataUrl: png } },
                        ],
                      },
                    ],
                  },
                ],
              },
              {
                type: 'docTableRow',
                content: [
                  {
                    type: 'docTableCell',
                    content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'b2' }] }],
                  },
                ],
              },
            ],
          },
        ],
      })
      expect(documentToMarkdown(editor)).toEqual({
        markdown:
          'see ![](__genoffice_media__/image1.png)\n\n' +
          '![](__genoffice_media__/image2.jpeg)\n\n' +
          '![](__genoffice_media__/image1.png)\n\n' +
          '| tall | cell ![](__genoffice_media__/image1.png) |\n| --- | --- |\n|  | b2 |\n',
        images: [
          { name: 'image1.png', src: png },
          { name: 'image2.jpeg', src: lazy },
        ],
      })
    } finally {
      editor.destroy()
    }
  })
})
