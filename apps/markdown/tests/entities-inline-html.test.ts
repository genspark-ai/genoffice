import { afterAll, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { decodeEntities } from '../src/renderer/editor/inlineTokens'
import { buildSourceMap, spliceMarkdown } from '../src/renderer/markdown/sourceSplice'

const editors: Editor[] = []
afterAll(() => {
  for (const e of editors) e.destroy()
})

function createEditor(): Editor {
  const editor = new Editor({
    extensions: buildExtensions({
      slashController: {
        onOpen: () => {},
        onUpdate: () => {},
        onKeyDown: () => false,
        onClose: () => {},
      },
      slashItems: () => [],
    }),
    content: '',
  })
  editors.push(editor)
  return editor
}

function html(source: string): string {
  const editor = createEditor()
  editor.commands.setContent(source, { contentType: 'markdown' })
  return editor.getHTML()
}

function endOfBlock(editor: Editor, index: number): number {
  let pos = 0
  for (let i = 0; i < index; i++) pos += editor.state.doc.child(i).nodeSize
  return pos + editor.state.doc.child(index).nodeSize - 1
}

/** load `source`, run `edit`, and save through the block splicer */
function editAndSave(
  source: string,
  edit: (editor: Editor) => void,
): { editor: Editor; saved: string } {
  const editor = createEditor()
  editor.commands.setContent(source, { contentType: 'markdown' })
  const map = buildSourceMap(editor, editor.state.doc, source)
  expect(map).not.toBeNull()
  edit(editor)
  return { editor, saved: spliceMarkdown(editor, editor.state.doc, map!) }
}

/** top-level blocks, without the caret paragraph the editor keeps after a trailing table */
function docJson(editor: Editor): unknown[] {
  const blocks: unknown[] = []
  editor.state.doc.forEach((node) => blocks.push(node.toJSON()))
  const last = blocks[blocks.length - 1] as { type: string; content?: unknown }
  return last?.type === 'paragraph' && !last.content ? blocks.slice(0, -1) : blocks
}

function parsedJson(editor: Editor, markdown: string): unknown[] {
  const doc = editor.schema.nodeFromJSON(editor.markdown!.parse(markdown))
  const blocks: unknown[] = []
  doc.forEach((node) => blocks.push(node.toJSON()))
  return blocks
}

const BR_PARAGRAPH = [
  '[**Visit the website.**](https://markdown-here.com)<br>',
  '[**Get it for Chrome.**](https://chrome.example/webstore)<br>',
  '[**Get it for Firefox.**](https://addons.example/firefox)<br>',
].join('\n')

// public-hygiene: fixture
const VUE_SPAN =
  '<span class="options-api">`this.$emit`</span><span class="composition-api">`emit`</span> の呼び出し'

describe('HTML entities', () => {
  it('decodes named and numeric entities in text', () => {
    expect(
      html('an&nbsp;index.md&nbsp;file &copy; R&amp;D &#123;x&#125; &#x1F600; &mdash; &#160;'),
    ).toBe('<p>an&nbsp;index.md&nbsp;file \u00a9 R&amp;D {x} \u{1F600} \u2014 &nbsp;</p>')
  })

  it('leaves unknown entities, double-encoded entities and escapes alone', () => {
    expect(html('&foo; &amp;lt; \\&amp; &#0;')).toBe('<p>&amp;foo; &amp;lt; &amp;amp; �</p>')
  })

  it('does not decode inside code spans or code blocks', () => {
    expect(html('`&nbsp; &amp;` and <code>&amp;lt;</code>')).toBe(
      '<p><code>&amp;nbsp; &amp;amp;</code> and <code>&amp;lt;</code></p>',
    )
    expect(html('```\n&nbsp;\n```')).toContain('<pre><code>&amp;nbsp;</code></pre>')
  })

  it('decodes in nested contexts', () => {
    expect(
      html('- a &copy;\n\n> q &nbsp;\n\n# H &#124;\n\n| c &copy; |\n|---|\n| &nbsp; |'),
    ).toContain('<li><p>a ©</p></li>')
    expect(decodeEntities('&lt;&gt;&quot;&apos;&#39;')).toBe("<>\"''")
  })

  it('writes a decoded no-break space back as &nbsp;', () => {
    const source = '| a |\n|---|\n| &nbsp;&nbsp;`x` |\n\nFront&nbsp;matter\n'
    const { editor, saved } = editAndSave(source, (e) => {
      e.commands.insertContentAt(endOfBlock(e, 0) - 3, 'w')
    })
    expect(saved).toContain('| &nbsp;&nbsp;`xw` |')
    expect(parsedJson(editor, saved)).toEqual(docJson(editor))
  })

  it('escapes a decoded pipe when an edited table cell is re-serialized', () => {
    const source = '| a | b |\n|---|---|\n| x &#124; y | z |'
    const { editor, saved } = editAndSave(source, (e) => {
      e.commands.insertContentAt(endOfBlock(e, 0) - 4, 'w')
    })
    expect(saved).toContain('| x \\| y |')
    expect(parsedJson(editor, saved)).toEqual(docJson(editor))
  })
})

describe('inline HTML tags', () => {
  it('drops inert tags and parses the markdown inside them', () => {
    const out = html(VUE_SPAN)
    expect(out).not.toContain('`')
    expect(out).toBe('<p><code>this.$emitemit</code> の呼び出し</p>')
  })

  it('keeps the content of tags without a schema mapping', () => {
    expect(
      html(
        'A <font color="red">**bold**</font> <center>*c*</center> <abbr title="t">ab</abbr> <small>sm</small> <big>bg</big> <mark>mk</mark> <kbd>**k**</kbd> <sub>s</sub> <sup>p</sup> <div>*dv*</div> <span/> <!-- c -->',
      ),
    ).toBe(
      '<p>A <strong>bold</strong> <em>c</em> ab sm bg mk <strong>k</strong> s p <em>dv</em>  </p>',
    )
  })

  it('leaves names marked accepts but CommonMark does not as literal text', () => {
    expect(html('{{<glossary_tooltip text="c" term_id="c">}} v<MERMAID_RELEASE_VERSION>+')).toBe(
      '<p>{{&lt;glossary_tooltip text="c" term_id="c"&gt;}} v&lt;MERMAID_RELEASE_VERSION&gt;+</p>',
    )
  })

  it('still maps formatting tags to marks', () => {
    expect(
      html(
        '<b>x</b> <strong>y</strong> <i>i</i> <em>e</em> <code>*c*</code> <del>d</del> <s>ss</s> <a href="https://x.y">l</a>',
      ),
    ).toBe(
      '<p><strong>x</strong> <strong>y</strong> <em>i</em> <em>e</em> <code>*c*</code> <s>d</s> <s>ss</s> <a target="_blank" rel="noopener noreferrer nofollow" href="https://x.y">l</a></p>',
    )
  })

  it('collapses the newline after <br>', () => {
    expect(html('a<br>\nb')).toBe('<p>a<br>b</p>')
    expect(html('**a<br>\nb**')).toBe('<p><strong>a</strong><br><strong>b</strong></p>')
  })
})

describe('saving', () => {
  const DOCUMENT = [
    'an&nbsp;index.md&nbsp;file &copy; &#124;',
    '',
    VUE_SPAN,
    '',
    BR_PARAGRAPH,
    '',
    '| a &#124; |',
    '|---|',
    '| <kbd>k</kbd> |',
    '',
    'end',
    '',
  ].join('\n')

  it('keeps an untouched document byte-identical', () => {
    const { saved } = editAndSave(DOCUMENT, () => {})
    expect(saved).toBe(DOCUMENT)
  })

  it('keeps an edited <br> paragraph as one paragraph, breaks written as <br>', () => {
    const source = `${BR_PARAGRAPH}\n\nnext\n`
    const { editor, saved } = editAndSave(source, (e) => {
      e.commands.insertContentAt(endOfBlock(e, 0), ' EDIT')
    })
    expect(saved).toBe(
      [
        '[**Visit the website.**](https://markdown-here.com)<br>',
        '[**Get it for Chrome.**](https://chrome.example/webstore)<br>',
        '[**Get it for Firefox.**](https://addons.example/firefox)<br>',
        ' EDIT',
        '',
        'next',
        '',
      ].join('\n'),
    )
    expect(parsedJson(editor, saved)).toEqual(docJson(editor))
  })

  it('writes a hard break inside a table cell as a single <br>', () => {
    const source = '| a |\n|---|\n| x<br>y |\n\nz<br>\nw\n'
    const { editor, saved } = editAndSave(source, (e) => {
      e.commands.insertContentAt(endOfBlock(e, 0) - 3, 'EDIT')
    })
    expect(saved).toContain('| x<br>yEDIT |')
    expect(parsedJson(editor, saved)).toEqual(docJson(editor))
  })

  it('follows the document when hard breaks are two spaces', () => {
    const { saved } = editAndSave('a  \nb\n\nnext\n', (e) => {
      e.commands.insertContentAt(endOfBlock(e, 0), ' EDIT')
    })
    expect(saved).toBe('a  \nb EDIT\n\nnext\n')
  })
})
