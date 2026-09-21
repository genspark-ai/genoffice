import { afterAll, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { escapeBlockStarts, fenceCodeSpan } from '../src/renderer/editor/markdownEscape'

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

/** serialize `source` and check the result reads back as the same document */
function roundTrip(source: string): string {
  const editor = createEditor()
  editor.commands.setContent(source, { contentType: 'markdown' })
  const doc = editor.state.doc.toJSON()
  const out = editor.getMarkdown()
  editor.commands.setContent(out, { contentType: 'markdown' })
  expect(editor.state.doc.toJSON(), out).toEqual(doc)
  return out
}

describe('nested code marks', () => {
  it.each([
    ['bold', 'Use a **`package.json`** for deps.'],
    ['italic', 'a *`b`* c'],
    ['bold+italic', 'a ***`d`*** c'],
    ['link', 'see [`e`](http://x) now'],
    ['strike', 'gone ~~`f`~~ now'],
    ['bold spanning code', 'a **x `y` z** b'],
    ['link around bold code', '[**`a`**](http://x)'],
  ])('renders code innermost: %s', (_name, source) => {
    expect(roundTrip(source)).toBe(source)
  })
})

describe('code spans with backticks', () => {
  it.each([
    ['plain', 'a', '`a`'],
    ['single backtick inside', 'a ` b', '``a ` b``'],
    ['triple backticks', '```', '```` ``` ````'],
    ['leading backtick', '`a', '`` `a ``'],
    ['trailing backtick', 'a`', '`` a` ``'],
    ['outer spaces', ' a ', '`  a  `'],
    ['single space', ' ', '` `'],
    ['leading space only', ' a', '` a`'],
  ])('fences %s', (_name, text, expected) => {
    expect(fenceCodeSpan(text)).toBe(expected)
  })

  it.each([
    'Triple backticks ```` ``` ```` here.',
    'Span ``a ` b`` x.',
    'Span `` `a` `` x and `  a  ` y and ` ` z.',
    'Tail `a` `` b` `` done.',
  ])('round-trips %s', (source) => {
    expect(roundTrip(source)).toBe(source)
  })

  it('keeps the spaces inside a code span', () => {
    // the parser strips one space on each side, so the document holds ` x `
    expect(roundTrip('a `  x  ` b')).toBe('a `  x  ` b')
  })
})

describe('block-start escapes in paragraphs', () => {
  it.each([
    ['\\## foo', '\\## foo'],
    ['\\# foo', '\\# foo'],
    ['\\- a', '\\- a'],
    ['\\+ b', '\\+ b'],
    ['\\* c', '\\* c'],
    ['2024\\. year', '2024\\. year'],
    ['1\\) first', '1\\) first'],
    ['\\> quoted', '\\> quoted'],
    ['\\---', '\\---'],
    ['foo  \n\\===', 'foo  \n\\==='],
    ['foo  \n\\---', 'foo  \n\\---'],
    ['foo  \n\\--', 'foo  \n\\--'],
    ['foo  \n\\-', 'foo  \n\\-'],
    ['foo  \n\\# bar', 'foo  \n\\# bar'],
    ['foo  \n\\- bar', 'foo  \n\\- bar'],
  ])('keeps %j a paragraph', (source, expected) => {
    expect(roundTrip(source)).toBe(expected)
  })

  it.each([
    '#hashtag',
    '1.5 x',
    'a - b',
    'x > y',
    '--- x',
    '===',
    '--',
    'a | b',
    '-> arrow',
    '2024.x',
  ])('leaves %j alone', (source) => {
    expect(roundTrip(source).replace(/&gt;/g, '>')).toBe(source)
  })

  it('edits inside a list item and heading stay put', () => {
    expect(roundTrip('- \\# not a heading\n- \\> not a quote').trimEnd()).toBe(
      '- \\# not a heading\n- \\> not a quote',
    )
    expect(roundTrip('# \\## foo').trimEnd()).toBe('# ## foo')
  })

  it('escapes a table delimiter row that would follow a hard break', () => {
    expect(escapeBlockStarts('a | b  \n--|--')).toBe('a | b  \n\\--|--')
    expect(escapeBlockStarts('| a |  \n|---|')).toBe('| a |  \n\\|---|')
    expect(escapeBlockStarts('|---|')).toBe('|---|')
  })

  it('does not touch table cells', () => {
    const source = '| a | b |\n| --- | --- |\n| # c | - d |'
    const editor = createEditor()
    editor.commands.setContent(source, { contentType: 'markdown' })
    expect(editor.getMarkdown()).toContain('| # c | - d |')
  })
})
