import { afterAll, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'
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

/** open `source` and save it untouched through the block splice */
function openAndSave(editor: Editor, source: string): string {
  editor.commands.setContent(source, { contentType: 'markdown' })
  const map = buildSourceMap(editor, editor.state.doc, source)
  expect(map).not.toBeNull()
  return spliceMarkdown(editor, editor.state.doc, map!)
}

function topLevel(editor: Editor): string[] {
  const out: string[] = []
  editor.state.doc.forEach((node) => out.push(node.type.name))
  return out
}

describe('tables', () => {
  it('a header/delimiter cell-count mismatch is a paragraph, not shredded text', () => {
    const editor = createEditor()
    const source = '| abc | def |\n| --- |\n| bar |\n'
    const saved = openAndSave(editor, source)
    expect(editor.getHTML()).toBe('<p>| abc | def |\n| --- |\n| bar |</p>')
    expect(saved).toBe(source)
  })

  it('more header cells than delimiter cells, inside a document', () => {
    const editor = createEditor()
    const source = 'intro\n\n| a | b | c |\n|---|---|\n| 1 | 2 |\n\nafter\n'
    expect(openAndSave(editor, source)).toBe(source)
    expect(topLevel(editor).filter((t) => t !== 'paragraph')).toEqual([])
    expect(editor.getText()).not.toMatch(/\n\|\n/)
  })

  it('body rows may have a different cell count', () => {
    const editor = createEditor()
    const source = '| a | b |\n|---|---|\n| 1 |\n| 3 | 4 | 5 |\n'
    openAndSave(editor, source)
    expect(topLevel(editor)[0]).toBe('table')
  })

  it('escaped pipes do not count as cell separators', () => {
    const editor = createEditor()
    openAndSave(editor, '| a \\| b | c |\n|---|---|\n| 1 | 2 |\n')
    expect(topLevel(editor)[0]).toBe('table')
    expect(editor.getHTML()).toContain('<th colspan="1" rowspan="1"><p>a | b</p></th>')
  })

  it('a dash row whose cells are not delimiters is not a table', () => {
    const editor = createEditor()
    const source = '| a | b |\n| -x- | --- |\n'
    expect(openAndSave(editor, source)).toBe(source)
    expect(topLevel(editor)[0]).toBe('paragraph')
  })
})

describe('indented fences', () => {
  it.each([
    [' ```\n aaa\naaa\n```\n', 'aaa\naaa'],
    ['  ```\naaa\n  aaa\naaa\n  ```\n', 'aaa\naaa\naaa'],
    ['   ```\n   aaa\n    aaa\n   aaa\n   ```\n', 'aaa\n aaa\naaa'],
  ])('%j parses as a code block, de-indented by the fence indent', (source, text) => {
    const editor = createEditor()
    const saved = openAndSave(editor, source)
    expect(topLevel(editor)).toEqual(['codeBlock', 'paragraph'])
    expect(editor.state.doc.child(0).textContent).toBe(text)
    expect(saved).toBe(source)
  })

  it('four spaces is indented code, not a fence', () => {
    const editor = createEditor()
    openAndSave(editor, '    ```\n    aaa\n    ```\n')
    expect(topLevel(editor)[0]).toBe('codeBlock')
    expect(editor.state.doc.child(0).textContent).toBe('```\naaa\n```')
  })

  it('an edited indented fence is re-serialized flush left', () => {
    const editor = createEditor()
    const source = 'intro\n\n ```js\n x\n ```\n'
    editor.commands.setContent(source, { contentType: 'markdown' })
    const map = buildSourceMap(editor, editor.state.doc, source)!
    editor.commands.insertContentAt(editor.state.doc.child(0).nodeSize + 1, 'y = ')
    expect(spliceMarkdown(editor, editor.state.doc, map)).toBe('intro\n\n```js\ny = x\n```\n')
  })
})

describe('review follow-ups', () => {
  it('a fence inside a tight ordered list does not make it loose', () => {
    const editor = createEditor()
    const src = '1. one\n   ```\n   code\n   ```\n2. two\n'
    editor.commands.setContent(src, { contentType: 'markdown' })
    expect(editor.state.doc.firstChild?.attrs.loose).toBe(false)
  })

  it('an escaped backslash before a pipe still separates cells', () => {
    const editor = createEditor()
    editor.commands.setContent('| a\\\\ | b |\n| --- | --- |\n| c | d |\n', {
      contentType: 'markdown',
    })
    expect(editor.getHTML()).toContain('<table')
  })
})
