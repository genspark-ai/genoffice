import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { OP_META, runOps, uiOp, validateOps, type OpMeta } from '../src/renderer/editor/ops'

const editors: Editor[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
})

function createEditor(md = ''): Editor {
  const editor = new Editor({
    extensions: buildExtensions({
      slashController: { onOpen() {}, onUpdate() {}, onKeyDown: () => false, onClose() {} },
      slashItems: () => [],
    }),
    content: '',
  })
  if (md) editor.commands.setContent(md, { contentType: 'markdown' })
  editors.push(editor)
  return editor
}

/** caret right after the first occurrence of `needle` */
function placeAfter(editor: Editor, needle: string): void {
  let found = -1
  editor.state.doc.descendants((node, pos) => {
    if (found >= 0) return false
    if (node.isText && node.text?.includes(needle))
      found = pos + node.text.indexOf(needle) + needle.length
    return found < 0
  })
  expect(found).toBeGreaterThan(0)
  editor.commands.setTextSelection(found)
}

function blocks(editor: Editor): string[] {
  const out: string[] = []
  editor.state.doc.forEach((n) => out.push(`${n.type.name}:${n.textContent}`))
  return out
}

const ai = (editor: Editor, ...ops: Parameters<typeof runOps>[1]) =>
  runOps(editor, ops, { source: 'ai' })

describe('validateOps', () => {
  it('accepts the documented shapes and rejects the rest', () => {
    expect('ops' in validateOps([{ op: 'deleteBlocks', target: 'selection' }])).toBe(true)
    expect('ops' in validateOps([{ op: 'setLink', target: { start: 0 }, href: null }])).toBe(true)
    expect(validateOps([])).toEqual({ error: 'ops must be a non-empty array' })
    expect(validateOps([{ op: 'setBlockType', target: 'selection', type: 'heading' }])).toEqual({
      error: 'ops[0] setBlockType: heading needs level 1-6',
    })
    expect(validateOps([{ op: 'toggleList', target: 'selection', list: 'nope' }])).toMatchObject({
      error: expect.stringContaining('list must be one of'),
    })
    expect(validateOps([{ op: 'insertContent', after: 'end', markdown: 'x' }])).toMatchObject({
      error: expect.stringContaining('after must be'),
    })
  })
})

describe('structural ops on index targets', () => {
  it('setBlockType converts a range and lifts list items first', () => {
    const editor = createEditor('- one\n- two\n\npara')
    const r = ai(editor, {
      op: 'setBlockType',
      target: { start: 0, end: 1 },
      type: 'heading',
      level: 2,
    })
    expect(r.applied).toBe(1)
    // the trailing empty paragraph is StarterKit's trailingNode, not ours
    expect(blocks(editor)).toEqual(['heading:one', 'heading:two', 'heading:para', 'paragraph:'])
    expect(r.results[0]).toMatchObject({ message: 'Set blocks 0-1 to h2.' })
  })

  it('setBlockType codeBlock sets and re-sets the language', () => {
    const editor = createEditor('print(1)')
    ai(editor, { op: 'setBlockType', target: { start: 0 }, type: 'codeBlock', language: 'python' })
    expect(editor.getMarkdown()).toContain('```python')
    const r = ai(editor, {
      op: 'setBlockType',
      target: { start: 0 },
      type: 'codeBlock',
      language: 'ruby',
    })
    expect(r.applied).toBe(1)
    expect(editor.getMarkdown()).toContain('```ruby')
    // no language = keep the fence as is
    expect(
      ai(editor, { op: 'setBlockType', target: { start: 0 }, type: 'codeBlock' }).applied,
    ).toBe(1)
    expect(editor.getMarkdown()).toContain('```ruby')
  })

  it('setBlockType blockquote wraps the range once and is idempotent', () => {
    const editor = createEditor('a\n\nb\n\nc')
    ai(editor, { op: 'setBlockType', target: { start: 0, end: 1 }, type: 'blockquote' })
    expect(blocks(editor)).toEqual(['blockquote:ab', 'paragraph:c'])
    // a range that starts inside a quote re-wraps everything into one quote
    ai(editor, { op: 'setBlockType', target: { start: 0, end: 1 }, type: 'blockquote' })
    expect(blocks(editor)).toEqual(['blockquote:abc', 'paragraph:'])
    expect(editor.getMarkdown()).not.toContain('> >')
    ai(editor, { op: 'setBlockType', target: { start: 0 }, type: 'paragraph' })
    expect(blocks(editor)).toEqual(['paragraph:a', 'paragraph:b', 'paragraph:c', 'paragraph:'])
  })

  it('nested wrappers unwrap fully and a doomed conversion leaves the document untouched', () => {
    const editor = createEditor('> - one\n> - two\n\n> ![p](assets/p.png)')
    const before = editor.getMarkdown()
    const quotedImage = ai(editor, {
      op: 'setBlockType',
      target: { start: 1 },
      type: 'heading',
      level: 2,
    })
    expect(quotedImage.results[0]).toMatchObject({
      ok: false,
      error: expect.stringContaining('no text'),
    })
    expect(editor.getMarkdown()).toBe(before)
    ai(editor, { op: 'setBlockType', target: { start: 0 }, type: 'paragraph' })
    expect(blocks(editor).slice(0, 2)).toEqual(['paragraph:one', 'paragraph:two'])
  })

  it('text-less targets (image, rule) are refused instead of restyling a neighbour', () => {
    const editor = createEditor('intro\n\n![p](assets/p.png)\n\n---\n\noutro')
    for (const op of [
      { op: 'setBlockType', target: { start: 1 }, type: 'heading', level: 1 },
      { op: 'toggleList', target: { start: 2 }, list: 'bullet' },
      { op: 'setStyle', target: { start: 1, end: 2 }, style: 'bold' },
      { op: 'setLink', target: { start: 2 }, href: 'https://x.test' },
    ] as Parameters<typeof runOps>[1]) {
      const r = ai(editor, op)
      expect(r.results[0]).toMatchObject({ ok: false, error: expect.stringContaining('no text') })
    }
    expect(blocks(editor)).toEqual([
      'paragraph:intro',
      'image:',
      'horizontalRule:',
      'paragraph:outro',
    ])
  })

  it('toggleList wraps paragraphs into one list and back', () => {
    const editor = createEditor('a\n\nb\n\nc')
    ai(editor, { op: 'toggleList', target: { start: 0, end: 1 }, list: 'bullet' })
    expect(blocks(editor)).toEqual(['bulletList:ab', 'paragraph:c'])
    ai(editor, { op: 'toggleList', target: { start: 0 }, list: 'task' })
    expect(editor.getMarkdown()).toContain('- [ ] a')
    ai(editor, { op: 'toggleList', target: { start: 0 }, list: 'task' })
    expect(blocks(editor)).toEqual(['paragraph:a', 'paragraph:b', 'paragraph:c'])
  })

  it('moveBlocks places the range after the anchor, -1 = start', () => {
    const editor = createEditor('a\n\nb\n\nc\n\nd')
    ai(editor, { op: 'moveBlocks', target: { start: 2, end: 3 }, after: -1 })
    expect(blocks(editor).map((b) => b.slice(-1))).toEqual(['c', 'd', 'a', 'b'])
    ai(editor, { op: 'moveBlocks', target: { start: 0 }, after: 3 })
    expect(blocks(editor).map((b) => b.slice(-1))).toEqual(['d', 'a', 'b', 'c'])
    const bad = ai(editor, { op: 'moveBlocks', target: { start: 0, end: 2 }, after: 1 })
    expect(bad.results[0]).toMatchObject({ ok: false, error: expect.stringContaining('inside') })
  })

  it('duplicateBlocks copies the range right after itself', () => {
    const editor = createEditor('a\n\nb')
    ai(editor, { op: 'duplicateBlocks', target: { start: 0, end: 1 } })
    expect(blocks(editor).map((b) => b.slice(-1))).toEqual(['a', 'b', 'a', 'b'])
  })

  it('setLink links matches and unlinks with href null', () => {
    const editor = createEditor('see the docs and the docs again')
    ai(editor, { op: 'setLink', target: { start: 0 }, find: 'docs', href: 'https://x.test' })
    expect(editor.getMarkdown().match(/\[docs\]\(https:\/\/x\.test\)/g)).toHaveLength(2)
    ai(editor, { op: 'setLink', target: { start: 0 }, href: null })
    expect(editor.getMarkdown()).not.toContain('](')
  })

  it('insertTable / insertHorizontalRule / insertImage land after the anchor block', () => {
    const editor = createEditor('a\n\nb')
    const r = ai(
      editor,
      { op: 'insertTable', after: 0, rows: 2, cols: 2, headerRow: false },
      { op: 'insertHorizontalRule', after: 1 },
      { op: 'insertImage', after: -1, src: 'assets/p.png', alt: 'pic' },
    )
    expect(r.applied).toBe(3)
    expect(blocks(editor).map((b) => b.split(':')[0])).toEqual([
      'image',
      'paragraph',
      'table',
      'paragraph',
      'horizontalRule',
      'paragraph',
    ])
    expect(r.blocksChanged).toBe(true)
  })

  it('editTable acts at the addressed cell of an index target', () => {
    const editor = createEditor('| h1 | h2 |\n| --- | --- |\n| a | b |')
    const r = ai(
      editor,
      { op: 'editTable', target: { start: 0 }, action: 'addRowAfter', row: 1 },
      { op: 'editTable', target: { start: 0 }, action: 'addColumnAfter', col: 1 },
    )
    expect(r.applied).toBe(2)
    const table = editor.state.doc.child(0)
    expect(table.type.name).toBe('table')
    expect(table.childCount).toBe(3)
    expect(table.child(0).childCount).toBe(3)
    const notTable = ai(editor, {
      op: 'editTable',
      target: { start: 0 },
      action: 'deleteRow',
      row: 9,
    })
    expect(notTable.results[0]).toMatchObject({ ok: false })
    const para = createEditor('text')
    expect(ai(para, { op: 'editTable', target: { start: 0 }, action: 'deleteTable' }).applied).toBe(
      0,
    )
  })

  it('every transaction an op dispatches carries the op meta', () => {
    const editor = createEditor('- item\n\npara')
    const metas: OpMeta[] = []
    editor.on('transaction', ({ transaction }) => {
      const m = transaction.getMeta(OP_META) as OpMeta | undefined
      if (m) metas.push(m)
      else if (transaction.docChanged) metas.push({ op: 'insertContent', source: 'ui', batch: -1 })
    })
    // index setBlockType on a list: unwrap + selection + convert, all stamped
    const r = ai(
      editor,
      { op: 'setBlockType', target: { start: 0 }, type: 'heading', level: 2 },
      { op: 'replaceText', target: { start: 1 }, find: 'para', replace: 'text' },
    )
    expect(r.applied).toBe(2)
    expect(metas.length).toBeGreaterThanOrEqual(2)
    expect(metas.every((m) => m.batch === metas[0]!.batch && m.source === 'ai')).toBe(true)
    expect(metas.map((m) => m.op)).toEqual(expect.arrayContaining(['setBlockType', 'replaceText']))
    // UI path (slash-style caret op) is stamped as ui
    metas.length = 0
    placeAfter(editor, 'text')
    uiOp(editor, { op: 'toggleList', target: 'selection', list: 'bullet' })
    expect(metas.length).toBeGreaterThan(0)
    expect(metas.every((m) => m.source === 'ui' && m.op === 'toggleList')).toBe(true)
  })

  it('mark-only AI ops (setStyle / setLink) highlight their matches', () => {
    const editor = createEditor('make TODO bold and link docs here')
    ai(
      editor,
      { op: 'setStyle', target: { start: 0 }, find: 'TODO', style: 'bold' },
      { op: 'setLink', target: { start: 0 }, find: 'docs', href: 'https://x.test' },
    )
    const marked = Array.from(editor.view.dom.querySelectorAll('.ai-changed')).map(
      (el) => el.textContent,
    )
    expect(marked).toEqual(['TODO', 'docs'])
  })

  it('a chain-based AI op (heading conversion) highlights the converted block', () => {
    const editor = createEditor('title\n\nbody')
    ai(editor, { op: 'setBlockType', target: { start: 0 }, type: 'heading', level: 1 })
    const marked = Array.from(editor.view.dom.querySelectorAll('.ai-changed')).map(
      (el) => el.textContent,
    )
    expect(marked.join('')).toContain('title')
    expect(marked.join('')).not.toContain('body')
  })

  it('an AI op highlights the touched range; UI ops do not', () => {
    const editor = createEditor('a\n\nb')
    ai(editor, { op: 'replaceText', target: { start: 0 }, find: 'a', replace: 'aaa' })
    const marked = editor.view.dom.querySelectorAll('.ai-changed')
    expect(marked.length).toBe(1)
    expect(marked[0]!.textContent).toBe('aaa')
    uiOp(editor, { op: 'setStyle', target: { start: 1 }, style: 'bold' })
    expect(editor.view.dom.querySelectorAll('.ai-changed').length).toBe(1)
  })
})

describe('UI ops on the selection', () => {
  it('insertTable from an empty paragraph replaces it and puts the caret in the first cell', () => {
    const editor = createEditor('intro\n\n')
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    expect(uiOp(editor, { op: 'insertTable', after: 'selection' })).toBe(true)
    expect(blocks(editor).map((b) => b.split(':')[0])).toEqual(['paragraph', 'table', 'paragraph'])
    expect(editor.isActive('table')).toBe(true)
  })

  it('inserting on a multi-block selection never swallows selected content', () => {
    const editor = createEditor('\n\nkeep me\n\nand me')
    editor.view.dispatch(
      editor.state.tr.setSelection(
        TextSelection.create(editor.state.doc, 1, editor.state.doc.content.size - 1),
      ),
    )
    uiOp(editor, { op: 'insertHorizontalRule', after: 'selection' })
    const md = editor.getMarkdown()
    expect(md).toContain('keep me')
    expect(md).toContain('and me')
    expect(md).toContain('---')
  })

  it('inserting from inside a list item lifts the item and inserts right after it', () => {
    const editor = createEditor('- one\n- two\n- three')
    placeAfter(editor, 'two')
    uiOp(editor, { op: 'insertHorizontalRule', after: 'selection' })
    expect(blocks(editor)).toEqual([
      'bulletList:one',
      'paragraph:two',
      'horizontalRule:',
      'bulletList:three',
      'paragraph:',
    ])
  })

  it('/quote inside a quote is a no-op (no split, no nesting); paragraph unwraps the caret block', () => {
    const editor = createEditor('> first\n>\n> second')
    placeAfter(editor, 'first')
    uiOp(editor, { op: 'setBlockType', target: 'selection', type: 'blockquote' })
    expect(blocks(editor)).toEqual(['blockquote:firstsecond', 'paragraph:'])
    expect(editor.getMarkdown()).not.toContain('> >')
    uiOp(editor, { op: 'setBlockType', target: 'selection', type: 'paragraph' })
    expect(blocks(editor)).toEqual(['paragraph:first', 'blockquote:second', 'paragraph:'])
  })

  it('insertHorizontalRule after the last block leaves a paragraph to type in', () => {
    const editor = createEditor('only')
    placeAfter(editor, 'only')
    uiOp(editor, { op: 'insertHorizontalRule', after: 'selection' })
    expect(blocks(editor).map((b) => b.split(':')[0])).toEqual([
      'paragraph',
      'horizontalRule',
      'paragraph',
    ])
    expect(editor.state.selection.$from.parent.type.name).toBe('paragraph')
    expect(editor.state.selection.from).toBeGreaterThan(editor.state.doc.child(0).nodeSize)
  })

  it('setStyle toggle on a selection behaves like the toolbar button', () => {
    const editor = createEditor('bold me please')
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 1, 8)))
    uiOp(editor, { op: 'setStyle', target: 'selection', style: 'bold', mode: 'toggle' })
    expect(editor.getMarkdown()).toContain('**bold me** please')
    uiOp(editor, { op: 'setStyle', target: 'selection', style: 'bold', mode: 'toggle' })
    expect(editor.getMarkdown()).not.toContain('**')
  })

  it('toggleList and setBlockType act on the caret block', () => {
    const editor = createEditor('a\n\nb')
    placeAfter(editor, 'b')
    uiOp(editor, { op: 'toggleList', target: 'selection', list: 'ordered' })
    expect(blocks(editor)).toEqual(['paragraph:a', 'orderedList:b', 'paragraph:'])
    uiOp(editor, { op: 'setBlockType', target: 'selection', type: 'heading', level: 3 })
    expect(blocks(editor)).toEqual(['paragraph:a', 'heading:b', 'paragraph:'])
  })

  it('editTable on the selection needs the caret inside a table', () => {
    const editor = createEditor('plain')
    placeAfter(editor, 'plain')
    expect(uiOp(editor, { op: 'editTable', target: 'selection', action: 'addRowAfter' })).toBe(
      false,
    )
  })
})
