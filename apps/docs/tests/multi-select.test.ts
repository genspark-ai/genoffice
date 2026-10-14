import { afterEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  clearMultiSelect,
  isMultiSelectModifier,
  multiSelectRanges,
} from '../src/renderer/editor/multi-select'

/**
 * Ctrl/⌘+click adds a non-contiguous selection range (#2013). ProseMirror's
 * Selection holds one {from, to}, so the added ranges live in plugin state and
 * are painted as decorations; the live selection stays the one the core owns.
 */

const para = (t: string) => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content: [{ type: 'text', text: t }],
})

const editors = new Set<Editor>()
afterEach(() => {
  for (const e of editors) e.destroy()
  editors.clear()
  vi.useRealTimers()
})

function makeEditor(text = 'alpha bravo charlie'): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content: [para(text)] },
  })
  editors.add(editor)
  return editor
}

/**
 * Fire a real ctrl+click on the editor's DOM at a character offset. The plugin
 * reads ProseMirror's own selection after the click, so the position has to be
 * set first — that is what the browser does by dragging/word-selecting.
 */
async function ctrlClick(editor: Editor, from: number, to: number): Promise<void> {
  editor.commands.setTextSelection({ from, to })
  const dom = editor.view.dom
  const ev = new MouseEvent('mousedown', {
    ctrlKey: true,
    bubbles: true,
    cancelable: true,
  })
  dom.dispatchEvent(ev)
  // the handler defers one macrotask so PM has placed its selection
  await new Promise((r) => setTimeout(r, 0))
}

describe('isMultiSelectModifier', () => {
  const ev = (init: Partial<MouseEvent>) =>
    ({ metaKey: false, ctrlKey: false, altKey: false, shiftKey: false, ...init }) as MouseEvent

  it('accepts ctrl or cmd on its own', () => {
    expect(isMultiSelectModifier(ev({ ctrlKey: true }))).toBe(true)
    expect(isMultiSelectModifier(ev({ metaKey: true }))).toBe(true)
  })

  it('rejects the combos that already mean something else', () => {
    expect(isMultiSelectModifier(ev({}))).toBe(false)
    expect(isMultiSelectModifier(ev({ ctrlKey: true, shiftKey: true }))).toBe(false)
    expect(isMultiSelectModifier(ev({ ctrlKey: true, altKey: true }))).toBe(false)
  })
})

describe('MultiSelect', () => {
  it('starts with no added ranges', () => {
    expect(multiSelectRanges(makeEditor().view)).toEqual([])
  })

  it('a ctrl+click adds the range the click produced', async () => {
    const editor = makeEditor()
    // "alpha" = positions 1..6
    await ctrlClick(editor, 1, 6)
    expect(multiSelectRanges(editor.view)).toEqual([{ from: 1, to: 6 }])
  })

  it('a second ctrl+click adds a second, disjoint range', async () => {
    const editor = makeEditor()
    await ctrlClick(editor, 1, 6)
    await ctrlClick(editor, 12, 19)
    expect(multiSelectRanges(editor.view)).toEqual([
      { from: 1, to: 6 },
      { from: 12, to: 19 },
    ])
  })

  it('the live selection still moves — PM keeps owning it', async () => {
    const editor = makeEditor()
    await ctrlClick(editor, 1, 6)
    await ctrlClick(editor, 12, 19)
    // the core's selection is the LAST clicked range, as the browser behaves
    expect(editor.state.selection.from).toBe(12)
    expect(editor.state.selection.to).toBe(19)
  })

  it('a collapsed ctrl+click adds nothing', async () => {
    const editor = makeEditor()
    editor.commands.setTextSelection(3)
    editor.view.dom.dispatchEvent(
      new MouseEvent('mousedown', { ctrlKey: true, bubbles: true, cancelable: true }),
    )
    await new Promise((r) => setTimeout(r, 0))
    expect(multiSelectRanges(editor.view)).toEqual([])
  })

  it('a plain click adds nothing', async () => {
    const editor = makeEditor()
    editor.commands.setTextSelection({ from: 1, to: 6 })
    editor.view.dom.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }))
    await new Promise((r) => setTimeout(r, 0))
    expect(multiSelectRanges(editor.view)).toEqual([])
  })

  it('an edit before a range moves it along with the text', async () => {
    const editor = makeEditor()
    await ctrlClick(editor, 12, 19)
    // type "X" at the very start: the range shifts right by one
    editor.commands.insertContentAt(1, 'X')
    expect(multiSelectRanges(editor.view)).toEqual([{ from: 13, to: 20 }])
  })

  it('a range the edit swallowed is dropped, not left as a zero-width sliver', async () => {
    const editor = makeEditor()
    await ctrlClick(editor, 1, 6)
    // delete exactly the range
    editor.commands.deleteRange({ from: 1, to: 6 })
    expect(multiSelectRanges(editor.view)).toEqual([])
  })

  it('clearMultiSelect drops every added range', async () => {
    const editor = makeEditor()
    await ctrlClick(editor, 1, 6)
    await ctrlClick(editor, 12, 19)
    expect(multiSelectRanges(editor.view)).toHaveLength(2)
    clearMultiSelect(editor.view)
    expect(multiSelectRanges(editor.view)).toEqual([])
  })

  it('clearMultiSelect on an empty state dispatches nothing', () => {
    const editor = makeEditor()
    const before = editor.state
    clearMultiSelect(editor.view)
    expect(editor.state).toBe(before)
  })

  it('paints a decoration over each added range', async () => {
    const editor = makeEditor()
    await ctrlClick(editor, 1, 6)
    // the plugin's decorations prop paints this class; assert through the DOM
    // that the range carries it and covers exactly the clicked text
    const el = editor.view.dom.querySelector('.doc-multi-range')
    expect(el).not.toBeNull()
    expect(el!.textContent).toBe('alpha')
  })

  it('leaves the document content itself untouched', async () => {
    const editor = makeEditor()
    await ctrlClick(editor, 1, 6)
    expect(editor.state.doc.textContent).toBe('alpha bravo charlie')
  })
})
