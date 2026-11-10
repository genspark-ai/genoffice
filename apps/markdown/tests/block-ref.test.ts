import { afterAll, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { wikiFragmentPosition } from '../src/renderer/editor/wikilink'

const editors: Editor[] = []
afterAll(() => {
  for (const e of editors) e.destroy()
})

function openDoc(source: string): Editor {
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
  editor.commands.setContent(source, { contentType: 'markdown' })
  return editor
}

describe('block ref fragments', () => {
  it('lands on the top-level block whose text ends with the ^id', () => {
    const editor = openDoc('First line.\n\nSecond line with the target. ^abc123\n\nAfter.')
    const pos = wikiFragmentPosition(editor.state.doc, '^abc123')
    expect(pos).not.toBeNull()
    // the position is inside the second block, not the first
    const $pos = editor.state.doc.resolve(pos!)
    expect($pos.parent.textContent).toContain('the target')
  })

  it('is null when no block carries the id', () => {
    const editor = openDoc('Nothing here.')
    expect(wikiFragmentPosition(editor.state.doc, '^missing')).toBeNull()
  })

  it('an id is distinct from its prefix (abc vs abc123)', () => {
    const editor = openDoc('Target line. ^abc123')
    expect(wikiFragmentPosition(editor.state.doc, '^abc')).toBeNull()
  })

  it('matches the id regardless of case', () => {
    // the fragment is lower-cased before lookup, so the marker test has to be
    // case-insensitive too — [[Note#^ABC123]] against ^AbC123 used to miss
    const editor = openDoc('Target line. ^AbC123')
    const pos = wikiFragmentPosition(editor.state.doc, '^ABC123')
    expect(pos).not.toBeNull()
    expect(editor.state.doc.resolve(pos!).parent.textContent).toContain('Target line')
  })

  it('lands on a list item that is not the last one', () => {
    // the marker sits inside the item's paragraph; walking only top-level
    // nodes tested the whole list's text and reported the list's start
    const editor = openDoc('- one ^aaa\n- two')
    const pos = wikiFragmentPosition(editor.state.doc, '^aaa')
    expect(pos).not.toBeNull()
    const $pos = editor.state.doc.resolve(pos!)
    expect($pos.parent.textContent).toBe('one ^aaa')
  })
})
