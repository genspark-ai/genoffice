import { afterAll, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { collectOutline } from '../src/renderer/editor/outline'

// Undestroyed views leave DOMObserver flush timers that fire after jsdom teardown
// ("document is not defined" unhandled error). Editors here are shared per describe,
// so destroy them once at the end of the file.
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

function setMarkdown(editor: Editor, md: string): void {
  editor.commands.setContent(md, { contentType: 'markdown' })
}

describe('collectOutline', () => {
  const editor = createEditor()

  it('collects headings in document order with levels and text', () => {
    setMarkdown(
      editor,
      '# Title\n\nSome body text.\n\n## Section one\n\n### Deep dive\n\n## Section two\n',
    )
    const items = collectOutline(editor)
    expect(items.map((i) => [i.level, i.text])).toEqual([
      [1, 'Title'],
      [2, 'Section one'],
      [3, 'Deep dive'],
      [2, 'Section two'],
    ])
  })

  it('skips empty headings so the pane never shows blank rows', () => {
    setMarkdown(editor, '# Kept\n\n#\n\n## Also kept\n')
    const items = collectOutline(editor)
    expect(items.map((i) => i.text)).toEqual(['Kept', 'Also kept'])
  })

  it('returns an empty list for a document without headings', () => {
    setMarkdown(editor, 'Just a paragraph.\n\n- a list\n- of things\n')
    expect(collectOutline(editor)).toEqual([])
  })

  it('positions resolve to a cursor inside the heading (jump targets are valid)', () => {
    setMarkdown(editor, '# First\n\n## Second\n')
    const items = collectOutline(editor)
    expect(items).toHaveLength(2)
    for (const item of items) {
      const $pos = editor.state.doc.resolve(item.pos)
      // must not throw: the outline click handler builds a selection here
      const selection = TextSelection.near($pos)
      expect(selection).toBeTruthy()
    }
  })

  it('reflects edits after the content changes', () => {
    setMarkdown(editor, '# Before\n')
    expect(collectOutline(editor).map((i) => i.text)).toEqual(['Before'])
    setMarkdown(editor, '# After\n\n## Added\n')
    expect(collectOutline(editor).map((i) => [i.level, i.text])).toEqual([
      [1, 'After'],
      [2, 'Added'],
    ])
  })
})
