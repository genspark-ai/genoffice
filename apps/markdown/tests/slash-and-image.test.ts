import { afterEach, describe, expect, it } from 'vitest'
import { buildSlashItems, filterSlashItems } from '../src/renderer/editor/slashCommand'
import { resolveImageSrc } from '../src/renderer/editor/localImage'

// Undestroyed views leave DOMObserver flush timers that fire after jsdom teardown
// ("document is not defined" unhandled error) — destroy every editor we create.
const editors: import('@tiptap/core').Editor[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
})

describe('filterSlashItems', () => {
  const items = buildSlashItems()

  it('returns everything for an empty query', () => {
    expect(filterSlashItems(items, '').length).toBe(items.length)
  })

  it('matches by id', () => {
    const hits = filterSlashItems(items, 'table')
    expect(hits.map((i) => i.id)).toContain('table')
  })

  it('matches by keyword', () => {
    const hits = filterSlashItems(items, 'todo')
    expect(hits.map((i) => i.id)).toContain('task')
  })

  it('image item only appears when an insert handler is provided', () => {
    expect(items.map((i) => i.id)).not.toContain('image')
    const withImage = buildSlashItems({ insertImage: () => {} })
    expect(withImage.map((i) => i.id)).toContain('image')
  })
})

describe('resolveImageSrc', () => {
  it('passes through absolute URLs', () => {
    expect(resolveImageSrc('https://x.test/a.png', '/docs')).toBe('https://x.test/a.png')
    expect(resolveImageSrc('data:image/png;base64,AA', '/docs')).toBe('data:image/png;base64,AA')
  })

  it('maps absolute POSIX paths to md-asset URLs', () => {
    expect(resolveImageSrc('/tmp/pic.png', null)).toBe('md-asset:///tmp/pic.png')
  })

  it('resolves relative paths against the document directory', () => {
    expect(resolveImageSrc('assets/p.png', '/Users/me/notes')).toBe(
      'md-asset:///Users/me/notes/assets/p.png',
    )
  })

  it('encodes special characters per segment', () => {
    expect(resolveImageSrc('assets/图 1.png', '/n')).toBe('md-asset:///n/assets/%E5%9B%BE%201.png')
  })

  it('returns the authored value when relative with no base dir', () => {
    expect(resolveImageSrc('assets/p.png', null)).toBe('assets/p.png')
  })
})

describe('slash block commands inside a list (lift-first behavior)', () => {
  async function editorWith(md: string) {
    const { Editor } = await import('@tiptap/core')
    const { buildExtensions } = await import('../src/renderer/editor/extensions')
    const editor = new Editor({
      extensions: buildExtensions({
        slashController: { onOpen() {}, onUpdate() {}, onKeyDown: () => false, onClose() {} },
        slashItems: () => [],
      }),
      content: '',
    })
    editor.commands.setContent(md, { contentType: 'markdown' })
    editors.push(editor)
    return editor
  }

  /** absolute position right after `needle`, mimicking a caret after typed "/" */
  function caretAfter(editor: import('@tiptap/core').Editor, needle: string): number {
    let found = -1
    editor.state.doc.descendants((node, pos) => {
      if (found >= 0) return false
      if (node.isText && node.text?.includes(needle)) {
        found = pos + node.text.indexOf(needle) + needle.length
      }
      return found < 0
    })
    expect(found).toBeGreaterThan(0)
    return found
  }

  it('callout from a list item lifts out of the list instead of silently failing', async () => {
    const editor = await editorWith('- first\n- second x/')
    const caret = caretAfter(editor, 'x/')
    editor.commands.setTextSelection(caret)
    const callout = buildSlashItems().find((i) => i.id === 'callout')!
    callout.run(editor, { from: caret - 1, to: caret })
    expect(editor.isActive('callout')).toBe(true)
    const md = editor.getMarkdown()
    expect(md).toContain(':::callout')
    expect(md).toContain('- first')
    expect(md).not.toContain('- second x')
  })

  it('heading from a task item lifts out of the task list', async () => {
    const editor = await editorWith('- [ ] todo x/')
    const caret = caretAfter(editor, 'x/')
    editor.commands.setTextSelection(caret)
    const h2 = buildSlashItems().find((i) => i.id === 'h2')!
    h2.run(editor, { from: caret - 1, to: caret })
    expect(editor.isActive('heading', { level: 2 })).toBe(true)
    expect(editor.getMarkdown()).toContain('## todo x')
  })

  it('toggle from a nested list item escapes every level', async () => {
    const editor = await editorWith('- outer\n  - inner x/')
    const caret = caretAfter(editor, 'x/')
    editor.commands.setTextSelection(caret)
    const toggle = buildSlashItems().find((i) => i.id === 'toggle')!
    toggle.run(editor, { from: caret - 1, to: caret })
    expect(editor.isActive('toggle')).toBe(true)
    expect(editor.isActive('listItem')).toBe(false)
  })
})

describe('block keymap commands (secondbrain parity)', () => {
  async function editorWith(md: string) {
    const { Editor } = await import('@tiptap/core')
    const { buildExtensions } = await import('../src/renderer/editor/extensions')
    const editor = new Editor({
      extensions: buildExtensions({
        slashController: { onOpen() {}, onUpdate() {}, onKeyDown: () => false, onClose() {} },
        slashItems: () => [],
      }),
      content: '',
    })
    editor.commands.setContent(md, { contentType: 'markdown' })
    editors.push(editor)
    return editor
  }

  function placeIn(editor: import('@tiptap/core').Editor, needle: string) {
    let found = -1
    editor.state.doc.descendants((node, pos) => {
      if (found >= 0) return false
      if (node.isText && node.text?.includes(needle)) found = pos + node.text.indexOf(needle) + 1
      return found < 0
    })
    editor.commands.setTextSelection(found)
  }

  it('duplicateBlock copies the current top-level block below itself', async () => {
    const editor = await editorWith('first\n\nsecond\n\nthird')
    placeIn(editor, 'second')
    expect(editor.commands.duplicateBlock()).toBe(true)
    const texts: string[] = []
    editor.state.doc.forEach((n) => texts.push(n.textContent))
    expect(texts).toEqual(['first', 'second', 'second', 'third'])
  })

  it('deleteBlock removes the block and keeps at least one paragraph', async () => {
    const editor = await editorWith('only block')
    placeIn(editor, 'only')
    expect(editor.commands.deleteBlock()).toBe(true)
    expect(editor.state.doc.childCount).toBe(1)
    expect(editor.state.doc.textContent).toBe('')
  })

  it('moveBlockUp / moveBlockDown swap with the sibling and stop at edges', async () => {
    const editor = await editorWith('alpha\n\nbeta\n\ngamma')
    placeIn(editor, 'beta')
    expect(editor.commands.moveBlockUp()).toBe(true)
    let texts: string[] = []
    editor.state.doc.forEach((n) => texts.push(n.textContent))
    expect(texts).toEqual(['beta', 'alpha', 'gamma'])
    // caret followed the block — moving up again hits the edge
    expect(editor.commands.moveBlockUp()).toBe(false)
    expect(editor.commands.moveBlockDown()).toBe(true)
    texts = []
    editor.state.doc.forEach((n) => texts.push(n.textContent))
    expect(texts).toEqual(['alpha', 'beta', 'gamma'])
  })
})

describe('resolveImageSrc on Windows paths', () => {
  it('keeps the drive colon unencoded in md-asset URLs', () => {
    expect(resolveImageSrc('C:\\Users\\me\\pic.png', null)).toBe('md-asset:///C:/Users/me/pic.png')
  })

  it('resolves relative paths against a Windows base dir', () => {
    expect(resolveImageSrc('assets/p.png', 'C:\\notes')).toBe('md-asset:///C:/notes/assets/p.png')
  })
})

describe('buildPrintHtml', () => {
  it('strips editor-only code block chrome from the print output', async () => {
    const { buildPrintHtml } = await import('../src/renderer/export/printHtml')
    const root = document.createElement('div')
    root.innerHTML =
      '<div class="md-codeblock"><div class="md-codeblock-bar"><select></select><button>Copy</button></div><pre><code>x = 1</code></pre></div>'
    const html = buildPrintHtml(root, 'Doc')
    expect(html).toContain('x = 1')
    expect(html).not.toContain('md-codeblock-bar')
    expect(html).not.toContain('<select')
  })
})
