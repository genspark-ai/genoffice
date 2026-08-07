import { afterAll, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'

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

/** parse → serialize → parse must be structurally stable */
function roundTrip(editor: Editor, md: string): { out: string; stable: boolean } {
  const manager = editor.markdown!
  const first = manager.parse(md)
  const out = manager.serialize(first)
  const second = manager.parse(out)
  return { out, stable: JSON.stringify(first) === JSON.stringify(second) }
}

describe('markdown round-trip for M2 nodes', () => {
  const editor = createEditor()

  it('callout with a non-default type', () => {
    const md = ':::callout {type="warning"}\nBe careful.\n:::'
    const { out, stable } = roundTrip(editor, md)
    expect(out).toContain(':::callout')
    expect(out).toContain('type="warning"')
    expect(out).toContain('Be careful.')
    expect(stable).toBe(true)
  })

  it('callout with the default type omits the attribute', () => {
    const { out, stable } = roundTrip(editor, ':::callout\nNote text.\n:::')
    expect(out).toContain(':::callout')
    expect(out).not.toContain('type=')
    expect(stable).toBe(true)
  })

  it('toggle keeps its summary and drops editor-only open state', () => {
    const md = ':::toggle {summary="More info"}\nHidden body.\n:::'
    const { out, stable } = roundTrip(editor, md)
    expect(out).toContain(':::toggle')
    expect(out).toContain('summary="More info"')
    expect(out).not.toContain('open=')
    expect(stable).toBe(true)
  })

  it('callout containing a list', () => {
    const md = ':::callout {type="tip"}\n- one\n- two\n:::'
    const { out, stable } = roundTrip(editor, md)
    expect(out).toContain('- one')
    expect(stable).toBe(true)
  })

  it('task list', () => {
    const md = '- [ ] open item\n- [x] done item'
    const { out, stable } = roundTrip(editor, md)
    expect(out).toContain('- [ ] open item')
    expect(out).toContain('- [x] done item')
    expect(stable).toBe(true)
  })

  it('table', () => {
    const md = '| Name | Value |\n| --- | --- |\n| a | 1 |\n| b | 2 |'
    const { out, stable } = roundTrip(editor, md)
    // the serializer pads cells to align columns
    expect(out).toContain('| Name | Value |')
    expect(out).toMatch(/\| a\s+\| 1\s+\|/)
    expect(stable).toBe(true)
  })

  it('image keeps the authored relative path in markdown', () => {
    const md = '![diagram](assets/diagram.png)'
    const { out, stable } = roundTrip(editor, md)
    expect(out).toContain('![diagram](assets/diagram.png)')
    expect(stable).toBe(true)
  })

  it('plain constructs still round-trip alongside the custom nodes', () => {
    const md = '# Title\n\n> quoted\n\n```js\ncode()\n```\n\n---'
    const { out, stable } = roundTrip(editor, md)
    expect(out).toContain('# Title')
    expect(out).toContain('```js')
    expect(stable).toBe(true)
  })
})

describe('highlight and code block language', () => {
  const editor = createEditor()

  it('==highlight== round-trips', () => {
    const { out, stable } = roundTrip(editor, 'some ==marked text== here')
    expect(out).toContain('==marked text==')
    expect(stable).toBe(true)
  })

  it('code block language survives the round-trip', () => {
    const { out, stable } = roundTrip(editor, '```python\nprint(1)\n```')
    expect(out).toContain('```python')
    expect(stable).toBe(true)
  })
})

describe('underline (ships inside StarterKit v3)', () => {
  const editor = createEditor()

  it('toggleUnderline applies the mark', () => {
    editor.commands.setContent('word', { contentType: 'markdown' })
    editor.commands.selectAll()
    expect(editor.commands.toggleUnderline()).toBe(true)
    expect(editor.isActive('underline')).toBe(true)
  })

  it('an underline applied in the editor survives serialize → reparse', () => {
    editor.commands.setContent('word', { contentType: 'markdown' })
    editor.commands.selectAll()
    editor.commands.toggleUnderline()
    const md = editor.getMarkdown()
    expect(md).toContain('word')
    editor.commands.setContent(md, { contentType: 'markdown' })
    editor.commands.selectAll()
    expect(editor.isActive('underline')).toBe(true)
  })
})
