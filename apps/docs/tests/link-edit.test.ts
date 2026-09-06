/**
 * Hyperlink viewing/editing (alpha ledger r164): the caret on an existing
 * link must expose its range and address (the Ctrl+K dialog pre-fills and
 * edits it), Remove Link strips the mark but keeps the text, and a
 * tooltip-less link renders its target as the hover title without the
 * fallback round-tripping into a stored tooltip.
 */
import { describe, expect, it } from 'vitest'
import { Editor, getMarkRange } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'

const makeEditor = (content?: object) =>
  new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: content ?? {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          content: [
            { type: 'text', text: 'visit ' },
            {
              type: 'text',
              text: 'our site',
              marks: [{ type: 'link', attrs: { href: 'https://old.example.com' } }],
            },
            { type: 'text', text: ' today' },
          ],
        },
      ],
    },
  })

const linkRangeAtCaret = (editor: Editor) => {
  const { $from } = editor.state.selection
  return getMarkRange($from, editor.state.schema.marks.link!) ?? null
}

describe('link edit at caret (r164)', () => {
  it('exposes the full link range and address from a caret inside it', () => {
    const editor = makeEditor()
    editor.commands.setTextSelection(10) // inside "our site"
    const range = linkRangeAtCaret(editor)
    expect(range).not.toBeNull()
    expect(editor.state.doc.textBetween(range!.from, range!.to)).toBe('our site')
    expect(editor.getAttributes('link').href).toBe('https://old.example.com')
    editor.destroy()
  })

  it('address-only edit re-marks the run and keeps the text', () => {
    const editor = makeEditor()
    editor.commands.setTextSelection(10)
    const range = linkRangeAtCaret(editor)!
    editor
      .chain()
      .setTextSelection({ from: range.from, to: range.to })
      .setMark('link', { href: 'https://new.example.com' })
      .run()
    editor.commands.setTextSelection(10)
    expect(editor.getAttributes('link').href).toBe('https://new.example.com')
    expect(editor.state.doc.textContent).toBe('visit our site today')
    editor.destroy()
  })

  it('remove keeps the text and drops the mark', () => {
    const editor = makeEditor()
    editor.commands.setTextSelection(10)
    const range = linkRangeAtCaret(editor)!
    editor.chain().setTextSelection({ from: range.from, to: range.to }).unsetMark('link').run()
    expect(editor.state.doc.textContent).toBe('visit our site today')
    editor.commands.setTextSelection(10)
    expect(editor.isActive('link')).toBe(false)
    editor.destroy()
  })

  it('reads the address from the run at the trailing edge of a link', () => {
    // inclusive:false drops the mark from $head.marks() at the trailing
    // boundary — the range-based read must still see the address (bugbot)
    const editor = makeEditor()
    editor.commands.setTextSelection(15) // right after "our site"
    const range = linkRangeAtCaret(editor)
    expect(range).not.toBeNull()
    const mark = editor.state.doc.nodeAt(range!.from)?.marks.find((m) => m.type.name === 'link')
    expect(mark?.attrs.href).toBe('https://old.example.com')
    editor.destroy()
  })

  it('renders the target as hover title without storing it as a tooltip', () => {
    const editor = makeEditor()
    expect(editor.getHTML()).toContain('title="https://old.example.com"')
    // round-trip: the display-only fallback must not come back as a tooltip
    const roundTripped = makeEditor()
    roundTripped.commands.setContent(editor.getHTML())
    let tooltip: unknown = 'unset'
    roundTripped.state.doc.descendants((node) => {
      const mark = node.marks.find((m) => m.type.name === 'link')
      if (mark) tooltip = mark.attrs.tooltip
    })
    expect(tooltip).toBeNull()
    editor.destroy()
    roundTripped.destroy()
  })

  it('keeps a real stored tooltip as the hover title', () => {
    const editor = makeEditor({
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          content: [
            {
              type: 'text',
              text: 'docs',
              marks: [
                {
                  type: 'link',
                  attrs: { href: 'https://a.example.com', tooltip: 'Open the docs' },
                },
              ],
            },
          ],
        },
      ],
    })
    expect(editor.getHTML()).toContain('title="Open the docs"')
    editor.destroy()
  })
})
