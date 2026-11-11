import { afterEach, describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { Editor, type JSONContent } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'

const editors: Editor[] = []
afterEach(() => editors.splice(0).forEach((e) => e.destroy()))

const para = (text: string): JSONContent => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content: text ? [{ type: 'text', text }] : [],
})
const heading = (level: number, text: string): JSONContent => ({
  type: 'docHeading',
  attrs: { level },
  content: [{ type: 'text', text }],
})

function createEditor(content: JSONContent[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  editors.push(editor)
  return editor
}

const STYLES = readFileSync(
  join(dirname(fileURLToPath(import.meta.url)), '../src/renderer/styles.css'),
  'utf8',
)

describe('story view styles', () => {
  it('hides page chrome: gaps, headers and footers', () => {
    expect(STYLES).toContain('.view-story .doc-page .page-gap')
    expect(STYLES).toContain('.view-story .doc-page .page-hf')
  })

  it('normalizes document formatting to one column but keeps emphasis', () => {
    expect(STYLES).toContain('.view-story .doc-page')
    expect(STYLES).toContain('width: min(720px, 92%)')
    // spans lose their authored font/color (display only)
    expect(STYLES).toMatch(/\.view-story[\s\S]*?font-family: inherit !important/)
    // bold/italic/underline are not stripped by the story rules
    const storyBlock = STYLES.split('.view-story')[1]
    expect(storyBlock).not.toMatch(/font-weight:\s*400\s*!important/)
    expect(storyBlock).not.toMatch(/font-style:\s*normal\s*!important[^}]*span/)
  })

  it('hangs markdown-style markers in the margin', () => {
    expect(STYLES).toContain('.view-story .doc-page h1::before')
    expect(STYLES).toContain("content: '#'")
    expect(STYLES).toContain("content: '##'")
    expect(STYLES).toContain("content: '###'")
    expect(STYLES).toContain("content: '-'")
    expect(STYLES).toContain("content: '>'")
  })

  it('keeps tables as tables with plain cell text', () => {
    expect(STYLES).toContain('.view-story .doc-table')
    expect(STYLES).toContain('.view-story .doc-table td')
  })
})

describe('story view editing', () => {
  it('keeps headings and paragraphs editable', () => {
    const editor = createEditor([heading(1, 'Title'), para('body text')])
    expect(editor.view.dom.querySelector('h1')?.textContent).toBe('Title')
    expect(editor.view.dom.querySelector('p')?.textContent).toBe('body text')
  })

  it('typing still works and lands in the document', () => {
    const editor = createEditor([para('hello')])
    editor.chain().focus().insertContent(' world').run()
    expect(editor.getText()).toContain('world')
    expect(editor.getText()).toContain('hello')
  })

  it('bold typing survives (emphasis is display-preserved, not stripped)', () => {
    const editor = createEditor([
      {
        type: 'docParagraph',
        attrs: { docxIndex: null },
        content: [{ type: 'text', text: 'plain', marks: [{ type: 'bold' }] }],
      },
    ])
    editor.chain().focus().insertContent(' text').run()
    expect(editor.getText()).toContain('plain')
    // the original bold run is untouched by the view (CSS-only hiding)
    expect(editor.getJSON().content?.[0]?.content?.[0]?.marks).toEqual([{ type: 'bold' }])
  })
})
