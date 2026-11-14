import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { Editor, type JSONContent } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import type { StoryStyles } from '../src/renderer/editor/story-view'

const STYLES_CSS = readFileSync(join(__dirname, '..', 'src', 'renderer', 'styles.css'), 'utf8')

function blocksFor(selector: string): string {
  const out: string[] = []
  let at = STYLES_CSS.indexOf(selector)
  while (at >= 0) {
    const open = STYLES_CSS.indexOf('{', at)
    const close = STYLES_CSS.indexOf('}', open)
    out.push(STYLES_CSS.slice(open + 1, close))
    at = STYLES_CSS.indexOf(selector, close)
  }
  return out.join('\n')
}

const text = (value: string): JSONContent => ({ type: 'text', text: value })
const bold = (value: string): JSONContent => ({
  type: 'text',
  text: value,
  marks: [{ type: 'bold' }],
})
const para = (...content: JSONContent[]): JSONContent => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content,
})
const heading = (value: string, level = 1): JSONContent => ({
  type: 'docHeading',
  attrs: { docxIndex: null, level },
  content: [text(value)],
})

const editors: Editor[] = []
afterEach(() => editors.splice(0).forEach((editor) => editor.destroy()))

function createEditor(content: JSONContent[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  editors.push(editor)
  return editor
}

const STRUCTURAL_DOC: JSONContent[] = [
  para(bold('BACKGROUND')),
  para(text('The parties agree as follows.')),
  para(text('RENEWAL TERMS')),
  para(text('The term begins on the first day.')),
]

describe('story view: inferred structure', () => {
  it('draws faded markers for the inferred headings', () => {
    const editor = createEditor(STRUCTURAL_DOC)
    editor.commands.setStoryViewActive(true)
    const markers = [...editor.view.dom.querySelectorAll('.story-marker')]
    expect(markers.map((marker) => marker.textContent)).toEqual(['#', '#'])
    expect(markers.every((marker) => marker.classList.contains('story-marker-inferred'))).toBe(true)
    expect(editor.view.dom.querySelectorAll('.story-heading-inferred')).toHaveLength(2)
  })

  it('leaves the document unchanged', () => {
    const editor = createEditor(STRUCTURAL_DOC)
    const before = editor.state.doc.toJSON()
    editor.commands.setStoryViewActive(true)
    expect(editor.state.doc.toJSON()).toEqual(before)
  })

  it('draws no inferred markers when the document has a real heading', () => {
    const editor = createEditor([
      heading('Title', 2),
      para(bold('BACKGROUND')),
      para(text('The parties agree as follows.')),
    ])
    editor.commands.setStoryViewActive(true)
    const markers = [...editor.view.dom.querySelectorAll('.story-marker')]
    expect(markers.map((marker) => marker.textContent)).toEqual(['##'])
    expect(markers[0].classList.contains('story-marker-inferred')).toBe(false)
    expect(editor.view.dom.querySelectorAll('.story-heading-inferred')).toHaveLength(0)
  })

  it('draws no inferred markers when a paragraph style is a heading', () => {
    const editor = createEditor([
      { type: 'docParagraph', attrs: { docxIndex: null, styleId: 'H2' }, content: [text('Title')] },
      para(bold('BACKGROUND')),
      para(text('The parties agree as follows.')),
    ])
    const styles: StoryStyles = new Map([['H2', { name: 'heading 2', headingLevel: 2 }]])
    editor.storage.listNumbering.styles = styles as never
    editor.commands.setStoryViewActive(true)
    expect(editor.view.dom.querySelectorAll('.story-marker-inferred')).toHaveLength(0)
    expect(editor.view.dom.querySelectorAll('.story-heading-inferred')).toHaveLength(0)
  })
})

describe('outline view: inferred structure', () => {
  const mousedown = (el: Element) =>
    el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }))

  it('keeps the inferred headings visible with a handle each', () => {
    const editor = createEditor(STRUCTURAL_DOC)
    editor.commands.setOutlineViewActive(true)
    const inferred = [...editor.view.dom.querySelectorAll('.inferred-heading')]
    expect(inferred.map((node) => node.textContent?.replace(/[⊖⊕]/g, ''))).toEqual([
      'BACKGROUND',
      'RENEWAL TERMS',
    ])
    expect(inferred.every((node) => node.classList.contains('inferred-h1'))).toBe(true)
    expect(editor.view.dom.querySelectorAll('.outline-handle')).toHaveLength(2)
  })

  it('folds an inferred subtree through its inferred level', () => {
    const editor = createEditor([
      para(text('I. Background')),
      para(text('First body paragraph.')),
      para(text('1. Scope')),
      para(text('Second body paragraph.')),
    ])
    editor.commands.setOutlineViewActive(true)
    const headings = [...editor.view.dom.querySelectorAll('.inferred-heading')]
    expect(headings.map((node) => node.classList.contains('inferred-h1'))).toEqual([true, false])
    expect(headings.map((node) => node.classList.contains('inferred-h2'))).toEqual([false, true])
    mousedown(editor.view.dom.querySelectorAll('.outline-handle')[0])
    expect(headings.map((node) => node.classList.contains('outline-collapsed'))).toEqual([
      false,
      true,
    ])
  })

  it('draws no inferred rows when the document has a real heading', () => {
    const editor = createEditor([
      heading('Title'),
      para(bold('BACKGROUND')),
      para(text('The parties agree as follows.')),
    ])
    editor.commands.setOutlineViewActive(true)
    expect(editor.view.dom.querySelectorAll('.inferred-heading')).toHaveLength(0)
    expect(editor.view.dom.querySelectorAll('.outline-handle')).toHaveLength(1)
  })
})

describe('structural heading styling', () => {
  it('shows inferred paragraphs in outline view and fades them', () => {
    expect(blocksFor('.view-outline .doc-page > .inferred-heading')).toContain('display: block')
    expect(blocksFor('.view-outline .doc-page .inferred-heading')).toContain('opacity: 0.62')
  })

  it('fades inferred headings and markers in story view', () => {
    expect(blocksFor('.view-story .doc-page .story-heading-inferred')).toContain('opacity: 0.62')
    expect(blocksFor('.view-story .story-marker-inferred')).toContain('opacity: 0.62')
  })
})
