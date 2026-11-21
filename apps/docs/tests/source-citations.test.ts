/** Citations are CITATION fields; the Style: picker rewrites them and the bibliography. */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import type { SourceInfo } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  bibliographyRange,
  citationFields,
  citedOrder,
  insertCitationField,
  restyleCitations,
} from '../src/renderer/editor/source-citations'
import { insertBibliography } from '../src/renderer/commands/reference-commands'

const SOURCES: SourceInfo[] = [
  { tag: 'Kay93', type: 'JournalArticle', author: 'Kay, Alan', title: 'Smalltalk', year: '1993' },
  {
    tag: 'Mat11',
    type: 'Book',
    author: 'Mattmann, Chris; Zitting, Jukka',
    title: 'Tika in Action',
    year: '2011',
    publisher: 'Manning',
  },
]

describe('field-backed citations', () => {
  let editor: Editor
  beforeEach(() => {
    editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          { type: 'docParagraph', content: [{ type: 'text', text: 'First ' }] },
          { type: 'docParagraph', content: [{ type: 'text', text: 'Second ' }] },
          { type: 'docParagraph', content: [{ type: 'text', text: 'end' }] },
        ],
      },
    })
  })
  afterEach(() => {
    editor.destroy()
    vi.restoreAllMocks()
  })

  const cite = (pos: number, tag: string, style: 'apa' | 'ieee' = 'apa') => {
    editor.commands.setTextSelection(pos)
    insertCitationField(
      editor,
      [],
      SOURCES.find((s) => s.tag === tag)!,
      SOURCES,
      style,
    )
  }
  const shown = () => citationFields(editor.state.doc).map((f) => f.text)
  const bibliography = () => {
    const out: string[] = []
    editor.state.doc.forEach((node) => {
      if (String(node.attrs.genXml ?? '').includes('w:pStyle w:val="Bibliography"'))
        out.push((node.attrs.fieldDisplay as { left: string }).left)
    })
    return out
  }

  it('a citation is a CITATION field showing the source in the style', () => {
    cite(7, 'Mat11')
    expect(citationFields(editor.state.doc)).toEqual([
      { from: 7, to: 33, instr: 'CITATION Mat11 \\l 1033', text: '(Mattmann & Zitting, 2011)' },
    ])
    // typing after it is plain text again
    editor.commands.insertContent('x')
    expect(shown()).toEqual(['(Mattmann & Zitting, 2011)'])
  })

  it('two citations of one source side by side stay two fields', () => {
    cite(7, 'Kay93')
    cite(editor.state.selection.to, 'Kay93')
    expect(shown()).toEqual(['(Kay, 1993)', '(Kay, 1993)'])
  })

  it('Style: rewrites every citation, and a numbered style numbers them as the text cites them', () => {
    cite(7, 'Mat11')
    cite(editor.state.doc.child(0).nodeSize + 8, 'Kay93')
    expect(citedOrder(editor.state.doc)).toEqual(['Mat11', 'Kay93'])
    restyleCitations(editor, [], SOURCES, 'ieee')
    expect(shown()).toEqual(['[1]', '[2]'])
    restyleCitations(editor, [], SOURCES, 'mla')
    expect(shown()).toEqual(['(Mattmann and Zitting)', '(Kay)'])
    // a citation put in before the first renumbers what follows
    restyleCitations(editor, [], SOURCES, 'ieee')
    cite(1, 'Kay93', 'ieee')
    expect(shown()).toEqual(['[1]', '[2]', '[1]'])
  })

  it('the bibliography is a BIBLIOGRAPHY field that follows the style and the sources', () => {
    cite(7, 'Mat11')
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    insertBibliography(editor, [], SOURCES, 'apa')
    expect(bibliography()).toEqual([
      'Kay, A. (1993). Smalltalk.',
      'Mattmann, C., & Zitting, J. (2011). Tika in Action. Manning.',
    ])
    expect(bibliographyRange(editor.state.doc, [])).not.toBeNull()
    restyleCitations(editor, [], SOURCES, 'ieee')
    expect(bibliography()).toEqual([
      '[1] C. Mattmann and J. Zitting, Tika in Action, Manning, 2011.',
      '[2] A. Kay, “Smalltalk”, 1993.',
    ])
    // asking for a bibliography again rewrites the one there
    insertBibliography(editor, [], SOURCES, 'ieee')
    expect(bibliography()).toHaveLength(2)
    const displays: string[] = []
    editor.state.doc.forEach((node) => displays.push(JSON.stringify(node.attrs.fieldDisplay ?? '')))
    expect(displays.join('')).toContain('{"text":"Tika in Action","italic":true}')
  })

  it('a citation that renumbers the ones after it and joins the bibliography is one step to undo', () => {
    cite(editor.state.doc.child(0).nodeSize + 8, 'Mat11', 'ieee')
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    insertBibliography(editor, [], SOURCES, 'ieee')
    const before = editor.getJSON()
    cite(7, 'Kay93', 'ieee')
    expect(shown()).toEqual(['[1]', '[2]'])
    expect(bibliography()[0]).toContain('A. Kay')
    editor.commands.undo()
    expect(editor.getJSON()).toEqual(before)
    editor.commands.redo()
    expect(shown()).toEqual(['[1]', '[2]'])
    // a change of style is one step too
    const numbered = editor.getJSON()
    restyleCitations(editor, [], SOURCES, 'apa')
    expect(shown()).toEqual(['(Kay, 1993)', '(Mattmann & Zitting, 2011)'])
    editor.commands.undo()
    expect(editor.getJSON()).toEqual(numbered)
  })

  it('with no sources there is nothing to list', () => {
    const alert = vi.spyOn(window, 'alert').mockImplementation(() => {})
    insertBibliography(editor, [], [], 'apa')
    expect(alert).toHaveBeenCalledTimes(1)
    expect(bibliography()).toEqual([])
  })
})
