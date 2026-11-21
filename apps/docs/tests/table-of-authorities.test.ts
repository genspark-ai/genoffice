/** Citations are marked after their text, collected by category, and listed in a table of authorities. */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { citationOf, citationSwitches } from '../src/renderer/editor/citations'
import {
  citationCategories,
  collectCitations,
  insertTableOfAuthorities,
  markCitation,
  updateTableOfAuthorities,
} from '../src/renderer/commands/reference-commands'

const heading = (category: number) => `Heading ${category}`

describe('table of authorities', () => {
  let editor: Editor
  beforeEach(() => {
    editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          { type: 'docParagraph', content: [{ type: 'text', text: 'See Baldwin.' }] },
          { type: 'docParagraph', content: [{ type: 'text', text: 'Under the statute.' }] },
          { type: 'docParagraph', content: [{ type: 'text', text: 'Baldwin again.' }] },
          { type: 'docParagraph', content: [{ type: 'text', text: 'end' }] },
        ],
      },
    })
    markCitation(editor, 12, { long: 'Baldwin v. Alberti', category: 1 })
    markCitation(editor, 32, { long: 'RCW 4.16', category: 2 })
    markCitation(editor, 49, { long: 'Baldwin v. Alberti', category: 1 })
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
  })
  afterEach(() => {
    editor.destroy()
    vi.restoreAllMocks()
  })

  const tableLines = () => {
    const out: Array<[string, string]> = []
    editor.state.doc.forEach((node) => {
      if (/TOAHeading|TableofAuthorities/.test(String(node.attrs.genXml ?? ''))) {
        const shown = node.attrs.fieldDisplay as { left: string; right: string }
        out.push([shown.left, shown.right])
      }
    })
    return out
  }

  it('reads and writes the switches of a citation', () => {
    expect(citationOf('\\l "Baldwin v. Alberti, 58 Wn. 2d 243" \\s "Baldwin" \\c 3')).toEqual({
      long: 'Baldwin v. Alberti, 58 Wn. 2d 243',
      short: 'Baldwin',
      category: 3,
    })
    // a citation that names no category is a case; a later mention carries the short form alone
    expect(citationOf('\\s "Baldwin"')).toEqual({ long: '', short: 'Baldwin', category: 1 })
    expect(citationSwitches({ long: 'A "quoted" case', category: 2 })).toBe(
      '\\l "A quoted case" \\c 2',
    )
  })

  it('finds the marked citations and the categories in use', () => {
    expect(collectCitations(editor).map((c) => [c.long, c.category])).toEqual([
      ['Baldwin v. Alberti', 1],
      ['RCW 4.16', 2],
      ['Baldwin v. Alberti', 1],
    ])
    expect(citationCategories(editor)).toEqual([1, 2])
  })

  it('lists one category under its heading, a repeated citation once with every page it is on', () => {
    insertTableOfAuthorities(editor, [], () => [2, 5], heading, 1)
    expect(tableLines()).toEqual([
      ['Heading 1', ''],
      ['Baldwin v. Alberti', '2, 5'],
    ])
  })

  it('lists every category in use when none is named, and leaves pages blank without a layout', () => {
    insertTableOfAuthorities(editor, [], () => null, heading)
    expect(tableLines()).toEqual([
      ['Heading 1', ''],
      ['Baldwin v. Alberti', ''],
      ['Heading 2', ''],
      ['RCW 4.16', ''],
    ])
  })

  it('a later mention marked by its short citation is listed under the long one, in its category', () => {
    markCitation(
      editor,
      5,
      { long: 'RCW 4.16.080 (2019)', short: 'the statute', category: 2 },
      [30],
    )
    expect(
      collectCitations(editor)
        .filter((c) => c.short === 'the statute')
        .map((c) => [c.long, c.category]),
    ).toEqual([
      ['RCW 4.16.080 (2019)', 2],
      ['RCW 4.16.080 (2019)', 2],
    ])
    insertTableOfAuthorities(editor, [], (positions) => positions.map((_, i) => i + 1), heading, 2)
    expect(tableLines()).toContainEqual(['RCW 4.16.080 (2019)', '1, 2'])
  })

  it('rebuilds the table in place for the category it lists', () => {
    insertTableOfAuthorities(editor, [], () => null, heading, 2)
    markCitation(editor, 5, { long: 'Another Act', category: 2 })
    expect(updateTableOfAuthorities(editor, [], () => [1, 3], heading)).toBe(true)
    expect(tableLines()).toEqual([
      ['Heading 2', ''],
      ['Another Act', '1'],
      ['RCW 4.16', '3'],
    ])
  })

  it('says so when a category has no citation, or there is no table to update', () => {
    const alert = vi.spyOn(window, 'alert').mockImplementation(() => {})
    insertTableOfAuthorities(editor, [], () => null, heading, 5)
    expect(updateTableOfAuthorities(editor, [], () => null, heading)).toBe(false)
    expect(alert).toHaveBeenCalledTimes(2)
    expect(tableLines()).toEqual([])
  })

  it('a copied citation keeps its switches, and a copied index entry its term', () => {
    const html = editor.view.serializeForClipboard(editor.state.doc.slice(0)).dom.innerHTML
    const pasted = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: html,
    })
    expect(collectCitations(pasted).map((c) => [c.long, c.category])).toEqual(
      collectCitations(editor).map((c) => [c.long, c.category]),
    )
    pasted.commands.setContent('<p>x<span data-xe-term="widgets"></span></p>')
    let term = ''
    pasted.state.doc.descendants((node) => {
      if (node.type.name === 'docXeMark') term = String(node.attrs.term)
    })
    expect(term).toBe('widgets')
    pasted.destroy()
  })
})
