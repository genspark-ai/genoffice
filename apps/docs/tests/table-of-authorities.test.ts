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

  it('puts the number of a sequence before each page (\\s), joined by \\d, and the pages apart by \\l', () => {
    const chapter = (n: number) => ({
      type: 'docProtected',
      attrs: {
        blockType: 'passthrough',
        genXml:
          '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> SEQ Chapter </w:instrText></w:r>' +
          `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${n}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`,
      },
    })
    // chapter 1 opens before the second paragraph, chapter 2 before the third
    editor.commands.insertContentAt(editor.state.doc.child(0).nodeSize, chapter(1) as never)
    const third =
      editor.state.doc.child(0).nodeSize +
      editor.state.doc.child(1).nodeSize +
      editor.state.doc.child(2).nodeSize
    editor.commands.insertContentAt(third, chapter(2) as never)
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    const pagesOf = (positions: number[]) => positions.map((_, i) => i + 4)
    insertTableOfAuthorities(editor, [], pagesOf, heading, undefined, {
      instr: ' TOA \\h \\s Chapter \\d ":" \\l "; " ',
      leader: 'dot',
    })
    expect(tableLines()).toEqual([
      ['Heading 1', ''],
      // the first citation stands ahead of the first chapter number
      ['Baldwin v. Alberti', '4; 2:5'],
      ['Heading 2', ''],
      ['RCW 4.16', '1:4'],
    ])
    // without \\d the two are joined by a hyphen
    expect(updateTableOfAuthorities(editor, [], pagesOf, heading)).toBe(true)
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    insertTableOfAuthorities(editor, [], pagesOf, heading, undefined, {
      instr: ' TOA \\c "2" \\s chapter ',
      leader: 'dot',
    })
    expect(tableLines()).toContainEqual(['RCW 4.16', '1-4']) // a sequence the document does not have numbers nothing: the pages are in page order
    insertTableOfAuthorities(editor, [], () => [9, 5], heading, undefined, {
      instr: ' TOA \\c "1" \\s Part ',
      leader: 'dot',
    })
    expect(tableLines()).toContainEqual(['Baldwin v. Alberti', '5, 9'])
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
