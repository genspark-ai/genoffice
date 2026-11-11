/**
 * Block ids the Obsidian way: `^id` is plain text at a line's end (or alone on
 * the line after a table or code block), the copied `[[Note#^id]]` resolves
 * through the same fragment lookup other tools' ids already take, and nothing
 * is written when the block carries an id already.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { blockIds, noteNameOf } from '../src/renderer/editor/blockId'
import { wikiFragmentPosition } from '../src/renderer/editor/wikilink'

const editors: Editor[] = []
afterAll(() => {
  for (const e of editors) e.destroy()
})

interface Harness {
  editor: Editor
  copied: string[]
  nothing: number
}

function openDoc(source: string, path = '/vault/Note.md'): Harness {
  const copied: string[] = []
  const state = { nothing: 0 }
  const editor = new Editor({
    extensions: buildExtensions({
      slashController: {
        onOpen: () => {},
        onUpdate: () => {},
        onKeyDown: () => false,
        onClose: () => {},
      },
      slashItems: () => [],
      blockLink: {
        getDocumentPath: () => path,
        write: (text) => {
          copied.push(text)
        },
        onNothing: () => {
          state.nothing++
        },
      },
    }),
    content: '',
  })
  editors.push(editor)
  editor.commands.setContent(source, { contentType: 'markdown' })
  return {
    editor,
    copied,
    get nothing() {
      return state.nothing
    },
  }
}

/** a position inside the first textblock whose text contains `text` */
function caretIn(editor: Editor, text: string): void {
  let at: number | null = null
  editor.state.doc.descendants((node, pos) => {
    if (at !== null) return false
    if (node.isTextblock && node.textContent.includes(text)) at = pos + 1
    return at === null
  })
  if (at === null) throw new Error(`no block containing ${text}`)
  editor.commands.setTextSelection(at)
}

const ID_IN_LINK = /\[\[(?:Note)?#\^([a-z0-9]{6})\]\]$/

function idOf(link: string): string {
  const match = ID_IN_LINK.exec(link)
  if (!match) throw new Error(`not a block link: ${link}`)
  return match[1]!
}

describe('copy link to block', () => {
  it('writes ` ^id` at the end of the paragraph and copies [[Note#^id]]', () => {
    const h = openDoc('First.\n\nSecond line here.\n\nThird.')
    caretIn(h.editor, 'Second')
    expect(h.editor.commands.copyBlockLink()).toBe(true)
    expect(h.copied).toHaveLength(1)
    const id = idOf(h.copied[0]!)
    expect(h.copied[0]).toBe(`[[Note#^${id}]]`)
    const md = h.editor.getMarkdown()
    expect(md).toContain(`Second line here. ^${id}`)
    // the other blocks are untouched
    expect(md).toContain('First.\n\nSecond')
    expect(md).toContain('\n\nThird.')
    // the link resolves to the block it was copied from
    const pos = wikiFragmentPosition(h.editor.state.doc, `^${id}`)
    expect(pos).not.toBeNull()
    expect(h.editor.state.doc.resolve(pos!).parent.textContent).toContain('Second line')
  })

  it('reuses the id a block already carries instead of adding a second one', () => {
    const h = openDoc('Line one.\n\nLine two.')
    caretIn(h.editor, 'Line two')
    h.editor.commands.copyBlockLink()
    const before = h.editor.getMarkdown()
    h.editor.commands.copyBlockLink()
    expect(h.copied).toHaveLength(2)
    expect(h.copied[1]).toBe(h.copied[0])
    expect(h.editor.getMarkdown()).toBe(before)
  })

  it('addresses an id another tool wrote, changing nothing', () => {
    const h = openDoc('Quoted line. ^quote-1\n\nNext.')
    const before = h.editor.getMarkdown()
    caretIn(h.editor, 'Quoted')
    h.editor.commands.copyBlockLink()
    expect(h.copied).toEqual(['[[Note#^quote-1]]'])
    expect(h.editor.getMarkdown()).toBe(before)
  })

  it('puts a list item’s id on the item’s own line', () => {
    const h = openDoc('- one\n- two\n- three')
    caretIn(h.editor, 'two')
    h.editor.commands.copyBlockLink()
    const id = idOf(h.copied[0]!)
    const md = h.editor.getMarkdown()
    expect(md).toContain(`- two ^${id}`)
    expect(md).toContain('- one\n')
    expect(md).toContain('- three')
    const pos = wikiFragmentPosition(h.editor.state.doc, `^${id}`)
    expect(h.editor.state.doc.resolve(pos!).parent.textContent).toBe(`two ^${id}`)
  })

  it('gives a table a lone ^id line after it, and the link lands in the table', () => {
    const h = openDoc('Intro.\n\n| A | B |\n| - | - |\n| a1 | b1 |\n\nOutro.')
    caretIn(h.editor, 'a1')
    h.editor.commands.copyBlockLink()
    const id = idOf(h.copied[0]!)
    const md = h.editor.getMarkdown()
    // the marker is its own paragraph between the table and what follows
    expect(md).toMatch(new RegExp(`\\|\\s*\\n\\n\\^${id}\\n\\nOutro\\.`))
    // no cell was written into
    expect(md).not.toContain(`a1 ^${id}`)
    const pos = wikiFragmentPosition(h.editor.state.doc, `^${id}`)
    expect(pos).not.toBeNull()
    const $pos = h.editor.state.doc.resolve(pos!)
    const ancestors = Array.from({ length: $pos.depth }, (_, i) => $pos.node(i + 1).type.name)
    expect(ancestors).toContain('table')
    // asking again finds the lone line and writes nothing more
    h.editor.commands.copyBlockLink()
    expect(h.copied[1]).toBe(h.copied[0])
    expect(h.editor.getMarkdown()).toBe(md)
  })

  it('gives a code block a lone ^id line after it, never a line inside it', () => {
    const h = openDoc('```js\nconsole.log(1)\n```\n\nAfter.')
    caretIn(h.editor, 'console')
    h.editor.commands.copyBlockLink()
    const id = idOf(h.copied[0]!)
    const md = h.editor.getMarkdown()
    expect(md).toContain('console.log(1)\n```\n\n^' + id + '\n\nAfter.')
    const pos = wikiFragmentPosition(h.editor.state.doc, `^${id}`)
    expect(h.editor.state.doc.resolve(pos!).parent.type.name).toBe('codeBlock')
  })

  it('gives a quotation a lone ^id line after it, as Obsidian documents for quotes', () => {
    const h = openDoc('> Quoted words.\n> More of them.\n\nAfter.')
    caretIn(h.editor, 'More of them')
    h.editor.commands.copyBlockLink()
    const id = idOf(h.copied[0]!)
    const md = h.editor.getMarkdown()
    expect(md).toContain(`More of them.\n\n^${id}\n\nAfter.`)
    expect(md).not.toContain(`them. ^${id}`)
    const pos = wikiFragmentPosition(h.editor.state.doc, `^${id}`)
    const $pos = h.editor.state.doc.resolve(pos!)
    const ancestors = Array.from({ length: $pos.depth }, (_, i) => $pos.node(i + 1).type.name)
    expect(ancestors).toContain('blockquote')
  })

  it('gives a selected block atom a lone ^id line after it', () => {
    const h = openDoc('Before.\n\n$$\nx^2\n$$\n\nAfter.')
    let mathPos = -1
    h.editor.state.doc.descendants((node, pos) => {
      if (node.type.name === 'blockMath') mathPos = pos
    })
    expect(mathPos).toBeGreaterThan(0)
    h.editor.view.dispatch(
      h.editor.state.tr.setSelection(NodeSelection.create(h.editor.state.doc, mathPos)),
    )
    h.editor.commands.copyBlockLink()
    const id = idOf(h.copied[0]!)
    const md = h.editor.getMarkdown()
    expect(md).toMatch(new RegExp(`\\$\\$\\s*\\n\\n\\^${id}\\n\\nAfter\\.`))
  })

  it('links a heading by its text, writing no id', () => {
    const h = openDoc('# Getting Started\n\nBody.')
    const before = h.editor.getMarkdown()
    caretIn(h.editor, 'Getting')
    h.editor.commands.copyBlockLink()
    expect(h.copied).toEqual(['[[Note#Getting Started]]'])
    expect(h.editor.getMarkdown()).toBe(before)
    expect(wikiFragmentPosition(h.editor.state.doc, 'Getting Started')).not.toBeNull()
  })

  it('addresses the note itself while the document is untitled', () => {
    const h = openDoc('Unsaved line.', '')
    caretIn(h.editor, 'Unsaved')
    h.editor.commands.copyBlockLink()
    const id = idOf(h.copied[0]!)
    expect(h.copied[0]).toBe(`[[#^${id}]]`)
  })

  it('does nothing on an empty line and says so', () => {
    const h = openDoc('Text.\n\nMore.')
    // an empty paragraph between the two: nothing to address
    const afterFirst = h.editor.state.doc.child(0).nodeSize
    h.editor.commands.insertContentAt(afterFirst, { type: 'paragraph' })
    let empty: number | null = null
    h.editor.state.doc.descendants((node, pos) => {
      if (empty === null && node.type.name === 'paragraph' && node.childCount === 0) empty = pos + 1
    })
    h.editor.commands.setTextSelection(empty!)
    const before = h.editor.getMarkdown()
    expect(h.editor.commands.copyBlockLink()).toBe(false)
    expect(h.copied).toEqual([])
    expect(h.nothing).toBe(1)
    expect(h.editor.getMarkdown()).toBe(before)
  })

  it('never inherits the marks at the line’s end', () => {
    const h = openDoc('Run `npm test` now `again`')
    caretIn(h.editor, 'Run')
    h.editor.commands.copyBlockLink()
    const id = idOf(h.copied[0]!)
    // inside the code span the id would have been code, not an address
    expect(h.editor.getMarkdown()).toContain('`again` ^' + id)
  })

  it('draws a fresh id that no block in the document already uses', () => {
    const h = openDoc('a ^abc123\n\nb ^def456\n\nthird block')
    expect([...blockIds(h.editor.state.doc)].sort()).toEqual(['abc123', 'def456'])
    caretIn(h.editor, 'third')
    h.editor.commands.copyBlockLink()
    const id = idOf(h.copied[0]!)
    expect(['abc123', 'def456']).not.toContain(id)
  })

  it('names the note by its stem', () => {
    expect(noteNameOf('/vault/folder/Daily Note.md')).toBe('Daily Note')
    expect(noteNameOf('C:\\notes\\Plan.markdown')).toBe('Plan')
    expect(noteNameOf('')).toBe('')
  })
})

describe('block id markers in the view', () => {
  const markers = (editor: Editor) =>
    [...editor.view.dom.querySelectorAll('.md-block-id')].map((el) => el.textContent)

  it('dims a trailing marker and a lone one', () => {
    const h = openDoc('Prose here. ^abc123\n\n| A |\n| - |\n| 1 |\n\n^tbl-id\n\nPlain.')
    expect(markers(h.editor)).toEqual(['^abc123', '^tbl-id'])
  })

  it('leaves code and a marker split by a mark as they read', () => {
    const h = openDoc('```\nx ^abc123\n```\n\nSplit ^ab**c1**')
    expect(markers(h.editor)).toEqual([])
  })

  it('follows an edit', () => {
    const h = openDoc('Line.')
    expect(markers(h.editor)).toEqual([])
    caretIn(h.editor, 'Line')
    h.editor.commands.copyBlockLink()
    expect(markers(h.editor)).toEqual([`^${idOf(h.copied[0]!)}`])
  })
})

describe('lone marker resolution', () => {
  it('a lone marker names the block before it', () => {
    const h = openDoc('```\ncode\n```\n\n^code-1\n\nAfter.')
    const pos = wikiFragmentPosition(h.editor.state.doc, '^code-1')
    expect(h.editor.state.doc.resolve(pos!).parent.type.name).toBe('codeBlock')
  })

  it('a lone marker with nothing before it names its own line', () => {
    const h = openDoc('^first\n\nAfter.')
    const pos = wikiFragmentPosition(h.editor.state.doc, '^first')
    expect(h.editor.state.doc.resolve(pos!).parent.textContent).toBe('^first')
  })
})
