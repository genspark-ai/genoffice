/** A table of figures lists the captions of one label, and is rebuilt in place. */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { generateCaptionXml, generateTocFieldXml } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { findTocRange } from '../src/renderer/editor/toc-update'
import {
  captionLabels,
  collectCaptions,
  insertTableOfFigures,
  updateTableOfFigures,
} from '../src/renderer/commands/reference-commands'

const caption = (label: string, n: number, text: string) => ({
  type: 'docProtected',
  attrs: { blockType: 'passthrough', genXml: generateCaptionXml(label, n, text) },
})
const paragraph = (text: string) => ({
  type: 'docParagraph',
  content: [{ type: 'text', text }],
})

describe('table of figures', () => {
  let editor: Editor
  beforeEach(() => {
    editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          paragraph('intro'),
          caption('Figure', 1, 'Overview'),
          caption('Table', 1, 'Costs <a&b>'),
          caption('Figure', 2, 'Detail'),
          paragraph('end'),
        ],
      },
    })
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
  })
  afterEach(() => {
    editor.destroy()
    vi.restoreAllMocks()
  })

  /** the generated field lines in the document: what each shows */
  const figureLines = () => {
    const out: Array<{ left: string; right: string }> = []
    editor.state.doc.forEach((node) => {
      if (String(node.attrs.genXml ?? '').includes('TableofFigures'))
        out.push(node.attrs.fieldDisplay as { left: string; right: string })
    })
    return out
  }

  it('finds the captions with their labels and text, in document order', () => {
    expect(collectCaptions(editor, []).map((c) => [c.label, c.text])).toEqual([
      ['Figure', 'Figure 1 Overview'],
      ['Table', 'Table 1 Costs <a&b>'],
      ['Figure', 'Figure 2 Detail'],
    ])
    expect(captionLabels(editor, [])).toEqual(['Figure', 'Table'])
  })

  it('inserts the captions of the chosen label with their pages', () => {
    insertTableOfFigures(editor, [], () => [3, 5], 'Figure')
    expect(figureLines()).toEqual([
      { left: 'Figure 1 Overview', right: '3', kind: 'tocLine', level: 1, leader: 'dot' },
      { left: 'Figure 2 Detail', right: '5', kind: 'tocLine', level: 1, leader: 'dot' },
    ])
  })

  it('takes the first label in use when none is named, and leaves pages blank without a layout', () => {
    insertTableOfFigures(editor, [], () => null)
    expect(figureLines().map((l) => [l.left, l.right])).toEqual([
      ['Figure 1 Overview', ''],
      ['Figure 2 Detail', ''],
    ])
  })

  it('rebuilds the table in place for the label it was made with', () => {
    insertTableOfFigures(editor, [], () => null, 'Table')
    expect(figureLines().map((l) => l.left)).toEqual(['Table 1 Costs <a&b>'])
    editor.commands.insertContentAt(1, caption('Table', 2, 'Later') as never)
    expect(updateTableOfFigures(editor, [], () => [1, 4])).toBe(true)
    expect(figureLines().map((l) => [l.left, l.right])).toEqual([
      ['Table 2 Later', '1'],
      ['Table 1 Costs <a&b>', '4'],
    ])
  })

  it('goes below a selected block instead of replacing it', () => {
    // the first caption, selected as a block, as it is right after Insert Caption
    let captionPos = -1
    editor.state.doc.forEach((node, pos) => {
      if (captionPos < 0 && node.type.name === 'docProtected') captionPos = pos
    })
    editor.commands.setNodeSelection(captionPos)
    insertTableOfFigures(editor, [], () => null, 'Figure')
    expect(collectCaptions(editor, []).map((c) => c.text)).toEqual([
      'Figure 1 Overview',
      'Table 1 Costs <a&b>',
      'Figure 2 Detail',
    ])
    expect(figureLines().map((l) => l.left)).toEqual(['Figure 1 Overview', 'Figure 2 Detail'])
  })

  it('says so when there is nothing to list or to update', () => {
    const alert = vi.spyOn(window, 'alert').mockImplementation(() => {})
    insertTableOfFigures(editor, [], () => null, 'Equation')
    expect(updateTableOfFigures(editor, [], () => null)).toBe(false)
    expect(alert).toHaveBeenCalledTimes(2)
    expect(figureLines()).toEqual([])
  })

  it('is not taken for the table of contents', () => {
    insertTableOfFigures(editor, [], () => null, 'Figure')
    expect(findTocRange(editor.state.doc, [])).toBeNull()
    const toc = generateTocFieldXml([{ level: 1, text: 'Chapter' }]).map((genXml) => ({
      type: 'docProtected',
      attrs: { blockType: 'passthrough', genXml },
    }))
    const end = editor.state.doc.content.size
    editor.commands.insertContentAt(end, toc as never)
    expect(findTocRange(editor.state.doc, [])?.from).toBe(end)
  })
})
