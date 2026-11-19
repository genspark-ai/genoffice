/** Symbol dialog, a symbol font picked: its characters insert as glyphs that save as w:sym. */
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  SYMBOL_FONTS,
  fontSymbolChar,
  fontSymbolChars,
  fontSymbolContent,
  fontSymbolPayload,
  parseFontSymbolPayload,
} from '../src/renderer/editor/font-symbol'
import { runMenuCommand } from '../src/renderer/commands/menu-commands'

describe('symbol font characters', () => {
  let editor: Editor
  beforeEach(() => {
    editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'ab' }] }],
      },
    })
  })
  afterEach(() => editor.destroy())

  it('names a character the way w:sym does and lists the ones that can be drawn', () => {
    expect(fontSymbolChar(0xfc)).toBe('F0FC')
    for (const font of SYMBOL_FONTS) {
      const chars = fontSymbolChars(font)
      expect(chars.length).toBeGreaterThan(50)
      expect(new Set(chars).size).toBe(chars.length)
      expect(chars.every((c) => /^F0[0-9A-F]{2}$/.test(c))).toBe(true)
    }
    // Wingdings FC is the check mark
    expect(fontSymbolContent('Wingdings', 'F0FC')).toEqual({
      type: 'text',
      text: '\u2713',
      marks: [{ type: 'docSym', attrs: { font: 'Wingdings', char: 'F0FC' } }],
    })
    expect(fontSymbolContent('Wingdings', 'zz')).toBeNull()
  })

  it('carries the font and the character through the command payload', () => {
    expect(parseFontSymbolPayload(fontSymbolPayload('Wingdings 2', 'F050'))).toEqual({
      font: 'Wingdings 2',
      char: 'F050',
    })
    expect(parseFontSymbolPayload('Wingdings')).toBeNull()
    expect(parseFontSymbolPayload('\tF050')).toBeNull()
  })

  it('inserts the character at the caret and saves it as a w:sym run', () => {
    editor.commands.setTextSelection(2)
    const host = { editor, hasDoc: true, hfEditor: null } as never
    runMenuCommand(host, 'insert-font-symbol', fontSymbolPayload('Wingdings', 'F0FC'))
    expect(editor.state.doc.firstChild!.textContent).toBe('a\u2713b')
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, [])
    const first = plan.saveBlocks[0]
    const runs = first.kind === 'generated' ? (first.block.runs ?? []) : []
    expect(runs.map((r) => [r.text, r.sym ?? null])).toEqual([
      ['a', null],
      ['\u2713', { font: 'Wingdings', char: 'F0FC' }],
      ['b', null],
    ])
  })
})
