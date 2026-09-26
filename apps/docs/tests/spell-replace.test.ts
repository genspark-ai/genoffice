/**
 * Right-click spelling suggestions go through Blink's replaceMisspelling
 * first and fall back to a ProseMirror replacement: the misspelled word
 * Chromium reports is located in the clicked textblock at word boundaries
 * (the occurrence under the click when it repeats), and the replacement keeps
 * the run's marks.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { applySpellingSuggestion, misspelledRangeAt } from '../src/renderer/editor/spell-replace'

const liveEditors: Editor[] = []
afterEach(() => {
  for (const editor of liveEditors.splice(0)) editor.destroy()
  vi.useRealTimers()
})

function makeEditor(text: string, marks?: Array<{ type: string }>) {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        { type: 'docParagraph', content: [{ type: 'text', text, ...(marks ? { marks } : {}) }] },
      ],
    },
  })
  liveEditors.push(editor)
  return editor
}

describe('misspelledRangeAt', () => {
  it('finds the word under the click at word boundaries', () => {
    const editor = makeEditor('Je vais a la mison ce soir')
    // paragraph content starts at 1; "mison" is at offset 13
    const range = misspelledRangeAt(editor.state.doc, 1 + 15, 'mison')
    expect(range).toEqual({ from: 14, to: 19 })
    expect(editor.state.doc.textBetween(range!.from, range!.to)).toBe('mison')
  })

  it('picks the occurrence containing the click when the word repeats', () => {
    const editor = makeEditor('mison encore la mison demain')
    expect(misspelledRangeAt(editor.state.doc, 1 + 2, 'mison')).toEqual({ from: 1, to: 6 })
    expect(misspelledRangeAt(editor.state.doc, 1 + 18, 'mison')).toEqual({ from: 17, to: 22 })
  })

  it('rejects substrings inside longer words and clicks away from the word', () => {
    const editor = makeEditor('the misonry wall mison')
    // "misonry" contains "mison" but is not a whole word
    expect(misspelledRangeAt(editor.state.doc, 1 + 6, 'mison')).toBeNull()
    // click on "wall": not on the misspelled word
    expect(misspelledRangeAt(editor.state.doc, 1 + 14, 'mison')).toBeNull()
  })

  it('replacing through insertText keeps the run marks', () => {
    const editor = makeEditor('a mison b', [{ type: 'bold' }])
    const range = misspelledRangeAt(editor.state.doc, 1 + 4, 'mison')!
    editor.commands.command(({ tr }) => {
      tr.insertText('maison', range.from, range.to)
      return true
    })
    expect(editor.state.doc.textContent).toBe('a maison b')
    const run = editor.state.doc.firstChild!.firstChild!
    expect(run.text).toBe('a maison b')
    expect(run.marks.some((m) => m.type.name === 'bold')).toBe(true)
  })
})

describe('applySpellingSuggestion', () => {
  const settle = async () => {
    await vi.advanceTimersByTimeAsync(1000)
  }

  it('replaces through ProseMirror when Blink is unavailable', () => {
    const editor = makeEditor('a mison b')
    applySpellingSuggestion(editor, 1 + 4, 'mison', 'maison')
    expect(editor.state.doc.textContent).toBe('a maison b')
  })

  /** Blink replaces the misspelling under the selection and the edit reaches ProseMirror as a DOM change */
  const blinkStandIn = (editor: Editor) =>
    vi.fn(async (word: string) => {
      const { from, to } = editor.state.selection
      editor.commands.command(({ tr }) => {
        tr.insertText(word, from, to)
        return true
      })
    })

  it('leaves the document to Blink when its replacement lands', async () => {
    vi.useFakeTimers()
    const editor = makeEditor('a mison b')
    const blink = blinkStandIn(editor)
    applySpellingSuggestion(editor, 1 + 4, 'mison', 'maison', blink)
    await settle()
    expect(blink).toHaveBeenCalledWith('maison')
    expect(editor.state.doc.textContent).toBe('a maison b')
  })

  it('aims Blink at the clicked word when the right-click landed inside a selection', async () => {
    vi.useFakeTimers()
    const editor = makeEditor('mison encore la mison demain')
    editor.commands.setTextSelection({ from: 1, to: 22 })
    applySpellingSuggestion(editor, 1 + 18, 'mison', 'maison', blinkStandIn(editor))
    await settle()
    expect(editor.state.doc.textContent).toBe('mison encore la maison demain')
  })

  it('does not call Blink over a selection when the clicked word cannot be located', async () => {
    vi.useFakeTimers()
    const editor = makeEditor('mison encore la mison demain')
    editor.commands.setTextSelection({ from: 1, to: 22 })
    const blink = blinkStandIn(editor)
    applySpellingSuggestion(editor, 1 + 8, 'mison', 'maison', blink)
    await settle()
    expect(blink).not.toHaveBeenCalled()
    expect(editor.state.doc.textContent).toBe('mison encore la mison demain')
  })

  it('falls back to ProseMirror when Blink does not act', async () => {
    vi.useFakeTimers()
    const editor = makeEditor('a mison b')
    applySpellingSuggestion(editor, 1 + 4, 'mison', 'maison', async () => {})
    expect(editor.state.doc.textContent).toBe('a mison b')
    await settle()
    expect(editor.state.doc.textContent).toBe('a maison b')
  })

  it('falls back to ProseMirror when the Blink call fails', async () => {
    vi.useFakeTimers()
    const editor = makeEditor('a mison b')
    applySpellingSuggestion(editor, 1 + 4, 'mison', 'maison', () => Promise.reject(new Error('x')))
    await settle()
    expect(editor.state.doc.textContent).toBe('a maison b')
  })

  it('does nothing when the click position is unknown and Blink does not act', async () => {
    vi.useFakeTimers()
    const editor = makeEditor('a mison b')
    applySpellingSuggestion(editor, null, 'mison', 'maison', async () => {})
    await settle()
    expect(editor.state.doc.textContent).toBe('a mison b')
  })
})
