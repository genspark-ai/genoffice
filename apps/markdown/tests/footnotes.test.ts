/**
 * GFM footnotes. The interesting failures are the same ones wikilinks and the
 * source splice guard against: a definition that lexes as something else (a
 * link reference definition, for single-word bodies), a save that reflows a
 * file the user did not edit, and an editor that traps the caret inside a
 * definition's last empty paragraph.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'
import {
  footnoteDefStart,
  footnoteRefStart,
  nextFootnoteLabel,
} from '../src/renderer/editor/footnote'
import { buildSourceMap, spliceMarkdown } from '../src/renderer/markdown/sourceSplice'

const editors: Editor[] = []
afterAll(() => {
  for (const e of editors) e.destroy()
})

function createEditor(): Editor {
  const editor = new Editor({
    extensions: buildExtensions({
      slashController: {
        onOpen: () => {},
        onUpdate: () => {},
        onKeyDown: () => false,
        onClose: () => {},
      },
      slashItems: () => [],
    }),
    content: '',
  })
  editors.push(editor)
  return editor
}

/** parse → serialize → parse must be structurally stable */
function roundTrip(editor: Editor, md: string): { out: string; stable: boolean } {
  const manager = editor.markdown!
  const first = manager.parse(md)
  const out = manager.serialize(first)
  const second = manager.parse(out)
  return { out, stable: JSON.stringify(first) === JSON.stringify(second) }
}

function load(editor: Editor, source: string) {
  editor.commands.setContent(source, { contentType: 'markdown' })
  return buildSourceMap(editor, editor.state.doc, source)
}

function save(editor: Editor, source: string): string {
  const map = load(editor, source)
  expect(map).not.toBeNull()
  return spliceMarkdown(editor, editor.state.doc, map!)
}

function blockTypes(editor: Editor): string[] {
  const out: string[] = []
  editor.state.doc.forEach((node) => out.push(node.type.name))
  // the editor keeps a trailing empty paragraph as the caret's rest point
  const last = editor.state.doc.lastChild
  if (last?.type.name === 'paragraph' && last.childCount === 0) out.pop()
  return out
}

function nodesOf(editor: Editor, type: string): { label: string; text: string }[] {
  const out: { label: string; text: string }[] = []
  editor.state.doc.descendants((node) => {
    if (node.type.name === type) {
      out.push({ label: String(node.attrs.label ?? ''), text: node.textContent })
    }
  })
  return out
}

describe('footnote parsing', () => {
  const editor = createEditor()

  it('a reference and its definition become their nodes', () => {
    editor.commands.setContent('See the note[^1] here.\n\n[^1]: The first footnote.', {
      contentType: 'markdown',
    })
    expect(blockTypes(editor)).toEqual(['paragraph', 'footnoteDefinition'])
    expect(nodesOf(editor, 'footnoteReference')).toEqual([{ label: '1', text: '' }])
    expect(nodesOf(editor, 'footnoteDefinition')).toEqual([
      { label: '1', text: 'The first footnote.' },
    ])
  })

  it('a paragraph is cut before a definition on the next line', () => {
    editor.commands.setContent('Some text\n[^1]: note', { contentType: 'markdown' })
    expect(blockTypes(editor)).toEqual(['paragraph', 'footnoteDefinition'])
  })

  it('a single-word body is not a link reference definition', () => {
    editor.commands.setContent('See[^1].\n\n[^1]: https://example.com', {
      contentType: 'markdown',
    })
    expect(blockTypes(editor)).toEqual(['paragraph', 'footnoteDefinition'])
    expect(nodesOf(editor, 'footnoteDefinition')[0]?.text).toBe('https://example.com')
  })

  it('consecutive definitions without a blank line stay separate blocks', () => {
    editor.commands.setContent('[^1]: one\n[^2]: two', { contentType: 'markdown' })
    expect(blockTypes(editor)).toEqual(['footnoteDefinition', 'footnoteDefinition'])
    expect(nodesOf(editor, 'footnoteDefinition').map((n) => n.label)).toEqual(['1', '2'])
  })

  it('an indented continuation line joins the body', () => {
    editor.commands.setContent('[^1]: First\n    continued.', { contentType: 'markdown' })
    const defs = nodesOf(editor, 'footnoteDefinition')
    expect(defs).toEqual([{ label: '1', text: 'First\ncontinued.' }])
  })

  it('a blank line plus indent opens a second paragraph inside the body', () => {
    editor.commands.setContent('[^1]: First\n\n    Second paragraph.', {
      contentType: 'markdown',
    })
    expect(blockTypes(editor)).toEqual(['footnoteDefinition'])
    const def = editor.state.doc.firstChild!
    const paragraphs: string[] = []
    def.forEach((child) => paragraphs.push(child.textContent))
    expect(paragraphs).toEqual(['First', 'Second paragraph.'])
  })

  it('a lazy continuation line joins the body', () => {
    editor.commands.setContent('[^1]: First\nmore words.', { contentType: 'markdown' })
    expect(nodesOf(editor, 'footnoteDefinition')).toEqual([
      { label: '1', text: 'First\nmore words.' },
    ])
  })

  it('a non-indented line after a blank line stays outside the definition', () => {
    editor.commands.setContent('[^1]: First\n\nAfter.', { contentType: 'markdown' })
    expect(blockTypes(editor)).toEqual(['footnoteDefinition', 'paragraph'])
  })

  it('a heading right below a definition is not swallowed', () => {
    editor.commands.setContent('[^1]: x\n# Heading', { contentType: 'markdown' })
    expect(blockTypes(editor)).toEqual(['footnoteDefinition', 'heading'])
  })
})

describe('footnote round-trip', () => {
  const editor = createEditor()

  it('references and definitions serialize back in their own form', () => {
    const { out, stable } = roundTrip(editor, 'See the note[^1] here.\n\n[^1]: The first footnote.')
    expect(out).toContain('See the note[^1] here.')
    expect(out).toContain('[^1]: The first footnote.')
    expect(stable).toBe(true)
  })

  it('indented continuation lines come back indented', () => {
    const { out, stable } = roundTrip(editor, '[^1]: First\n    continued.')
    expect(out).toBe('[^1]: First\n    continued.')
    expect(stable).toBe(true)
  })

  it('a second body paragraph comes back blank-line separated and indented', () => {
    const { out, stable } = roundTrip(editor, '[^1]: First\n\n    Second paragraph.')
    expect(out).toBe('[^1]: First\n\n    Second paragraph.')
    expect(stable).toBe(true)
  })

  it('a reference without a definition survives as a reference', () => {
    const { out, stable } = roundTrip(editor, 'See[^1].')
    expect(out).toBe('See[^1].')
    expect(stable).toBe(true)
  })

  it('an unreferenced definition survives', () => {
    const { out, stable } = roundTrip(editor, 'Text.\n\n[^orphan]: alone')
    expect(out).toContain('Text.')
    expect(out).toContain('[^orphan]: alone')
    expect(stable).toBe(true)
  })

  it('an escaped definition stays escaped text', () => {
    const { out, stable } = roundTrip(editor, '\\[^1]: x')
    // the serializer escapes both brackets when a line reads as a definition
    expect(out).toBe('\\[^1\\]: x')
    expect(stable).toBe(true)
  })

  it('a reference inside inline code is not a reference', () => {
    const { out, stable } = roundTrip(editor, 'Use `[^1]` here.')
    expect(out).toBe('Use `[^1]` here.')
    expect(nodesOf(editor, 'footnoteReference')).toEqual([])
    expect(stable).toBe(true)
  })

  it('a definition inside a fenced block is not a definition', () => {
    const { out, stable } = roundTrip(editor, '```\n[^1]: not a note\n```')
    expect(out).toContain('```')
    expect(nodesOf(editor, 'footnoteDefinition')).toEqual([])
    expect(stable).toBe(true)
  })

  it('definitions in the middle of a document round-trip', () => {
    const md = 'Top[^a].\n\n[^a]: early note\n\nMiddle.\n\n[^b]: late note\n\nBottom[^b].'
    const { stable } = roundTrip(editor, md)
    expect(stable).toBe(true)
  })
})

describe('footnote probes', () => {
  it('footnoteDefStart finds line-start definitions only', () => {
    expect(footnoteDefStart('Some text\n[^1]: x')).toBe(10)
    expect(footnoteDefStart('a [^1]: x')).toBe(-1)
    expect(footnoteDefStart('no markers')).toBe(-1)
    expect(footnoteDefStart('text\nnot one\n[^2]: y')).toBe(13)
  })

  it('footnoteRefStart finds references but not escaped ones', () => {
    expect(footnoteRefStart('see [^1] here')).toBe(4)
    expect(footnoteRefStart('an \\[^1] escape')).toBe(-1)
    expect(footnoteRefStart('plain text')).toBe(-1)
  })
})

describe('footnotes through the block source map', () => {
  const editor = createEditor()

  const SOURCE =
    'Intro with a note[^1] and a named one[^labeled].\n\n' +
    'Body paragraph stays.\n\n' +
    '[^1]: The first note.\n' +
    '    wrapped onto a second line.\n\n' +
    '[^labeled]: Another note\n\n' +
    '    Its own second paragraph.\n\n' +
    'Closing text[^1].\n'

  it('an untouched footnote file saves byte-identical', () => {
    expect(save(editor, SOURCE)).toBe(SOURCE)
  })

  it('editing another block keeps every definition verbatim', () => {
    const map = load(editor, SOURCE)
    expect(map).not.toBeNull()
    let at = -1
    editor.state.doc.descendants((node, pos) => {
      if (at < 0 && node.isText && node.text?.startsWith('Body paragraph')) at = pos
    })
    expect(at).toBeGreaterThan(0)
    editor.view.dispatch(editor.state.tr.insertText('Edited', at, at + 4))
    const saved = spliceMarkdown(editor, editor.state.doc, map!)
    expect(saved).toContain('Edited paragraph stays.')
    expect(saved).toContain('[^1]: The first note.\n    wrapped onto a second line.')
    expect(saved).toContain('[^labeled]: Another note\n\n    Its own second paragraph.')
    expect(saved.startsWith('Intro with a note[^1] and a named one[^labeled].')).toBe(true)
  })

  it('editing a definition re-serializes it in place', () => {
    const map = load(editor, SOURCE)
    expect(map).not.toBeNull()
    let at = -1
    editor.state.doc.descendants((node, pos) => {
      if (at < 0 && node.isText && node.text?.startsWith('The first note')) at = pos
    })
    expect(at).toBeGreaterThan(0)
    editor.view.dispatch(editor.state.tr.insertText('Changed', at, at + 14))
    const saved = spliceMarkdown(editor, editor.state.doc, map!)
    expect(saved).toContain('[^1]: Changed.\n    wrapped onto a second line.')
    expect(saved).toContain('Body paragraph stays.')
    expect(saved).toContain('[^labeled]: Another note\n\n    Its own second paragraph.')
  })
})

describe('footnote editing', () => {
  /** a real keydown, not the keyboardShortcut command: that command replays the
   * captured steps through a second mapping, which relocates the second step of
   * a multi-step transaction (the live key path dispatches directly) */
  function pressEnter(editor: Editor): void {
    editor.view.dom.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }),
    )
  }

  it('Enter in the only, empty paragraph of a definition replaces the definition', () => {
    const editor = createEditor()
    editor.commands.setContent({
      type: 'doc',
      content: [
        {
          type: 'footnoteDefinition',
          attrs: { label: '1' },
          content: [{ type: 'paragraph' }],
        },
      ],
    })
    editor.commands.setTextSelection(2) // inside the definition's empty paragraph
    pressEnter(editor)
    expect(blockTypes(editor)).toEqual(['paragraph'])
  })

  it('Enter in the trailing empty paragraph leaves the definition', () => {
    const editor = createEditor()
    editor.commands.setContent({
      type: 'doc',
      content: [
        {
          type: 'footnoteDefinition',
          attrs: { label: '1' },
          content: [
            { type: 'paragraph', content: [{ type: 'text', text: 'First' }] },
            { type: 'paragraph' },
          ],
        },
      ],
    })
    editor.commands.setTextSelection(9) // inside the trailing empty paragraph
    pressEnter(editor)
    expect(blockTypes(editor)).toEqual(['footnoteDefinition', 'paragraph'])
    const def = nodesOf(editor, 'footnoteDefinition')
    expect(def).toEqual([{ label: '1', text: 'First' }])
  })

  it('nextFootnoteLabel continues the highest numeric label', () => {
    const editor = createEditor()
    editor.commands.setContent('A[^3] and B[^note].\n\n[^3]: third\n\n[^note]: named', {
      contentType: 'markdown',
    })
    expect(nextFootnoteLabel(editor.state.doc)).toBe('4')
    editor.commands.setContent('No footnotes here.')
    expect(nextFootnoteLabel(editor.state.doc)).toBe('1')
  })
})
