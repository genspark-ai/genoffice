import { afterEach, beforeAll, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  clearMultiSelect,
  isMultiSelectModifier,
  multiSelectRanges,
} from '../src/renderer/editor/multi-select'

/**
 * Ctrl/⌘+click adds a non-contiguous selection range (genoffice#2013).
 * ProseMirror's Selection holds one {from, to}, so the added ranges live in
 * plugin state and are painted as decorations; the live selection stays the
 * one the core owns.
 *
 * jsdom has no layout, so ProseMirror's hit-test is stubbed with the document
 * position a click would land on; everything from there on — the mouseup-side
 * capture, the word under the cursor, the suppression of ProseMirror's
 * select-clicked-node — is the plugin's real path. The honest end-to-end check
 * runs in e2e/docs-multiselect.spec.ts against a real Chromium.
 */

// A dispatched mousedown walks the editor's whole hit-test chain — the
// float-image probe (caretRangeFromPoint, elementsFromPoint) — and jsdom
// implements none of it (the renderer is Chromium-only). The stubs read as
// "no glyph, no floating picture"; the click's document position is stubbed
// onto the view per test.
beforeAll(() => {
  document.caretRangeFromPoint = () => null
  document.elementsFromPoint = () => []
  document.elementFromPoint = () => null
})

const para = (t: string) => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content: [{ type: 'text', text: t }],
})

const editors = new Set<Editor>()
afterEach(() => {
  for (const e of editors) {
    const host = e.options.element
    e.destroy()
    if (host instanceof HTMLElement) host.remove()
  }
  editors.clear()
})

function makeEditor(text = 'alpha bravo charlie'): Editor {
  // the editor must sit in the document: ProseMirror listens for mouseup on
  // the root, and a detached element never bubbles there — the whole click
  // path would silently not run
  const host = document.createElement('div')
  document.body.appendChild(host)
  const editor = new Editor({
    element: host,
    extensions: editorExtensions,
    content: { type: 'doc', content: [para(text)] },
  })
  editors.add(editor)
  return editor
}

interface ClickAt {
  readonly x: number
  readonly y: number
}

const AT: ClickAt = { x: 0, y: 0 }

function dispatchMouse(editor: Editor, type: 'mousedown' | 'mouseup', init: MouseEventInit): void {
  editor.view.dom.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, ...init }))
}

/**
 * Fire a full ctrl+click whose hit-test lands on `pos` (a document position):
 * mousedown and mouseup both dispatch on the editor DOM, so ProseMirror's
 * mouseup processing — and with it the plugin's handleClick — runs for real.
 * Both modifiers are set because prosemirror-view picks its own
 * select-clicked-node modifier per platform (metaKey on mac, ctrlKey
 * elsewhere): a click must suppress the node selection on either one.
 */
function ctrlClick(editor: Editor, pos: number, at: ClickAt = AT): void {
  // inside mirrors a real click on the paragraph text: the nearest view desc
  // starts at 1 (the paragraph sits at 0), and ProseMirror's selectClickedNode
  // only engages when inside is a node position — with -1 the suppression
  // could not be tested at all
  editor.view.posAtCoords = () => ({ pos, inside: 1 })
  const init = { ctrlKey: true, metaKey: true, clientX: at.x, clientY: at.y }
  dispatchMouse(editor, 'mousedown', init)
  dispatchMouse(editor, 'mouseup', init)
}

describe('isMultiSelectModifier', () => {
  const ev = (init: Partial<MouseEvent>) =>
    ({ metaKey: false, ctrlKey: false, altKey: false, shiftKey: false, ...init }) as MouseEvent

  it("accepts the platform's modifier on its own and ignores the other", () => {
    // jsdom reports no platform, so the plugin takes the non-Mac branch
    expect(isMultiSelectModifier(ev({ ctrlKey: true }))).toBe(true)
    expect(isMultiSelectModifier(ev({ metaKey: true }))).toBe(false)
  })

  it('rejects the combos that already mean something else', () => {
    expect(isMultiSelectModifier(ev({}))).toBe(false)
    expect(isMultiSelectModifier(ev({ ctrlKey: true, shiftKey: true }))).toBe(false)
    expect(isMultiSelectModifier(ev({ ctrlKey: true, altKey: true }))).toBe(false)
  })
})

describe('MultiSelect', () => {
  it('starts with no added ranges', () => {
    expect(multiSelectRanges(makeEditor().view)).toEqual([])
  })

  it('a ctrl+click adds the pre-click range and the word under the cursor', () => {
    const editor = makeEditor()
    editor.commands.setTextSelection({ from: 1, to: 6 }) // "alpha"
    ctrlClick(editor, 9) // inside "bravo" (7..12)
    expect(multiSelectRanges(editor.view)).toEqual([
      { from: 1, to: 6 },
      { from: 7, to: 12 },
    ])
  })

  it('a ctrl+click with nothing selected adds just the word', () => {
    const editor = makeEditor()
    ctrlClick(editor, 9)
    expect(multiSelectRanges(editor.view)).toEqual([{ from: 7, to: 12 }])
  })

  it('makes the clicked word the live selection, never a node selection', () => {
    const editor = makeEditor()
    ctrlClick(editor, 9)
    const sel = editor.state.selection
    // PM's own selectNodeModifier would otherwise turn the paragraph into a
    // NodeSelection, which the next click would store as a "range"
    expect(sel instanceof TextSelection).toBe(true)
    expect([sel.from, sel.to]).toEqual([7, 12])
  })

  it('a second ctrl+click adds a second, disjoint range', () => {
    const editor = makeEditor()
    ctrlClick(editor, 9) // bravo
    ctrlClick(editor, 16) // charlie
    expect(multiSelectRanges(editor.view)).toEqual([
      { from: 7, to: 12 },
      { from: 13, to: 20 },
    ])
    expect(editor.state.selection.empty).toBe(false)
    expect(editor.state.selection.to - editor.state.selection.from).toBe(7)
  })

  it('re-clicking the same word does not stack a duplicate', () => {
    const editor = makeEditor()
    ctrlClick(editor, 9)
    ctrlClick(editor, 10)
    expect(multiSelectRanges(editor.view)).toEqual([{ from: 7, to: 12 }])
  })

  it('a ctrl+click on a link leaves the gesture to the link, adding nothing', () => {
    const editor = makeEditor()
    editor.commands.setTextSelection({ from: 1, to: 6 })
    editor.view.posAtCoords = () => ({ pos: 9, inside: 1 })
    const link = document.createElement('a')
    link.href = 'https://example.com'
    editor.view.dom.appendChild(link)
    const init = { ctrlKey: true, metaKey: true, clientX: 0, clientY: 0 }
    link.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, ...init }))
    link.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, ...init }))
    expect(multiSelectRanges(editor.view)).toEqual([])
  })

  it('a plain click clears every added range', () => {
    const editor = makeEditor()
    ctrlClick(editor, 9)
    ctrlClick(editor, 16)
    expect(multiSelectRanges(editor.view)).toHaveLength(2)
    editor.view.posAtCoords = () => ({ pos: 3, inside: -1 })
    dispatchMouse(editor, 'mousedown', { clientX: 0, clientY: 0 })
    dispatchMouse(editor, 'mouseup', { clientX: 0, clientY: 0 })
    expect(multiSelectRanges(editor.view)).toEqual([])
  })

  it('a ctrl+click on no word keeps the pre-click range and adds nothing else', () => {
    const editor = makeEditor('alpha  bravo charlie') // two spaces
    editor.commands.setTextSelection({ from: 1, to: 6 })
    ctrlClick(editor, 7) // the second space, between no two word characters
    expect(multiSelectRanges(editor.view)).toEqual([{ from: 1, to: 6 }])
  })

  it('a ctrl+drag adds the pre-drag range and the dragged one', () => {
    const editor = makeEditor()
    editor.commands.setTextSelection({ from: 1, to: 6 }) // pre-drag: "alpha"
    editor.view.posAtCoords = () => ({ pos: 9, inside: -1 })
    dispatchMouse(editor, 'mousedown', { ctrlKey: true, clientX: 0, clientY: 0 })
    // what the browser's drag selects while the button is down
    editor.commands.setTextSelection({ from: 7, to: 12 })
    dispatchMouse(editor, 'mouseup', { ctrlKey: true, clientX: 20, clientY: 0 })
    expect(multiSelectRanges(editor.view)).toEqual([
      { from: 1, to: 6 },
      { from: 7, to: 12 },
    ])
    expect(
      editor.state.doc.textBetween(editor.state.selection.from, editor.state.selection.to),
    ).toBe('bravo')
  })

  it('a ctrl+drag released outside the editor still adds both ranges', () => {
    const editor = makeEditor()
    editor.commands.setTextSelection({ from: 1, to: 6 })
    dispatchMouse(editor, 'mousedown', { ctrlKey: true, clientX: 0, clientY: 0 })
    editor.commands.setTextSelection({ from: 7, to: 12 })
    document.body.dispatchEvent(
      new MouseEvent('mouseup', { bubbles: true, ctrlKey: true, clientX: 20, clientY: 0 }),
    )
    expect(multiSelectRanges(editor.view)).toEqual([
      { from: 1, to: 6 },
      { from: 7, to: 12 },
    ])
  })

  it('a ctrl+click that jittered past the drag threshold still selects the word', () => {
    const editor = makeEditor()
    editor.commands.setTextSelection({ from: 1, to: 6 })
    editor.view.posAtCoords = () => ({ pos: 9, inside: -1 })
    dispatchMouse(editor, 'mousedown', { ctrlKey: true, clientX: 0, clientY: 0 })
    // the wobble dragged no text: the browser left the selection collapsed
    editor.commands.setTextSelection({ from: 9, to: 9 })
    dispatchMouse(editor, 'mouseup', { ctrlKey: true, clientX: 6, clientY: 0 })
    expect(multiSelectRanges(editor.view)).toEqual([
      { from: 1, to: 6 },
      { from: 7, to: 12 },
    ])
    expect(
      editor.state.doc.textBetween(editor.state.selection.from, editor.state.selection.to),
    ).toBe('bravo')
  })

  it('an edit before a range moves it along with the text', () => {
    const editor = makeEditor()
    ctrlClick(editor, 9)
    // type "X" at the very start: the range shifts right by one
    editor.commands.insertContentAt(1, 'X')
    expect(multiSelectRanges(editor.view)).toEqual([{ from: 8, to: 13 }])
  })

  it('a range the edit swallowed is dropped, not left as a zero-width sliver', () => {
    const editor = makeEditor()
    ctrlClick(editor, 9)
    // delete exactly the range
    editor.commands.deleteRange({ from: 7, to: 12 })
    expect(multiSelectRanges(editor.view)).toEqual([])
  })

  it('clearMultiSelect drops every added range', () => {
    const editor = makeEditor()
    ctrlClick(editor, 9)
    ctrlClick(editor, 16)
    expect(multiSelectRanges(editor.view)).toHaveLength(2)
    clearMultiSelect(editor.view)
    expect(multiSelectRanges(editor.view)).toEqual([])
  })

  it('clearMultiSelect on an empty state dispatches nothing', () => {
    const editor = makeEditor()
    const before = editor.state
    clearMultiSelect(editor.view)
    expect(editor.state).toBe(before)
  })

  it('paints a decoration over each added range', () => {
    const editor = makeEditor()
    ctrlClick(editor, 9)
    ctrlClick(editor, 16)
    // the plugin's decorations prop paints this class; assert through the DOM
    // that each range carries it and covers exactly the clicked text
    const els = Array.from(editor.view.dom.querySelectorAll('.doc-multi-range'))
    expect(els.map((el) => el.textContent)).toEqual(['bravo', 'charlie'])
  })

  it('leaves the document content itself untouched', () => {
    const editor = makeEditor()
    ctrlClick(editor, 9)
    expect(editor.state.doc.textContent).toBe('alpha bravo charlie')
  })
})
