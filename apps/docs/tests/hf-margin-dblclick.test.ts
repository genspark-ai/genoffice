/**
 * Word opens a page's header on a double-click anywhere in its top margin and
 * its footer anywhere in its bottom margin, empty or not. The bands come from
 * the cuts: the gap's bottom part also holds the page's unused body height,
 * which stays the body's, and the gray band between two papers is nobody's.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { hfBandAt, pageMarginBands } from '../src/renderer/editor/hf-margin-hit'
import { applyGapLedger, buildPageCuts, type GapLedgerEntry } from '../src/renderer/pagination-cuts'
import type { BlockBox } from '../src/renderer/pagination-types'

const para = (top: number, height: number): BlockBox => ({
  top,
  height,
  el: document.createElement('p'),
})

// three 400 px pages; each turn: a 96 px bottom margin under 50 px of unused
// body, the 28 px band, then a 120 px top margin (a header pushes the body down)
const entry = (page: number): GapLedgerEntry => ({
  page,
  kind: 'block',
  heightPx: 146 + 28 + 120,
  marginBottomPx: 146,
  marginTopPx: 120,
  marginLeftPx: 90,
  topPushPx: 24,
  pullUpPx: 0,
  repeatHeaderPx: 0,
})

function cutsOf(ledger: GapLedgerEntry[]) {
  const cuts = buildPageCuts(
    [para(0, 400), para(400, 400), para(800, 400)],
    [
      { start: 0, end: 400, section: 0 },
      { start: 400, end: 800, section: 0 },
      { start: 800, end: 1200, section: 0 },
    ],
    [{ contentHeight: 400, forceBreak: false }],
    [{ headerPx: 0, footerPx: 0 }],
  )
  applyGapLedger(cuts, ledger)
  return cuts
}

describe('page margin bands', () => {
  const bands = pageMarginBands(cutsOf([entry(1), entry(2)]), 96, 1980, (p) => (p === 2 ? 110 : 96))

  it('places every page top and bottom margin on the flow', () => {
    expect(bands).toEqual([
      { page: 0, headerTop: 0, headerBottom: 96, footerTop: 546, footerBottom: 642 },
      { page: 1, headerTop: 670, headerBottom: 790, footerTop: 1240, footerBottom: 1336 },
      { page: 2, headerTop: 1364, headerBottom: 1484, footerTop: 1870, footerBottom: 1980 },
    ])
  })

  it('resolves a Y to the header or footer of its page', () => {
    expect(hfBandAt(bands, 10)).toEqual({ kind: 'header', page: 0 })
    expect(hfBandAt(bands, 600)).toEqual({ kind: 'footer', page: 0 })
    expect(hfBandAt(bands, 700)).toEqual({ kind: 'header', page: 1 })
    expect(hfBandAt(bands, 1300)).toEqual({ kind: 'footer', page: 1 })
    expect(hfBandAt(bands, 1975)).toEqual({ kind: 'footer', page: 2 })
  })

  it('leaves the body, its unused height and the band between papers alone', () => {
    expect(hfBandAt(bands, 96)).toBeNull()
    expect(hfBandAt(bands, 300)).toBeNull()
    // unused body height above page 1's bottom margin
    expect(hfBandAt(bands, 520)).toBeNull()
    // the gray band between the papers
    expect(hfBandAt(bands, 655)).toBeNull()
    expect(hfBandAt(bands, 1860)).toBeNull()
  })

  it('keeps one band per page when a page has no gap of its own', () => {
    const shared = pageMarginBands(cutsOf([entry(2)]), 96, 1980, () => 96)
    expect(shared.map((b) => b.page)).toEqual([0, 1, 2])
    expect(hfBandAt(shared, 10)).toEqual({ kind: 'header', page: 0 })
    expect(hfBandAt(shared, 1975)).toEqual({ kind: 'footer', page: 2 })
  })
})

describe('a margin double-click selects no body word', () => {
  let editor: Editor | null = null
  afterEach(() => {
    editor?.destroy()
    editor = null
  })

  function doubleClickAt(claimed: boolean | null) {
    const element = document.createElement('div')
    document.body.appendChild(element)
    editor = new Editor({
      element,
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'Bagian satu' }] }],
      },
    })
    editor.storage.doubleClickWord.claimed = claimed === null ? null : () => claimed
    editor.commands.setTextSelection(1)
    const view = editor.view
    const event = new MouseEvent('mousedown', { detail: 2, button: 0, clientX: 0 })
    const handled = view.someProp('handleDoubleClick', (f) => f(view, 3, event))
    const { from, to } = editor.state.selection
    return { handled, selected: editor.state.doc.textBetween(from, to) }
  }

  it('selects the word under an ordinary double-click', () => {
    expect(doubleClickAt(null)).toEqual({ handled: true, selected: 'Bagian ' })
    expect(doubleClickAt(false).selected).toBe('Bagian ')
  })

  it('keeps the selection when the host claims the double-click', () => {
    expect(doubleClickAt(true)).toEqual({ handled: true, selected: '' })
  })
})
