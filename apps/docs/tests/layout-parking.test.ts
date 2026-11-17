/**
 * Layout parking (route-B prototype, #2781): the gate that keeps whole-
 * document style recalc/layout/paint off the typing path. The decisions
 * must be exact from both sides — parking a visible block hides live text,
 * failing to park an offscreen block keeps the O(document) cost the
 * prototype exists to remove. The decoration side is pinned through a real
 * editor: parked entries render as node decorations carrying the
 * remembered height, they map across transactions, and the __parking
 * update meta replaces the set wholesale.
 */
import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { editorExtensions } from '../src/renderer/editor/extensions'
import * as mod from '../src/renderer/editor/layout-parking'
import {
  parkingEntriesFor,
  parkingKey,
  parkBand,
  parkDecisions,
  type ParkEntry,
} from '../src/renderer/editor/layout-parking'

const band = { top: 0, bottom: 1000 }

describe('parkDecisions', () => {
  it('parks everything outside the band', () => {
    expect(
      parkDecisions(
        [
          { top: -900, height: 400 },
          { top: -50, height: 300 },
          { top: 200, height: 400 },
          { top: 1500, height: 300 },
        ],
        band,
      ),
    ).toEqual([true, false, false, true])
  })

  it('keeps a box straddling the band edge unparked', () => {
    expect(parkDecisions([{ top: 990, height: 40 }], band)).toEqual([false])
    expect(parkDecisions([{ top: -20, height: 40 }], band)).toEqual([false])
  })

  it('parks a box that ends exactly at the band top', () => {
    expect(parkDecisions([{ top: -400, height: 400 }], band)).toEqual([true])
  })

  it('answers nothing for no boxes', () => {
    expect(parkDecisions([], band)).toEqual([])
  })
})

describe('parkBand', () => {
  it('grows the scroll position by one viewport of headroom each way', () => {
    expect(parkBand({ scrollTop: 5000, clientHeight: 600 })).toEqual({
      top: 4400,
      bottom: 6200,
    })
  })
})

describe('parkingArmDecided: the sticky per-block cost gate', () => {
  const { parkingArmDecided } = mod

  it('never arms a cheap-to-lay-out document (plain text 183 pages)', () => {
    // measured: 166 ms over 5182 blocks = 0.032 ms/block
    expect(parkingArmDecided(166, 5182, false)).toBe(false)
    expect(parkingArmDecided(441, 5182, false)).toBe(false)
  })

  it('arms an expensive-to-lay-out document (table-heavy)', () => {
    // measured: 295 ms over 122 blocks = 2.4 ms/block
    expect(parkingArmDecided(295, 122, false)).toBe(true)
    expect(parkingArmDecided(151, 122, false)).toBe(true)
  })

  it('stays armed once a document has proven expensive', () => {
    expect(parkingArmDecided(10, 122, true)).toBe(true)
  })

  it('disarms when the document shrinks below the floor', () => {
    expect(parkingArmDecided(295, 40, true)).toBe(false)
    expect(parkingArmDecided(295, 0, true)).toBe(false)
  })
})

describe('parkingEntriesFor', () => {
  const doc = {
    childCount: 3,
    child(i: number) {
      return { nodeSize: [12, 15, 9][i] }
    },
  } as unknown as Parameters<typeof parkingEntriesFor>[0]

  it('collects positions only for parked blocks, with their heights', () => {
    const entries = parkingEntriesFor(doc, [false, true, true], [40, 200, 90])
    expect(entries).toEqual([
      { pos: 12, h: 200 },
      { pos: 27, h: 90 },
    ])
  })

  it('stops at the document child count', () => {
    const entries = parkingEntriesFor(doc, [true, true, true, true], [1, 2, 3, 4, 5])
    expect(entries).toHaveLength(3)
  })
})

describe('the parking plugin through a real editor', () => {
  const paragraphs = Array.from({ length: 6 }, (_, i) => ({
    type: 'docParagraph',
    attrs: { docxIndex: null },
    content: [{ type: 'text', text: `Paragraph ${i} body copy.` }],
  }))
  const content = { type: 'doc', content: paragraphs }

  const makeEditor = (): Editor =>
    new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: content as never,
    })

  const setParking = (editor: Editor, entries: ParkEntry[]) => {
    editor.view.dispatch(editor.state.tr.setMeta(parkingKey, { entries }))
  }

  it('applies node decorations with the remembered height for parked entries', () => {
    const editor = makeEditor()
    // park the second and fifth paragraphs
    const positions: number[] = []
    let pos = 0
    for (let i = 0; i < 6; i++) {
      if (i === 1 || i === 4) positions.push(pos)
      pos += editor.state.doc.child(i).nodeSize
    }
    setParking(editor, [
      { pos: positions[0]!, h: 111 },
      { pos: positions[1]!, h: 222 },
    ])
    const el1 = editor.view.nodeDOM(positions[0]!) as HTMLElement
    const el2 = editor.view.nodeDOM(positions[1]!) as HTMLElement
    expect(el1.classList.contains('cv-parked')).toBe(true)
    expect(el1.style.getPropertyValue('--park-h')).toBe('111px')
    expect(el2.classList.contains('cv-parked')).toBe(true)
    expect(el2.style.getPropertyValue('--park-h')).toBe('222px')
  })

  it('typing into an unparked paragraph keeps the parked set applied', () => {
    const editor = makeEditor()
    const positions: number[] = []
    let pos = 0
    for (let i = 0; i < 6; i++) {
      if (i === 1) positions.push(pos)
      pos += editor.state.doc.child(i).nodeSize
    }
    setParking(editor, [{ pos: positions[0]!, h: 111 }])
    // plain typing away from the parked entry: insertionText at position 4
    apply(editor, (tr) => tr.insertText('x', 4))
    // read through the mapped position: the entry maps forward past the
    // insertion, and PM may rebuild the node's DOM when it applies the
    // decoration set
    const mapped = parkingKey.getState(editor.state)!.entries[0]!
    const el = editor.view.nodeDOM(mapped.pos) as HTMLElement
    expect(el).not.toBeNull()
    expect(el.classList.contains('cv-parked')).toBe(true)
    expect(el.style.getPropertyValue('--park-h')).toBe('111px')
  })

  it('maps a parked entry forward when an earlier paragraph grows', () => {
    const editor = makeEditor()
    const positions: number[] = []
    let pos = 0
    for (let i = 0; i < 6; i++) {
      if (i === 2) positions.push(pos)
      pos += editor.state.doc.child(i).nodeSize
    }
    setParking(editor, [{ pos: positions[0]!, h: 90 }])
    const before = parkingKey.getState(editor.state)!.entries[0]!.pos
    apply(editor, (tr) => tr.insertText('growing text ', 4))
    const after = parkingKey.getState(editor.state)!.entries[0]!
    expect(after.pos).toBeGreaterThan(before)
    expect(after.h).toBe(90)
    const el = editor.view.nodeDOM(after.pos) as HTMLElement
    expect(el.classList.contains('cv-parked')).toBe(true)
  })

  it('replacing the wholesale set reveals the previously parked blocks', () => {
    const editor = makeEditor()
    const positions: number[] = []
    let pos = 0
    for (let i = 0; i < 6; i++) {
      positions.push(pos)
      pos += editor.state.doc.child(i).nodeSize
    }
    setParking(editor, [{ pos: positions[1]!, h: 50 }])
    const el1 = editor.view.nodeDOM(positions[1]!) as HTMLElement
    expect(el1.classList.contains('cv-parked')).toBe(true)
    setParking(editor, [])
    expect(el1.classList.contains('cv-parked')).toBe(false)
  })

  function apply(editor: Editor, fn: (tr: import('@tiptap/pm/state').Transaction) => void) {
    const tr = editor.state.tr
    fn(tr)
    editor.view.dispatch(tr)
  }
})
