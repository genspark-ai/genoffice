import { describe, expect, it } from 'vitest'
import {
  decideLineShrinks,
  sameLine,
  type ShrinkGap,
  type ShrinkLine,
} from '../src/renderer/editor/justify-shrink'

/**
 * Fixtures mirror the Word for Mac probe matrix (2026-08-27): TNR 14pt,
 * column 468pt, uniform 3.5pt spaces. A line of n words with the candidate
 * word overflowing the column by delta.
 */

const SP = 3.5
const COL = 468

function gap(width = SP, chars = 1, from = 0, to = 1): ShrinkGap {
  return { width, chars, from, to }
}

/** a justified line of `words` words whose pull candidate overflows by delta */
function line(words: number, delta: number, nextWordWidth: number, spaceW = SP): ShrinkLine {
  const gaps = Array.from({ length: words - 1 }, (_, i) => gap(spaceW, 1, i * 2, i * 2 + 1))
  const boundary = gap(spaceW, 1, 100, 101)
  // words sum so that: sum + interior + boundary + nextWord = COL + delta
  const sum = COL + delta - spaceW * words - nextWordWidth
  return {
    wordWidths: Array.from({ length: words }, () => sum / words),
    gaps,
    boundary,
    avail: COL,
    nextWordWidth,
  }
}

describe('decideLineShrinks — Word compat15 pull rule', () => {
  it('pulls within the tradeoff threshold (S=10, short word "de")', () => {
    // probe: w=13.21 pulls at delta=5.89, refuses at 6.27
    const [pull] = decideLineShrinks([line(10, 5.89, 13.21)])
    expect(pull).not.toBeNull()
    expect(pull!.gaps).toHaveLength(10)
    expect(pull!.perChar).toBeCloseTo((5.89 + 0.5) / 10, 5)
    const [no] = decideLineShrinks([line(10, 6.27, 13.21)])
    expect(no).toBeNull()
  })

  it('caps total shrink at 25% of the space width (S=10, "pour")', () => {
    // probe: w=25.66 pulls at 8.62 (cap 8.75), refuses at 9.00
    expect(decideLineShrinks([line(10, 8.62, 25.66)])[0]).not.toBeNull()
    expect(decideLineShrinks([line(10, 9.0, 25.66)])[0]).toBeNull()
  })

  it('cap scales with the space count (S=5 / S=15)', () => {
    // probe: S=5 pulls at 4.11 (cap 4.375), refuses at 4.44
    expect(decideLineShrinks([line(5, 4.11, 35)])[0]).not.toBeNull()
    expect(decideLineShrinks([line(5, 4.44, 35)])[0]).toBeNull()
    // probe: S=15 pulls at 12.89 (cap 13.125), refuses at 13.27
    expect(decideLineShrinks([line(15, 12.89, 35)])[0]).not.toBeNull()
    expect(decideLineShrinks([line(15, 13.27, 35)])[0]).toBeNull()
  })

  it('thresholds scale with the font size (28pt: space 7pt)', () => {
    // probe: 28pt "de" (26.43pt) pulls at 11.83, refuses at 12.59
    expect(decideLineShrinks([line(10, 11.83, 26.43, 7)])[0]).not.toBeNull()
    expect(decideLineShrinks([line(10, 12.59, 26.43, 7)])[0]).toBeNull()
  })

  it('keeps the compression of an already-pulled line', () => {
    const l = line(10, 5, 13.21)
    // simulate the pulled state: the word joined the line, natural > avail by 5
    l.wordWidths = [...l.wordWidths, 13.21]
    l.gaps = [...l.gaps, l.boundary!]
    l.boundary = null
    l.nextWordWidth = null
    const [keep] = decideLineShrinks([l])
    expect(keep).not.toBeNull()
    expect(keep!.perChar).toBeCloseTo((5 + 0.5) / 10, 3)
  })

  it('never pulls across a hard break or without a boundary space', () => {
    const l = line(10, 3, 13.21)
    l.boundary = null
    expect(decideLineShrinks([l])[0]).toBeNull()
  })

  it('pulls a word that overflows by a single LayoutUnit (Word fits it by twip rounding)', () => {
    // Cyrillic Times 14pt line whose natural width equals the column to the twip:
    // Chromium wraps the last word by ~1/64px, Word keeps it
    const [pull] = decideLineShrinks([line(10, 1 / 64, 40)])
    expect(pull).not.toBeNull()
    expect(pull!.perChar).toBeCloseTo((1 / 64 + 0.5) / 10, 4)
  })

  it('keeps the pull when the joined line overflows by a single LayoutUnit', () => {
    const l = line(10, 1 / 64, 40)
    l.wordWidths = [...l.wordWidths, 40]
    l.gaps = [...l.gaps, l.boundary!]
    l.boundary = null
    l.nextWordWidth = null
    const [keep] = decideLineShrinks([l])
    expect(keep).not.toBeNull()
  })

  it('leaves naturally fitting lines alone', () => {
    const l = line(10, -2, 13.21)
    expect(decideLineShrinks([l])[0]).toBeNull()
  })

  it('never divides by zero on a single-gap line', () => {
    const l: ShrinkLine = {
      wordWidths: [400],
      gaps: [],
      boundary: gap(),
      avail: COL,
      nextWordWidth: 60,
    }
    expect(decideLineShrinks([l])[0]).toBeNull()
  })
})

describe('sameLine — rendered-line grouping of word boxes', () => {
  // glyph boxes are ascent+descent tall regardless of line-height, so under
  // tight leading consecutive lines overlap by a few px; chaining them into one
  // "line" makes decideLineShrinks compress the whole paragraph (word-spacing
  // of minus hundreds of px, words overprinting in a narrow column)
  const box = (top: number, height: number) => ({ top, bottom: top + height, left: 0, right: 10 })

  it('keeps consecutive lines apart under w:line=204 auto (Arial 9pt: 14px box, 11.7px pitch)', () => {
    const l1 = box(100, 14)
    const l2 = box(100 + 11.73, 14)
    expect(sameLine(l1, l2)).toBe(false)
  })

  it('keeps consecutive lines apart at a 1.0 face factor (Helvetica 11pt: 16px box, 14.67px pitch)', () => {
    expect(sameLine(box(50, 16), box(50 + 14.67, 16))).toBe(false)
  })

  it('joins words on one baseline with different font sizes and a superscript', () => {
    const big = { top: 100, bottom: 100 + 29, left: 0, right: 10 } // 22pt
    const small = { top: 100 + 29 - 4 - 13, bottom: 100 + 29 - 4, left: 0, right: 10 } // 9pt, shared baseline
    const sup = { top: 100 + 29 - 4 - 13 - 5, bottom: 100 + 29 - 4 - 5, left: 0, right: 10 } // raised 5px
    expect(sameLine(big, small)).toBe(true)
    expect(sameLine(small, big)).toBe(true)
    expect(sameLine(small, sup)).toBe(true)
  })

  it('still separates lines that merely touch', () => {
    expect(sameLine(box(0, 14), box(14, 14))).toBe(false)
    expect(sameLine(box(0, 14), box(13.5, 14))).toBe(false)
  })
})

describe('shrink meta survives plugin state application', () => {
  it("keeps the meta array intact for 'transaction' listeners", async () => {
    const { Editor } = await import('@tiptap/core')
    const { Decoration } = await import('@tiptap/pm/view')
    const { editorExtensions } = await import('../src/renderer/editor/extensions')
    const { justifyShrinkPluginKey } = await import('../src/renderer/editor/justify-shrink')
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { align: 'justify' },
            content: [{ type: 'text', text: 'un paragraphe justifié avec des espaces' }],
          },
        ],
      },
    })
    try {
      const decos = [
        Decoration.inline(1, 3, { class: 'doc-jshrink', style: 'word-spacing:-0.1px' }),
        Decoration.inline(4, 8, { class: 'doc-jshrink', style: 'word-spacing:-0.2px' }),
      ]
      // DecorationSet.create (run by the plugin's apply) nulls out consumed
      // entries of the array it is handed; the App-level listener then read
      // null.from and broke Enter, paste and save on shrink-active documents.
      const seen: Array<unknown | null> = []
      editor.on('transaction', ({ transaction }) => {
        const meta = transaction.getMeta(justifyShrinkPluginKey) as Array<{ from: number } | null>
        if (meta) seen.push(...meta)
      })
      editor.view.dispatch(
        editor.state.tr.setMeta(justifyShrinkPluginKey, decos).setMeta('addToHistory', false),
      )
      expect(seen).toHaveLength(2)
      expect(seen.every((d) => d !== null)).toBe(true)
      expect(seen.map((d) => (d as { from: number }).from)).toEqual([1, 4])
      // the original array must not have been consumed either
      expect(decos.every((d) => d !== null)).toBe(true)
    } finally {
      editor.destroy()
    }
  })
})
