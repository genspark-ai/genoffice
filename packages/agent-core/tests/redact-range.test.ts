import { describe, expect, it } from 'vitest'
import { redactTextBetween, type WithheldNode } from '../src/redact-range'

/**
 * The range walker, on its own.
 *
 * It takes a plain node rather than a ProseMirror document on purpose, so the
 * properties that matter can be pinned here without a schema: a withheld span
 * never contributes its words, a range that clips one still gets a whole
 * marker, and an unwithheld tree comes back out unchanged.
 *
 * The offsets are the part that can go quietly wrong, so `sizeOf` below is a
 * second, independent implementation of the arithmetic, and the round trip
 * checks the two against each other over a range of positions rather than a
 * single happy path.
 */

const MARK = 'redaction'
const options = { markName: MARK, marker: (label: string) => `{{${label}}}` }

/** the same count ProseMirror makes, written the obvious way */
function sizeOf(node: WithheldNode): number {
  if (typeof node.text === 'string') return node.text.length
  if (node.type === 'docInlineImage' || node.type === 'image') return 1
  if (!Array.isArray(node.content)) return 1
  return 2 + node.content.reduce((sum, c) => sum + sizeOf(c), 0)
}

const doc = (paragraphs: WithheldNode[][]): WithheldNode => ({
  type: 'doc',
  content: paragraphs.map((content) => ({ type: 'docParagraph', content }) as WithheldNode),
})

const t = (text: string, label?: string): WithheldNode => ({
  type: 'text',
  text,
  ...(label ? { marks: [{ type: MARK, attrs: { label } }] } : {}),
})

const SECRET = '13800138000'

describe('a range of a document with nothing withheld', () => {
  const plain = doc([
    [t('Base '), t('added '), t('and '), t('gone '), t('words.')],
    [t('Second block here')],
  ])

  it('starts on the character the range starts on', () => {
    // One paragraph, several runs: the offsets are what can go quietly wrong,
    // and an off-by-one drops the first character. A single paragraph is what
    // makes `slice(from - 1)` line up with the position space at all — across a
    // block boundary the two are offset by the block's own tokens.
    const one = doc([[t('Base '), t('added '), t('and '), t('gone '), t('words.')]])
    const end = 2 + sizeOf(one.content![0]!)
    const all = 'Base added and gone words.'
    for (let from = 1; from <= end; from++) {
      expect(redactTextBetween(one, from, end, options).replace(/\s+/g, ''), `from ${from}`).toBe(
        all.slice(from - 1).replace(/\s+/g, ''),
      )
    }
  })

  it('reads every block, in order, across a boundary', () => {
    const end = 2 + plain.content!.reduce((sum, b) => sum + sizeOf(b), 0)
    expect(redactTextBetween(plain, 1, end, options)).toBe(
      'Base added and gone words.\nSecond block here',
    )
  })
})

describe('a range with a withheld span in it', () => {
  const marked = doc([[t('Call '), t(SECRET, '客户电话'), t(' about the invoice')], [t('Second')]])

  it('never returns the words', () => {
    const end = 2 + marked.content!.reduce((sum, b) => sum + sizeOf(b), 0)
    for (let from = 1; from <= end; from++) {
      expect(redactTextBetween(marked, from, end, options), `from ${from}`).not.toContain(SECRET)
    }
  })

  it('stands the whole marker in, even for a range that clips the span', () => {
    const out = redactTextBetween(marked, 7, 10, options)
    expect(out).toContain('{{客户电话}}')
    expect(out).not.toMatch(/\{\{[^}]*$/)
    expect(out).not.toMatch(/^[^{]*\}\}/)
  })

  it('leaves the text around the span alone', () => {
    const out = redactTextBetween(marked, 1, 99, options)
    expect(out).toContain('Call ')
    expect(out).toContain('about the invoice')
    expect(out).toContain('Second')
  })
})

describe('leaves it alone', () => {
  it('when the range is empty or reversed', () => {
    const plain = doc([[t('hello')]])
    expect(redactTextBetween(plain, 5, 5, options)).toBe('')
    expect(redactTextBetween(plain, 9, 2, options)).toBe('')
  })

  it('when the node has no blocks at all', () => {
    expect(redactTextBetween({ type: 'doc' }, 1, 9, options)).toBe('')
  })
})

describe('a range the withheld content sits outside', () => {
  it('a leaf before the selection contributes nothing', () => {
    const d: WithheldNode = {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          content: [{ type: 'image' }, { type: 'text', text: 'XY' }],
        },
      ],
    }
    const isLeaf = (n: WithheldNode) => n.type === 'image'
    // content starts at 1: image occupies [1,2), 'XY' at [2,4)
    expect(redactTextBetween(d, 2, 4, { ...options, isLeaf, leafText: '\u00b7' })).toBe('XY')
    expect(redactTextBetween(d, 1, 4, { ...options, isLeaf, leafText: '\u00b7' })).toBe('\u00b7XY')
  })

  it('a marked block the selection never reaches contributes no marker', () => {
    // defensive shape: the mark on the block node itself, selection in block 1
    const d: WithheldNode = {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          marks: [{ type: MARK, attrs: { label: 'secret' } }],
          content: [t('AAAA')],
        },
        { type: 'docParagraph', content: [t('B')] },
      ],
    }
    // block 0 occupies [0,6), block 1 content at 7
    expect(redactTextBetween(d, 7, 8, options)).toBe('B')
    expect(redactTextBetween(d, 1, 3, options)).toBe('{{secret}}')
  })

  it('an empty paragraph between two selected ones still costs its line break', () => {
    const d: WithheldNode = {
      type: 'doc',
      content: [
        { type: 'docParagraph', content: [t('One')] },
        { type: 'docParagraph', content: [] },
        { type: 'docParagraph', content: [t('Two')] },
      ],
    }
    // block sizes: 5, 2, 5 — 'One' at [1,4), 'Two' at [8,11)
    expect(redactTextBetween(d, 1, 11, options)).toBe('One\n\nTwo')
  })
})

describe('a range inside a nested inline container', () => {
  it('reads the characters at their own offsets, not one earlier', () => {
    const d: WithheldNode = {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          content: [
            { type: 'text', text: 'AB' },
            { type: 'bold', content: [{ type: 'text', text: 'cd' }] },
            { type: 'text', text: 'EF' },
          ],
        },
      ],
    }
    // content starts at 1: 'AB' [1,3), bold open 3, 'cd' [4,6), close 6, 'EF' [7,9)
    expect(redactTextBetween(d, 4, 6, options)).toBe('cd')
    expect(redactTextBetween(d, 1, 9, options)).toBe('ABcdEF')
    // a range clipping only the second character still slices the right one
    expect(redactTextBetween(d, 5, 6, options)).toBe('d')
  })
})
