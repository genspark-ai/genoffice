/** Global find/replace: in-run matching, case toggle, firstOnly/element scope, dynamic-field skip, byte fidelity. */
import { describe, it, expect } from 'vitest'
import { parseSlide } from '../src/parse'
import { patchedElementXml, replaceAllInDeck } from '../src/index'
import type { SlideDeck, TextElement } from '../src/types'

const slideWith = (sps: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a"><p:cSld>' +
  `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${sps}</p:spTree></p:cSld></p:sld>`
const sp = (runs: string) =>
  '<p:sp><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>' +
  `<p:txBody><a:bodyPr/><a:p>${runs}</a:p></p:txBody></p:sp>`

const deckWith = (...slidesSps: string[]): SlideDeck => {
  const slides = slidesSps.map((sps, i) =>
    parseSlide({ path: `ppt/slides/slide${i + 1}.xml`, slideXml: slideWith(sps), ctx: {} }),
  )
  return { slides } as unknown as SlideDeck
}

describe('replaceAllInDeck', () => {
  it('replaces across slides + counts + marks only changed slides', () => {
    const deck = deckWith(
      sp('<a:r><a:t>Hello world, hello</a:t></a:r>'),
      sp('<a:r><a:t>no match here</a:t></a:r>'),
      sp('<a:r><a:t>Hello again</a:t></a:r>'),
    )
    const r = replaceAllInDeck(deck, 'hello', 'hi')
    expect(r.count).toBe(3)
    expect(r.changedSlides).toEqual([0, 2])
    const el = deck.slides[0]!.elements[0] as TextElement
    expect(el.text!.paragraphs[0]!.runs[0]!.text).toBe('hi world, hi')
    expect(el.dirty).toBe(true)
  })

  it('matchCase is case-sensitive', () => {
    const deck = deckWith(sp('<a:r><a:t>Hello hello</a:t></a:r>'))
    const r = replaceAllInDeck(deck, 'hello', 'x', { matchCase: true })
    expect(r.count).toBe(1)
    expect((deck.slides[0]!.elements[0] as TextElement).text!.paragraphs[0]!.runs[0]!.text).toBe(
      'Hello x',
    )
  })

  it('firstOnly + element scope replaces only the first match', () => {
    const deck = deckWith(sp('<a:r><a:t>aaa</a:t></a:r>'), sp('<a:r><a:t>aaa</a:t></a:r>'))
    const el0 = deck.slides[0]!.elements[0]!
    const r = replaceAllInDeck(deck, 'a', 'b', {
      firstOnly: true,
      slideIndex: 0,
      elementId: el0.id,
    })
    expect(r.count).toBe(1)
    expect((el0 as TextElement).text!.paragraphs[0]!.runs[0]!.text).toBe('baa')
    expect((deck.slides[1]!.elements[0] as TextElement).text!.paragraphs[0]!.runs[0]!.text).toBe(
      'aaa',
    )
  })

  it('dynamic field runs skipped; rPr keeps original bytes after replacement (in-place patch)', () => {
    const deck = deckWith(
      sp(
        '<a:r><a:rPr sz="1800" b="1"><a:latin typeface="Calibri"/></a:rPr><a:t>foo bar</a:t></a:r>',
      ) + sp('<a:fld id="{X}" type="slidenum"><a:rPr/><a:t>foo</a:t></a:fld>'),
    )
    const r = replaceAllInDeck(deck, 'foo', 'baz')
    expect(r.count).toBe(1) // 'foo' inside fld is untouched
    const el = deck.slides[0]!.elements[0]!
    const out = patchedElementXml(el)
    expect(out).toContain('baz bar')
    // Font size/bold/font declarations kept (explicit b/i booleans overriding inheritance is existing semantics of the alignment path)
    expect(out).toMatch(
      /<a:rPr[^>]*\bsz="1800"[^>]*\bb="1"[^>]*><a:latin typeface="Calibri"\/><\/a:rPr>/,
    )
  })

  it('regex special characters match literally', () => {
    const deck = deckWith(sp('<a:r><a:t>1+1=2 (a.b)</a:t></a:r>'))
    expect(replaceAllInDeck(deck, '1+1', '2').count).toBe(1)
    expect(replaceAllInDeck(deck, '(a.b)', 'c').count).toBe(1)
    expect((deck.slides[0]!.elements[0] as TextElement).text!.paragraphs[0]!.runs[0]!.text).toBe(
      '2=2 c',
    )
  })
})

describe('replaceAllInDeck across runs (#1005)', () => {
  const runs0 = (deck: SlideDeck) =>
    (deck.slides[0]!.elements[0] as TextElement).text!.paragraphs[0]!.runs

  it('a match spanning two runs is replaced once and keeps each run its own text', () => {
    const deck = deckWith(
      sp('<a:r><a:rPr b="1"/><a:t>Hel</a:t></a:r><a:r><a:t>lo world</a:t></a:r>'),
    )
    expect(runs0(deck)).toHaveLength(2)
    expect(replaceAllInDeck(deck, 'Hello', 'Hi').count).toBe(1)
    expect(runs0(deck).map((r) => r.text)).toEqual(['Hi', ' world'])
  })

  it('the replacement takes the formatting of the run holding the match start', () => {
    const deck = deckWith(
      sp('<a:r><a:rPr b="1"/><a:t>Hel</a:t></a:r><a:r><a:rPr i="1"/><a:t>lo</a:t></a:r>'),
    )
    replaceAllInDeck(deck, 'Hello', 'Hi')
    const [first, second] = runs0(deck)
    expect(first!.text).toBe('Hi')
    expect(first!.bold).toBe(true)
    expect(first!.italic).toBe(false)
    expect(second!.text).toBe('')
    expect(second!.italic).toBe(true)
  })

  it('a match spanning three runs cuts the middle one without touching its rPr', () => {
    const deck = deckWith(
      sp(
        '<a:r><a:rPr b="1"/><a:t>ab</a:t></a:r>' +
          '<a:r><a:rPr i="1"/><a:t>cd</a:t></a:r>' +
          '<a:r><a:rPr u="sng"/><a:t>ef</a:t></a:r>',
      ),
    )
    expect(replaceAllInDeck(deck, 'bcde', 'X').count).toBe(1)
    const runs = runs0(deck)
    expect(runs.map((r) => r.text)).toEqual(['aX', '', 'f'])
    expect(runs[1]!.italic).toBe(true)
    const out = patchedElementXml(deck.slides[0]!.elements[0]!)
    expect(out).toContain('<a:t>aX</a:t>')
    expect(out).toContain('<a:t>f</a:t>')
    expect(out).toMatch(/<a:rPr[^>]*\bi="1"[^>]*\/>/)
  })

  it('each cross-run match is counted and replaced, not just the first', () => {
    const deck = deckWith(sp('<a:r><a:rPr b="1"/><a:t>x-y </a:t></a:r><a:r><a:t>x-y</a:t></a:r>'))
    expect(replaceAllInDeck(deck, 'x-y', 'Z').count).toBe(2)
    expect(runs0(deck).map((r) => r.text)).toEqual(['Z ', 'Z'])
  })

  it('firstOnly stops at the first match even when a later one spans runs', () => {
    const deck = deckWith(sp('<a:r><a:rPr b="1"/><a:t>ab</a:t></a:r><a:r><a:t>cd</a:t></a:r>'))
    expect(replaceAllInDeck(deck, 'bc', 'X', { firstOnly: true }).count).toBe(1)
    expect(runs0(deck).map((r) => r.text)).toEqual(['aX', 'd'])
  })

  it('matchCase applies across the run boundary', () => {
    const deck = deckWith(sp('<a:r><a:rPr b="1"/><a:t>Hel</a:t></a:r><a:r><a:t>LO x</a:t></a:r>'))
    expect(replaceAllInDeck(deck, 'hello', 'hi', { matchCase: true }).count).toBe(0)
    expect(runs0(deck).map((r) => r.text)).toEqual(['Hel', 'LO x'])
  })

  it('a dynamic field run is a barrier: no match spans it', () => {
    const deck = deckWith(
      sp(
        '<a:r><a:rPr b="1"/><a:t>ab</a:t></a:r>' +
          '<a:fld id="{X}" type="slidenum"><a:rPr/><a:t>cd</a:t></a:fld>' +
          '<a:r><a:t>ef</a:t></a:r>',
      ),
    )
    // 'b' and 'e' are only adjacent once the field is skipped: no match may span it
    expect(replaceAllInDeck(deck, 'be', 'X').count).toBe(0)
    expect(runs0(deck).map((r) => r.text)).toEqual(['ab', 'cd', 'ef'])
    // nor may a match touch the field's own text
    expect(replaceAllInDeck(deck, 'bcd', 'X').count).toBe(0)
    expect(replaceAllInDeck(deck, 'cd', 'X').count).toBe(0)
    expect(runs0(deck).map((r) => r.text)).toEqual(['ab', 'cd', 'ef'])
  })

  it('runs adjacent to a field still replace within themselves', () => {
    const deck = deckWith(
      sp(
        '<a:r><a:rPr b="1"/><a:t>ab</a:t></a:r>' +
          '<a:fld id="{X}" type="slidenum"><a:rPr/><a:t>cd</a:t></a:fld>' +
          '<a:r><a:t>ef</a:t></a:r>',
      ),
    )
    expect(replaceAllInDeck(deck, 'ef', 'Z').count).toBe(1)
    expect(replaceAllInDeck(deck, 'b', 'Y').count).toBe(1)
    expect(runs0(deck).map((r) => r.text)).toEqual(['aY', 'cd', 'Z'])
  })

  it('a cross-run match still works on either side of a field', () => {
    const deck = deckWith(
      sp(
        '<a:r><a:t>ab</a:t></a:r><a:r><a:t>cd</a:t></a:r>' +
          '<a:fld id="{X}" type="slidenum"><a:rPr/><a:t>#</a:t></a:fld>' +
          '<a:r><a:t>ef</a:t></a:r><a:r><a:t>gh</a:t></a:r>',
      ),
    )
    expect(replaceAllInDeck(deck, 'bc', 'X').count).toBe(1)
    expect(replaceAllInDeck(deck, 'fg', 'Y').count).toBe(1)
    expect(runs0(deck).map((r) => r.text)).toEqual(['aX', 'd', '#', 'eY', 'h'])
  })

  it('a needle containing a newline never blanks an <a:br> soft-break sentinel run', () => {
    const deck = deckWith(sp('<a:r><a:t>ab</a:t></a:r><a:br/><a:r><a:t>cd</a:t></a:r>'))
    expect(runs0(deck).map((r) => r.text)).toEqual(['ab', '\n', 'cd'])
    expect(replaceAllInDeck(deck, 'b\nc', 'X').count).toBe(0)
    expect(replaceAllInDeck(deck, '\n', 'X').count).toBe(0)
    expect(runs0(deck).map((r) => r.text)).toEqual(['ab', '\n', 'cd'])
    expect(patchedElementXml(deck.slides[0]!.elements[0]!)).toContain('<a:br/>')
    // text either side of the break still replaces on its own
    expect(replaceAllInDeck(deck, 'cd', 'Z').count).toBe(1)
    expect(runs0(deck).map((r) => r.text)).toEqual(['ab', '\n', 'Z'])
  })

  it('text either side of a cross-run match keeps its run split', () => {
    const deck = deckWith(
      sp('<a:r><a:rPr b="1"/><a:t>keep-</a:t></a:r><a:r><a:t>-keep</a:t></a:r>'),
    )
    expect(replaceAllInDeck(deck, '--', '=').count).toBe(1)
    expect(runs0(deck).map((r) => r.text)).toEqual(['keep=', 'keep'])
  })
})
