/**
 * The parse map is rebuilt on every commit, so the rebuild has to stay cheap
 * on large documents: nearest-sid matching is asserted to touch the previous
 * entries a bounded number of times rather than once per candidate pair, and
 * the sid seed must survive a document with more entries than the engine's
 * argument limit.
 */
import { describe, expect, it } from 'vitest'
import { buildParseMap, type ElementEntry, type ParseMap } from '../src/renderer/document/parse-map'

const FRAGMENT = '<p>x</p>'
const FRAGMENT_PATH = 'p:nth-of-type(1)'
const UNRELATED_PATH = 'div:nth-of-type(1) > p:nth-of-type(1)'

function synthetic(
  entries: Array<Pick<ElementEntry, 'sid' | 'tag' | 'path' | 'startTag'>>,
): ParseMap {
  return {
    version: 0,
    elements: entries.map((e) => ({
      ...e,
      parentSid: null,
      depth: 0,
      range: [e.startTag[0], e.startTag[1]],
      startTag: e.startTag,
      endTag: null,
      inner: [e.startTag[1], e.startTag[1]],
      textNodes: [],
    })),
    bySid: new Map(entries.map((e) => [e.sid, e])),
    errorCount: 0,
  }
}

describe('nearest-sid matching', () => {
  const at = FRAGMENT.indexOf('<p>')

  it('reuses the previous sid for an unchanged element', () => {
    const first = buildParseMap(FRAGMENT, 1)
    const second = buildParseMap(FRAGMENT, 2, first)
    expect(second.elements.map((e) => e.sid)).toEqual(first.elements.map((e) => e.sid))
  })

  it('picks the closest of several previous entries sharing a tag and path', () => {
    const previous = synthetic([
      { sid: 41, tag: 'p', path: FRAGMENT_PATH, startTag: [at - 40, at - 37] },
      { sid: 42, tag: 'p', path: FRAGMENT_PATH, startTag: [at - 2, at + 1] },
      { sid: 43, tag: 'p', path: FRAGMENT_PATH, startTag: [at + 60, at + 63] },
    ])
    const next = buildParseMap(FRAGMENT, 1, previous)
    const p = next.elements.find((e) => e.tag === 'p')!
    expect(p.sid).toBe(42)
  })

  it('breaks an equidistant tie toward the smaller start offset', () => {
    const previous = synthetic([
      { sid: 51, tag: 'p', path: FRAGMENT_PATH, startTag: [at + 2, at + 5] },
      { sid: 52, tag: 'p', path: FRAGMENT_PATH, startTag: [at - 2, at - 1] },
    ])
    const next = buildParseMap(FRAGMENT, 1, previous)
    expect(next.elements.find((e) => e.tag === 'p')!.sid).toBe(52)
  })
})

describe('sid reuse across a sequence of edits', () => {
  it('keeps sids unique and preserves the ones that did not move', () => {
    const versions = [
      '<div><p>one</p><p>two</p></div>',
      '<div><p>one</p><p>two</p><p>three</p></div>',
      '<div><h2>head</h2><p>one</p><p>two</p><p>three</p></div>',
      '<div><p>one</p><h2>head</h2><p>two</p><p>three</p></div>',
    ]
    let previous: ParseMap | null = null
    for (const [i, text] of versions.entries()) {
      const next = buildParseMap(text, i + 1, previous)
      const sids = next.elements.map((e) => e.sid)
      expect(new Set(sids).size).toBe(sids.length)
      if (previous) {
        for (const entry of next.elements) {
          const old = previous.bySid.get(entry.sid)
          if (old) expect(old.path).toBe(entry.path)
        }
      }
      previous = next
    }
  })
})

describe('rebuild cost', () => {
  const doc = (sections: number): string =>
    `<body>${Array.from({ length: sections }, (_, i) => `<section><h2>t${i}</h2><p>x</p></section>`).join('')}</body>`

  const counted = (previous: ParseMap, text: string): number => {
    let reads = 0
    const elements = new Proxy(previous.elements, {
      get(target, prop, receiver) {
        if (typeof prop === 'string' && /^\d+$/.test(prop)) reads++
        return Reflect.get(target, prop, receiver) as unknown
      },
    })
    buildParseMap(text, 2, { ...previous, elements })
    return reads
  }

  it('reads the previous entries a bounded number of times, not once per element', () => {
    const small = buildParseMap(doc(500), 1)
    const large = buildParseMap(doc(2000), 1)
    const smallReads = counted(small, doc(500))
    const largeReads = counted(large, doc(2000))

    expect(small.elements.length).toBeLessThan(3000)
    expect(large.elements.length).toBeGreaterThan(small.elements.length)

    // A per-candidate scan is quadratic: 4x the elements would be ~16x the reads.
    // The bucketed index is linear, so reads must stay within a small multiple
    // of the element count.
    expect(largeReads).toBeLessThan(large.elements.length * 8)
    expect(largeReads / large.elements.length).toBeLessThan(8)
    expect(largeReads).toBeLessThan(
      smallReads * (large.elements.length / small.elements.length + 1),
    )
  })
})

describe('sid seed on documents larger than the argument limit', () => {
  it('does not throw a RangeError when spreading every sid', () => {
    const total = 200_000
    const previous = synthetic(
      Array.from({ length: total }, (_, i) => ({
        sid: i + 1,
        tag: 'p',
        path: UNRELATED_PATH,
        startTag: [i * 7, i * 7 + 3] as [number, number],
      })),
    )
    const next = buildParseMap(FRAGMENT, 1, previous)
    const p = next.elements.find((e) => e.tag === 'p')!
    expect(p.sid).toBe(total + 1)
    expect(new Set(next.elements.map((e) => e.sid)).size).toBe(next.elements.length)
  })

  it('seeds from 1 with no previous map', () => {
    expect(buildParseMap(FRAGMENT, 1).elements[0]!.sid).toBe(1)
  })
})
