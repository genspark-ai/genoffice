import { describe, expect, it } from 'vitest'
import { Fragment, Schema, type Node as PmNode } from '@tiptap/pm/model'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import {
  fastFindIndex,
  fastForChild,
  withOriginalFindIndex,
  withOriginalForChild,
} from '../src/renderer/editor/prosemirror-perf'

const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: { content: 'text*', group: 'block' },
    quote: { content: 'paragraph+', group: 'block' },
    text: {},
  },
})

// deterministic LCG so a failure reproduces
let seed = 12345
const rnd = (n: number) => {
  seed = (seed * 1103515245 + 12345) & 0x7fffffff
  return seed % n
}

function makeDoc(blocks: number): PmNode {
  const children: PmNode[] = []
  for (let i = 0; i < blocks; i++) {
    const para = () => schema.node('paragraph', null, schema.text('x'.repeat(1 + rnd(12))))
    children.push(rnd(5) === 0 ? schema.node('quote', null, [para(), para()]) : para())
  }
  return schema.node('doc', null, children)
}

function randomDecorations(doc: PmNode, count: number): Decoration[] {
  const size = doc.content.size
  const out: Decoration[] = []
  for (let i = 0; i < count; i++) {
    const kind = rnd(4)
    if (kind === 0) {
      // node decoration on a top-level or nested block
      const pos = rnd(size)
      const $pos = doc.resolve(pos)
      const depth = Math.max(1, $pos.depth)
      const before = $pos.before(depth)
      const node = doc.nodeAt(before)!
      out.push(Decoration.node(before, before + node.nodeSize, { class: `n${i}` }))
    } else if (kind === 1) {
      out.push(Decoration.widget(rnd(size + 1), () => document.createElement('span'), { id: i }))
    } else {
      // inline decoration; every other one deliberately crosses a block boundary
      const a = rnd(size)
      const b = Math.min(size, a + 1 + rnd(kind === 2 ? 3 : 40))
      if (a < b) out.push(Decoration.inline(a, b, { class: `i${i}` }))
    }
  }
  return out
}

type Source = DecorationSet | { members: DecorationSet[] }
const flatten = (set: Source): Decoration[] =>
  'members' in set ? set.members.flatMap((m) => m.find()) : set.find()
const dump = (set: Source) =>
  flatten(set)
    .map((d) => `${d.from}-${d.to}:${JSON.stringify(d.spec)}`)
    .sort()

function compareAll(doc: PmNode, set: DecorationSet, parent: PmNode, base: number): number {
  let checked = 0
  parent.forEach((child, offset) => {
    const at = base + offset
    const fast = fastForChild.call(set, at, child)
    const slow = withOriginalForChild(() => set.forChild(at, child))
    expect(dump(fast as unknown as Source)).toEqual(dump(slow as unknown as Source))
    checked++
    if (
      !child.isLeaf &&
      !child.isTextblock &&
      fast !== DecorationSet.empty &&
      fast instanceof DecorationSet
    ) {
      checked += compareAll(doc, fast, child, at + 1)
    }
  })
  return checked
}

describe('DecorationSet.forChild override', () => {
  it('matches the upstream result on random documents and decoration mixes', () => {
    let checked = 0
    for (let round = 0; round < 40; round++) {
      const doc = makeDoc(5 + rnd(60))
      const set = DecorationSet.create(doc, randomDecorations(doc, rnd(80)))
      checked += compareAll(doc, set, doc, 0)
    }
    expect(checked).toBeGreaterThan(1000)
  })

  it('returns the shared empty set and child subtrees by identity like upstream', () => {
    const doc = makeDoc(30)
    const decos = randomDecorations(doc, 60).filter(
      (d) => !(d as unknown as { inline: boolean }).inline || d.to - d.from < 3,
    )
    const set = DecorationSet.create(doc, decos)
    doc.forEach((child, offset) => {
      const fast = fastForChild.call(set, offset, child)
      const slow = withOriginalForChild(() => set.forChild(offset, child))
      expect(fast).toBe(slow)
    })
  })
})

const findIndexOf = (frag: Fragment, pos: number) =>
  (frag as unknown as { findIndex: (p: number) => { index: number; offset: number } }).findIndex(
    pos,
  )

describe('Fragment.findIndex override', () => {
  it('matches the upstream walk for every position of small and large fragments', () => {
    for (const blocks of [1, 3, 63, 64, 65, 200, 1000]) {
      const doc = makeDoc(blocks)
      const frag = doc.content
      for (let pos = 0; pos <= frag.size; pos++) {
        const fast = { ...fastFindIndex.call(frag, pos) }
        const slow = withOriginalFindIndex(() => ({ ...findIndexOf(frag, pos) }))
        expect(fast).toEqual(slow)
      }
      expect(() => fastFindIndex.call(frag, frag.size + 1)).toThrow(RangeError)
      expect(() => fastFindIndex.call(frag, -1)).toThrow(RangeError)
    }
  })

  it('keeps resolve / nodeAt / childAfter exact on a long document', () => {
    const doc = makeDoc(500)
    for (let round = 0; round < 300; round++) {
      const pos = rnd(doc.content.size + 1)
      const fast = doc.resolve(pos)
      const slow = withOriginalFindIndex(() => doc.resolve(pos))
      expect([fast.depth, fast.parentOffset, fast.index(0)]).toEqual([
        slow.depth,
        slow.parentOffset,
        slow.index(0),
      ])
      expect(doc.nodeAt(pos)).toBe(withOriginalFindIndex(() => doc.nodeAt(pos)))
    }
    expect(findIndexOf(Fragment.empty, 0)).toEqual({ index: 0, offset: 0 })
  })
})
