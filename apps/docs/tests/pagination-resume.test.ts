import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { firstChangedTopLevelIndex } from '../src/renderer/editor/pagination-fast-path'
import {
  isResumePage,
  sliceWithLineSplit,
  type BlockBox,
  type PageSlice,
  type PassResume,
  type SectionGeom,
  type SliceOutputs,
} from '../src/renderer/pagination'

// deterministic pseudo-random documents: paragraphs with line boxes, keep
// rules, explicit breaks, empty marks, oversized blocks, several sections
// (forced / continuous, title-page capacity, parity, two columns)
function seeded(seed: number): () => number {
  let s = seed >>> 0
  return () => {
    s = (s * 1664525 + 1013904223) >>> 0
    return s / 2 ** 32
  }
}

interface Doc {
  blocks: BlockBox[]
  geoms: SectionGeom[]
  total: number
}

function randomDoc(seed: number, n = 160): Doc {
  const rnd = seeded(seed)
  const pick = <T>(xs: T[]) => xs[Math.floor(rnd() * xs.length)]
  const nSec = 1 + Math.floor(rnd() * 3)
  const geoms: SectionGeom[] = []
  for (let s = 0; s < nSec; s++) {
    const contentHeight = 300 + Math.floor(rnd() * 200)
    geoms.push({
      contentHeight,
      forceBreak: s === 0 || rnd() < 0.5,
      ...(rnd() < 0.3 ? { firstContentHeight: contentHeight - 60 } : {}),
      ...(s > 0 && rnd() < 0.2 ? { startType: 'evenPage' as const } : {}),
      ...(rnd() < 0.2 ? { cols: 2 } : {}),
    })
  }
  const secStarts = new Set<number>()
  for (let s = 1; s < nSec; s++) secStarts.add(Math.floor(rnd() * n))
  const blocks: BlockBox[] = []
  let y = 0
  let section = 0
  for (let i = 0; i < n; i++) {
    if (secStarts.has(i) && section < nSec - 1) section++
    const spaceBeforePx = pick([0, 0, 6, 12])
    const spaceAfterPx = pick([0, 0, 8])
    const empty = rnd() < 0.1
    const lines = empty ? 1 : rnd() < 0.05 ? 30 : 1 + Math.floor(rnd() * 8)
    const lh = 14 + Math.floor(rnd() * 8)
    const lineBoxes = Array.from({ length: lines }, (_, k) => ({
      offsetInBlock: spaceBeforePx + k * lh,
      height: lh,
    }))
    const height = spaceBeforePx + lines * lh + spaceAfterPx
    blocks.push({
      top: y,
      height,
      lineBoxes,
      spaceBeforePx,
      spaceAfterPx,
      section,
      docxIndex: i,
      ...(empty ? { emptyPara: true } : {}),
      ...(rnd() < 0.1 ? { keepNext: true } : {}),
      ...(rnd() < 0.03 ? { breakBefore: true } : {}),
      ...(rnd() < 0.02 ? { breakAfter: true } : {}),
      ...(rnd() < 0.01 ? { breakBefore: true, extraBreaksBefore: 1 } : {}),
      ...(rnd() < 0.1 ? { widowControl: false } : {}),
    })
    y += height
  }
  return { blocks, geoms, total: y }
}

const clone = <T>(x: T): T => JSON.parse(JSON.stringify(x)) as T

function slice(doc: Doc, resume?: PassResume): { slices: PageSlice[]; out: SliceOutputs } {
  const out: SliceOutputs = { rowFills: [], floatVShifts: [], oversizeClips: [] }
  const slices = sliceWithLineSplit(
    clone(doc.blocks),
    doc.geoms,
    doc.total,
    1,
    undefined,
    out,
    resume,
  )
  return { slices, out }
}

/** pages a pass may resume from: a distinct page opening on a block top */
function resumePoints(doc: Doc, pre: PageSlice[]): Array<{ page: number; block: number }> {
  const points: Array<{ page: number; block: number }> = []
  for (let p = 1; p < pre.length; p++) {
    if (!isResumePage(pre, p)) continue
    const page = pre[p]
    const k = doc.blocks.findIndex((b) => Math.abs(b.top - page.start) < 0.5)
    if (k < 0 || doc.blocks[k].floated) continue
    points.push({ page: p, block: k })
  }
  return points
}

describe('resumed slicing equals slicing the whole document', () => {
  it('from every resumable page of many random documents', () => {
    let resumed = 0
    for (let seed = 1; seed <= 60; seed++) {
      const doc = randomDoc(seed)
      const full = slice(doc)
      const pre = full.out.preParity!
      expect(pre.length).toBeGreaterThan(2)
      for (const { page, block } of resumePoints(doc, pre)) {
        const again = slice(doc, { prevSlices: pre, page, block, prevOut: full.out })
        expect(again.slices, `seed ${seed} page ${page}`).toEqual(full.slices)
        expect(again.out.preParity, `seed ${seed} page ${page} preParity`).toEqual(pre)
        resumed++
      }
    }
    expect(resumed).toBeGreaterThan(500)
  })

  it('a page sharing its start with an extra-break blank is not a resume point', () => {
    const lines = Array.from({ length: 5 }, (_, k) => ({ offsetInBlock: k * 20, height: 20 }))
    const blocks: BlockBox[] = Array.from({ length: 6 }, (_, i) => ({
      top: i * 100,
      height: 100,
      lineBoxes: lines,
      section: 0,
      docxIndex: i,
      ...(i === 3 ? { breakBefore: true, extraBreaksBefore: 1 } : {}),
    }))
    const doc = { blocks, geoms: [{ contentHeight: 250, forceBreak: true }], total: 600 }
    const pre = slice(doc).out.preParity!
    const blank = pre.findIndex((p, i) => i > 0 && p.end === p.start)
    expect(blank).toBeGreaterThan(0)
    expect(pre[blank + 1].start).toBe(pre[blank].start)
    expect(isResumePage(pre, blank)).toBe(false)
    expect(isResumePage(pre, blank + 1)).toBe(false)
    expect(isResumePage(pre, blank + 2)).toBe(true)
  })

  it('a title-page section resumed on its second page keeps the default capacity', () => {
    const lines = Array.from({ length: 10 }, (_, k) => ({ offsetInBlock: k * 20, height: 20 }))
    const blocks: BlockBox[] = Array.from({ length: 12 }, (_, i) => ({
      top: i * 200,
      height: 200,
      lineBoxes: lines,
      section: 0,
      docxIndex: i,
    }))
    const geoms: SectionGeom[] = [{ contentHeight: 400, firstContentHeight: 200, forceBreak: true }]
    const doc = { blocks, geoms, total: 2400 }
    const full = slice(doc)
    const pre = full.out.preParity!
    expect(pre[0].end - pre[0].start).toBe(200)
    expect(pre[1].end - pre[1].start).toBe(400)
    const again = slice(doc, { prevSlices: pre, page: 2, block: 3, prevOut: full.out })
    expect(again.slices).toEqual(full.slices)
  })

  it('keeps the previous pass patches for blocks above the resumed page', () => {
    const doc = randomDoc(7)
    const full = slice(doc)
    const pre = full.out.preParity!
    const { page, block } = resumePoints(doc, pre).at(-1)!
    const prevOut: SliceOutputs = {
      ...full.out,
      oversizeClips: [
        { blockTop: doc.blocks[0].top, clipPx: 3 },
        { blockTop: doc.blocks[block].top, clipPx: 5 },
      ],
    }
    const again = slice(doc, { prevSlices: pre, page, block, prevOut })
    expect(again.out.oversizeClips).toEqual([{ blockTop: doc.blocks[0].top, clipPx: 3 }])
  })
})

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
}

const para = (t: string): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content: [{ type: 'text', text: t }],
})

const liveEditors: Editor[] = []
afterEach(() => {
  for (const e of liveEditors.splice(0)) e.destroy()
})

function createEditor(content: JsonNode[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  liveEditors.push(editor)
  return editor
}

describe('firstChangedTopLevelIndex', () => {
  const three = () => createEditor([para('alpha'), para('beta'), para('gamma')])

  it('typing inside the second paragraph', () => {
    const editor = three()
    const pos = editor.state.doc.child(0).nodeSize + 3
    const tr = editor.state.tr.insertText('X', pos)
    expect(firstChangedTopLevelIndex(tr)).toBe(1)
  })

  it('splitting the first paragraph reports the first', () => {
    const editor = three()
    const end = editor.state.doc.child(0).nodeSize - 1
    const tr = editor.state.tr.setSelection(TextSelection.create(editor.state.doc, end)).split(end)
    expect(firstChangedTopLevelIndex(tr)).toBe(0)
  })

  it('joining the third paragraph onto the second reports the second', () => {
    const editor = three()
    const { doc } = editor.state
    const boundary = doc.child(0).nodeSize + doc.child(1).nodeSize
    const tr = editor.state.tr.join(boundary)
    expect(firstChangedTopLevelIndex(tr)).toBe(1)
  })

  it('a mark on the third paragraph reports the third', () => {
    const editor = three()
    const { doc } = editor.state
    const start = doc.child(0).nodeSize + doc.child(1).nodeSize + 1
    const tr = editor.state.tr.addMark(start, start + 3, editor.schema.marks.bold.create())
    expect(firstChangedTopLevelIndex(tr)).toBe(2)
  })

  it('a document attribute change is unknown', () => {
    const editor = three()
    const tr = editor.state.tr.setDocAttribute?.('x', 1) ?? editor.state.tr
    expect(firstChangedTopLevelIndex(tr)).toBeNull()
  })
})
