/**
 * Regression benchmark for the 122-page open.
 *
 * The report that prompted the hardening measured a 122-page document at 97
 * seconds to first content. Three heuristics now stand between that number
 * and the current one — a progressive first mount, a duty-cycled tail, and a
 * carry that lets each streaming pagination pass reuse the previous pass's
 * line samples instead of re-deriving them for the whole document. None of
 * the three was pinned by a test, so any later change could undo one of them
 * and the suite would stay green.
 *
 * The bound asserted here is a work count, not wall-clock, on purpose. jsdom
 * lays out nothing (the same caveat pagination-carry-samples.test.ts records),
 * so a timer in this suite would measure ProseMirror bookkeeping instead of
 * the layout and pagination that dominate those 97 seconds — and it would be
 * at the mercy of how busy the runner is besides. What *is* deterministic here
 * is exactly what the heuristics claim to bound: how many blocks the first
 * mount touches, how much each streaming step does, and how much of the
 * document a streaming pagination pass has to re-sample.
 *
 * 122 pages is the corpus the report used. Its block count follows the figure
 * that commit recorded for it — a 122-page document "lands in ~10 chunks" —
 * which fixes it at PHASE1_BLOCKS + 10 * PHASE_CHUNK_BLOCKS = 1344 blocks.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { PHASED_APPEND } from '../src/renderer/editor/streaming-tail-guard'
import { appendStreamedNodes } from '../src/renderer/file-actions'
import { carryStreamedSamples, fillLineBoxes } from '../src/renderer/pagination-lines'
import type { BlockBox, SectionGeom } from '../src/renderer/pagination-types'
import {
  PHASE1_BLOCKS,
  PHASE_CHUNK_BLOCKS,
  cancelPhasedContent,
  isPhasedContentPending,
  setContentPhased,
  type PhasedContentHost,
} from '../src/renderer/phased-content'

/** top-level blocks in the reported 122-page document: 64 + 10 * 128 */
const PAGES_122_BLOCKS = PHASE1_BLOCKS + 10 * PHASE_CHUNK_BLOCKS
/** the tail of that document arrives in this many appended chunks */
const PAGES_122_CHUNKS = 10
/** A4-ish content area; the benchmark counts sampled blocks, not pages */
const GEOMS: SectionGeom[] = [{ contentHeight: 10_000, forceBreak: false }]

const para = (i: number) => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content: [{ type: 'text', text: `block ${i}` }],
})
const docOf = (blocks: number) => ({
  type: 'doc',
  content: Array.from({ length: blocks }, (_, i) => para(i)),
})

const liveEditors: Editor[] = []
afterEach(() => {
  cancelPhasedContent()
  for (const e of liveEditors.splice(0)) e.destroy()
})

function createEditor(): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content: [para(0)] },
  })
  liveEditors.push(editor)
  return editor
}

/** mounts the first phase and returns the scheduled tail chunks */
function startStreamingOpen(editor: Editor, blocks: number): Array<() => void> {
  const chunks: Array<() => void> = []
  const host: PhasedContentHost = {
    setContent: (doc) =>
      editor
        .chain()
        .setMeta('addToHistory', false)
        .setMeta(PHASED_APPEND, true)
        .setContent(doc as never)
        .run(),
    appendNodes: (nodes) => appendStreamedNodes(editor, nodes as never),
    isDestroyed: () => editor.isDestroyed,
    resetHistory: () => {},
    setLoading: () => {},
    getDirty: () => false,
    setDirty: () => {},
  }
  setContentPhased(host, docOf(blocks) as never, (cb) => chunks.push(cb))
  return chunks
}

const rectOf = (top: number, height: number, width = 600) =>
  ({
    top,
    height,
    bottom: top + height,
    left: 0,
    right: width,
    width,
    x: 0,
    y: top,
    toJSON: () => ({}),
  }) as DOMRect

/**
 * The freshly measured blocks of one pagination pass, built from the editor's
 * real DOM children with synthetic page geometry. jsdom returns no rects, so
 * the geometry is stubbed; the block *count* is what the pass pays for and it
 * is real. Height is a function of the block's own index, so a block's
 * geometry is stable across passes exactly as an unchanged block's would be.
 */
function measurePass(editor: Editor, top: number): BlockBox[] {
  const els = Array.from(editor.view.dom.children) as HTMLElement[]
  expect(els.length).toBe(editor.state.doc.childCount)
  let y = top
  return els.map((el, i) => {
    const height = 60 + (i % 4) * 12
    el.getBoundingClientRect = () => rectOf(y, height)
    const box = { top: y, height, el, widthPx: 600 } as BlockBox
    y += height
    return box
  })
}

/** one line sample, as a real layout pass would have produced */
const sample = [{ offsetInBlock: 0, height: 14 }] as unknown as BlockBox['lineBoxes']

/**
 * Runs the pass the streaming pipeline runs, and reports how many blocks it
 * had to sample. Blocks that arrive carrying the previous pass's samples are
 * skipped by fillLineBoxes; the rest are the pass's work.
 */
function paginationPass(blocks: BlockBox[], prev: BlockBox[], pending: boolean): number {
  carryStreamedSamples(blocks, prev, {
    pending,
    dirty: prev.length,
    lastPassChildCount: prev.length,
  })
  const resampled = blocks.filter((b) => !b.lineBoxes).length
  fillLineBoxes(blocks, GEOMS, 1)
  // stand in for the line boxes a real layout pass derives, so the next pass
  // has samples to carry
  for (const b of blocks) if (!b.lineBoxes) b.lineBoxes = sample
  return resampled
}

describe('the 122-page phased open', () => {
  it('mounts a bounded first screen, streams the tail in bounded steps, and samples the document once', () => {
    const editor = createEditor()
    const chunks = startStreamingOpen(editor, PAGES_122_BLOCKS)

    // First content is bounded and does not grow with the document: the whole
    // 122 pages are not built to show page 1.
    expect(editor.state.doc.childCount).toBe(PHASE1_BLOCKS)
    // duty-cycled: the tail waits for a painted frame, it does not run on in
    // the same task as the first mount
    expect(chunks).toHaveLength(1)
    expect(isPhasedContentPending()).toBe(true)

    // a pagination pass over the first screen
    const perPass: number[] = []
    let prev = measurePass(editor, 0)
    perPass.push(paginationPass(prev, [], false))

    // every streamed chunk is followed by one pass, and each pass samples only
    // the chunk that just arrived
    while (chunks.length) {
      chunks.shift()!()
      const blocks = measurePass(editor, 0)
      perPass.push(paginationPass(blocks, prev, isPhasedContentPending()))
      prev = blocks
    }

    expect(isPhasedContentPending()).toBe(false)
    // the tail landed in the chunk count the fix recorded for this corpus
    expect(perPass).toHaveLength(1 + PAGES_122_CHUNKS)
    expect(editor.state.doc.childCount).toBe(PAGES_122_BLOCKS)

    // The bound. While the tail streams, no pass may cost more than the chunk
    // that caused it — the carried samples are what make a streaming pass a
    // tail-sized pass rather than a document-sized one. Only the settled pass
    // measures the whole document, and the settled event asks for exactly that
    // once; without the carry every one of these passes would cost all
    // PAGES_122_BLOCKS, which is the regression this benchmark exists to catch.
    const streaming = perPass.slice(1, -1)
    for (const resampled of streaming) expect(resampled).toBeLessThanOrEqual(PHASE_CHUNK_BLOCKS)
    expect(Math.max(...streaming)).toBeLessThan(PAGES_122_BLOCKS / 10)
    expect(perPass.at(-1)).toBe(PAGES_122_BLOCKS)
    expect(perPass[0]).toBe(PHASE1_BLOCKS)
    // exactly one whole-document pass in the whole open, and it is the last
    expect(perPass.filter((n) => n > PHASE_CHUNK_BLOCKS)).toHaveLength(1)

    // so the open costs one document's worth of sampling plus the streamed
    // chunks — not one document's worth per pass
    const total = perPass.reduce((a, b) => a + b, 0)
    expect(total).toBe(
      PAGES_122_BLOCKS + PHASE1_BLOCKS + (PAGES_122_CHUNKS - 1) * PHASE_CHUNK_BLOCKS,
    )

    // and the chunking did not cost content: every block landed, in order
    const last = editor.state.doc.childCount - 1
    expect(editor.state.doc.child(last).textContent).toBe(`block ${last}`)
  })

  it('mounts the same first screen whatever the document size', () => {
    const small = createEditor()
    startStreamingOpen(small, PAGES_122_BLOCKS)
    const large = createEditor()
    startStreamingOpen(large, PAGES_122_BLOCKS * 4)
    // four times the pages must not cost four times the first screen: this is
    // the property that took the 97 seconds off first content
    expect(large.state.doc.childCount).toBe(small.state.doc.childCount)
    expect(small.state.doc.childCount).toBe(PHASE1_BLOCKS)
  })
})
