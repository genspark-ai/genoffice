import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  PAINTED_FRAME_FALLBACK_MS,
  PHASE1_BLOCKS,
  PHASE_CHUNK_BLOCKS,
  cancelPhasedContent,
  isPhasedContentPending,
  setContentPhased,
  type PhasedContentHost,
} from '../src/renderer/phased-content'
import type { PmNode } from '../src/renderer/editor/convert'

const blocks = (n: number): PmNode[] =>
  Array.from({ length: n }, (_, i) => ({ type: 'paragraph', attrs: { docxIndex: i } }) as PmNode)

function host(): PhasedContentHost & { mounted: number } {
  const h = {
    mounted: 0,
    setContent(doc: PmNode) {
      h.mounted = doc.content?.length ?? 0
    },
    appendNodes(nodes: PmNode[]) {
      h.mounted += nodes.length
    },
    isDestroyed: () => false,
    resetHistory() {},
    setLoading() {},
    getDirty: () => false,
    setDirty() {},
  }
  return h
}

describe('phased open in a background view', () => {
  beforeEach(() => {
    vi.useFakeTimers()
  })
  afterEach(() => {
    cancelPhasedContent()
    vi.useRealTimers()
    vi.unstubAllGlobals()
  })

  it('keeps streaming on the timer fallback when frames never come', () => {
    // a hidden shell view is handed a frame a second at best; here none at all
    vi.stubGlobal('requestAnimationFrame', () => 1)
    vi.stubGlobal('cancelAnimationFrame', () => {})
    const h = host()
    const total = PHASE1_BLOCKS + PHASE_CHUNK_BLOCKS * 3
    setContentPhased(h, { type: 'doc', content: blocks(total) } as PmNode)
    expect(h.mounted).toBe(PHASE1_BLOCKS)
    expect(isPhasedContentPending()).toBe(true)
    vi.advanceTimersByTime(PAINTED_FRAME_FALLBACK_MS)
    expect(h.mounted).toBe(PHASE1_BLOCKS + PHASE_CHUNK_BLOCKS)
    vi.advanceTimersByTime(PAINTED_FRAME_FALLBACK_MS * 2)
    expect(h.mounted).toBe(total)
    expect(isPhasedContentPending()).toBe(false)
  })

  it('appends a chunk once when the frame wins the race', () => {
    const frames: FrameRequestCallback[] = []
    vi.stubGlobal('requestAnimationFrame', (cb: FrameRequestCallback) => frames.push(cb))
    vi.stubGlobal('cancelAnimationFrame', () => {})
    const h = host()
    const total = PHASE1_BLOCKS + PHASE_CHUNK_BLOCKS * 2
    setContentPhased(h, { type: 'doc', content: blocks(total) } as PmNode)
    // two frames: the paint frame, then the append
    frames.shift()!(0)
    frames.shift()!(0)
    expect(h.mounted).toBe(PHASE1_BLOCKS + PHASE_CHUNK_BLOCKS)
    // the first chunk's losing timer must not append it again; the timer of
    // the next race lands the last chunk exactly once
    vi.advanceTimersByTime(PAINTED_FRAME_FALLBACK_MS * 3)
    expect(h.mounted).toBe(total)
    expect(isPhasedContentPending()).toBe(false)
  })
})
