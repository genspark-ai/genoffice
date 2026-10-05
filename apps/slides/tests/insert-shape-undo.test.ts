import { tmpdir } from 'node:os'
import { beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'

const handlers = new Map<string, (...args: never[]) => unknown>()

vi.mock('electron', () => ({
  app: {
    getPath: () => tmpdir(),
    getName: () => 'slides-test',
    getVersion: () => '0',
    whenReady: () => Promise.resolve(),
  },
  dialog: {},
  clipboard: { writeBuffer: () => {} },
  ipcMain: {
    handle: (channel: string, fn: (...args: never[]) => unknown) => handlers.set(channel, fn),
    on: () => {},
    removeHandler: () => {},
  },
  BrowserWindow: class {},
  webContents: { getAllWebContents: () => [], fromId: () => null },
  Menu: { buildFromTemplate: () => ({}), setApplicationMenu: () => {} },
  session: { defaultSession: { protocol: {}, setDisplayMediaRequestHandler: () => {} } },
  nativeImage: {},
  shell: {},
  desktopCapturer: { getSources: async () => [] },
  WebContentsView: class {},
}))
// The shaping chain loads a harfbuzz wasm that cannot resolve in the unit-test env;
// nothing measured here renders text.
vi.mock('../src/main/shaped-metrics', () => ({
  shapedMetricsReady: async () => {},
  refineComplexWidths: async () => {},
  complexScriptOf: () => null,
  initShapedMetrics: () => {},
  gtMeasure: (t: string) => ({ width: t.length * 8 }),
  shapedFamily: () => null,
  shapedMeasure: () => ({ width: 0 }),
}))

import { createBlankPptx, openPptx } from '@genoffice/pptx-engine'
import { registerSlidesIpc } from '../src/main/slides-main'
import { sessions, type Session } from '../src/main/session-state'
import { insertShapeAt } from '../src/renderer/insert-actions'
import type { ActionCtx } from '../src/renderer/action-context'
import type { RenderSlide } from '@genoffice/pptx-render'

const WC = 1
const event = { sender: { id: WC } } as never

/** Forward a renderer call straight into the real main-process handler. */
function ipc(channel: string, ...args: unknown[]): unknown {
  const fn = handlers.get(channel)
  if (!fn) throw new Error(`handler not registered: ${channel}`)
  return (fn as (...a: unknown[]) => unknown)(event, ...args)
}

let session: Session
let applied: RenderSlide | null

beforeAll(() => {
  // registerSlidesIpc is guarded by a module-level flag, so it must run exactly once
  registerSlidesIpc()
})

beforeEach(async () => {
  sessions.clear()
  session = {
    path: '',
    fitWidthPx: 1280,
    undoStack: [],
    redoStack: [],
    opened: await openPptx(await createBlankPptx()),
  }
  sessions.set(WC, session)
  applied = null

  // Only the four calls insertShapeAt makes; each lands on the real handler, so
  // the history stack under test is the one the user would actually undo.
  vi.stubGlobal('window', {
    slidesApi: {
      addElement: (op: unknown) => ipc('slides:add-element', op),
      flipElements: (op: unknown) => ipc('slides:flip-elements', op),
      beginHistoryBatch: () => ipc('slides:history-batch-begin'),
      endHistoryBatch: () => ipc('slides:history-batch-end'),
    },
  })
})

function ctx(): ActionCtx {
  return {
    slide: session.opened.deck.slides[0],
    current: 0,
    applySlide: (_i: number, updated: RenderSlide) => {
      applied = updated
    },
    setSelectedIds: () => {},
  } as unknown as ActionCtx
}

function nodeCount(): number {
  return applied ? applied.nodes.length : 0
}

describe('drawing a line right-to-left is one undo step', () => {
  it('adds the line and both flips as a single undoable insert', async () => {
    await insertShapeAt(ctx(), 'lineArrow', { x: 400, y: 100, w: 200, h: 0, flipH: true })

    expect(nodeCount()).toBe(1)
    expect(session.undoStack).toHaveLength(1)

    ipc('slides:undo')

    // one undo must drop the line outright, not merely un-flip it
    expect(session.opened.deck.slides[0]!.elements).toHaveLength(0)
    expect(session.undoStack).toHaveLength(0)
  })

  it('collapses a bottom-to-top drag into one undo step too', async () => {
    await insertShapeAt(ctx(), 'lineArrow', { x: 100, y: 400, w: 0, h: 200, flipV: true })

    expect(nodeCount()).toBe(1)
    expect(session.undoStack).toHaveLength(1)
    ipc('slides:undo')
    expect(session.opened.deck.slides[0]!.elements).toHaveLength(0)
  })

  it('keeps a plain unflipped shape insert at one undo step', async () => {
    await insertShapeAt(ctx(), 'rect', { x: 10, y: 10, w: 100, h: 100 })

    expect(nodeCount()).toBe(1)
    expect(session.undoStack).toHaveLength(1)
    ipc('slides:undo')
    expect(session.opened.deck.slides[0]!.elements).toHaveLength(0)
  })
})
