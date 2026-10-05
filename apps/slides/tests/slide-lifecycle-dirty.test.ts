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
  session: {
    defaultSession: { protocol: {}, setDisplayMediaRequestHandler: () => {} },
  },
  nativeImage: {},
  shell: {},
  desktopCapturer: { getSources: async () => [] },
  WebContentsView: class {},
}))
// The shaping chain loads a harfbuzz wasm that cannot resolve in the unit-test env;
// none of the handlers under test measure text.
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
import { registerSlidesIpc, slidesIsDirty } from '../src/main/slides-main'
import { sessions, type Session } from '../src/main/session-state'

const WC = 1
const FIT = 1280
const event = { sender: { id: WC } } as never

function call(channel: string, ...args: unknown[]): unknown {
  const fn = handlers.get(channel)
  if (!fn) throw new Error(`handler not registered: ${channel}`)
  return (fn as (...a: unknown[]) => unknown)(event, ...args)
}

let session: Session

/** The real predicate behind the close guard, the 30s autosave tick and the AutoSave poll. */
function dirty(): boolean {
  return slidesIsDirty(WC)
}

beforeAll(() => {
  // registerSlidesIpc is guarded by a module-level flag, so it must run exactly once
  registerSlidesIpc()
})

beforeEach(async () => {
  sessions.clear()
  session = {
    path: '',
    fitWidthPx: FIT,
    undoStack: [],
    redoStack: [],
    opened: await openPptx(await createBlankPptx()),
  }
  sessions.set(WC, session)
})

function firstLayoutPath(): string {
  const r = call('slides:get-layouts') as { layouts: { path: string }[] } | null
  const path = r?.layouts[0]?.path
  if (!path) throw new Error('blank deck exposed no slide layout')
  return path
}

describe('slide lifecycle marks the session dirty', () => {
  it('insertBlankSlide leaves the session dirty', () => {
    expect(dirty()).toBe(false)
    call('slides:add-blank-slide', { sourceIndex: 0, before: false, fitWidthPx: FIT })
    expect(session.opened.deck.slides).toHaveLength(2)
    expect(dirty()).toBe(true)
  })

  it('insertSlideWithLayout leaves the session dirty', () => {
    const layoutPath = firstLayoutPath()
    expect(dirty()).toBe(false)
    call('slides:add-slide-with-layout', { sourceIndex: 0, layoutPath, fitWidthPx: FIT })
    expect(session.opened.deck.slides).toHaveLength(2)
    expect(dirty()).toBe(true)
  })

  it('duplicateSlide leaves the session dirty', () => {
    expect(dirty()).toBe(false)
    call('slides:add-slide', { sourceIndex: 0, fitWidthPx: FIT })
    expect(session.opened.deck.slides).toHaveLength(2)
    expect(dirty()).toBe(true)
  })

  it('pasteSlide leaves the session dirty', () => {
    call('slides:copy-slides', { slideIndexes: [0] })
    expect(dirty()).toBe(false)
    call('slides:paste-slide', { afterIndex: 0, fitWidthPx: FIT })
    expect(session.opened.deck.slides).toHaveLength(2)
    expect(dirty()).toBe(true)
  })

  it('deleteSlide leaves the session dirty', () => {
    call('slides:add-blank-slide', { sourceIndex: 0, before: false, fitWidthPx: FIT })
    // the insert above already flagged the session; a clean deck proves the delete is what flags it
    session.metaDirty = false
    expect(dirty()).toBe(false)
    call('slides:delete-slide', 0)
    expect(session.opened.deck.slides).toHaveLength(1)
    expect(dirty()).toBe(true)
  })
})
