import { describe, expect, it, vi } from 'vitest'
import type { DeepLinkPayload, HomeApi } from '../src/shared/home-api'
import { HOME_CHANNELS } from '../src/shared/home-api'

/**
 * The preload's early-event buffers: a deep link or a palette summon can be
 * delivered right after did-finish-load, while the renderer is still waiting
 * on its locale/bootstrap round-trips and has not attached a handler yet. The
 * module-scope subscriptions must buffer those events and replay them on
 * attach — asynchronously, so the handler never runs inside the attach call
 * stack — and keep buffering after a detach (only the renderer handler goes
 * away, the module-scope subscription stays).
 */

type RendererListener = (event: unknown, ...args: unknown[]) => void

const electronMocks = vi.hoisted(() => {
  const listeners = new Map<string, Set<RendererListener>>()
  return {
    exposed: new Map<string, unknown>(),
    invoke: vi.fn(),
    send: vi.fn(),
    on: vi.fn((channel: string, listener: RendererListener) => {
      const set = listeners.get(channel) ?? new Set()
      set.add(listener)
      listeners.set(channel, set)
    }),
    removeListener: vi.fn((channel: string, listener: RendererListener) => {
      listeners.get(channel)?.delete(listener)
    }),
    listenerCount: vi.fn((channel: string) => listeners.get(channel)?.size ?? 0),
    /** deliver a main-process send on a channel to every current listener */
    emit: (channel: string, ...args: unknown[]) => {
      for (const listener of [...(listeners.get(channel) ?? [])]) listener({}, ...args)
    },
  }
})

vi.mock('electron', () => ({
  contextBridge: {
    exposeInMainWorld: (name: string, api: unknown) => electronMocks.exposed.set(name, api),
  },
  ipcRenderer: {
    invoke: electronMocks.invoke,
    send: electronMocks.send,
    on: electronMocks.on,
    removeListener: electronMocks.removeListener,
    listenerCount: electronMocks.listenerCount,
  },
  // the preload imports @genoffice/electron-utils (drop-open bridge), which
  // binds webUtils at module scope even though node env never installs it
  webUtils: { getPathForFile: () => '' },
}))

import '../src/preload/index'

const homeApi = electronMocks.exposed.get('aiOffice') as HomeApi

/** run the currently queued microtasks (the replays are queueMicrotask'd) */
async function flushMicrotasks(): Promise<void> {
  await Promise.resolve()
  await Promise.resolve()
}

describe('onDeepLink buffering', () => {
  it('buffers a deep link that arrives before any handler and replays it on attach', async () => {
    electronMocks.emit(HOME_CHANNELS.deepLink, { view: 'ask', query: 'hello from a link' })
    const seen: DeepLinkPayload[] = []
    const off = homeApi.onDeepLink((payload) => seen.push(payload))
    // not yet: the replay is asynchronous, never inside the attach call stack
    expect(seen).toEqual([])
    await flushMicrotasks()
    expect(seen).toEqual([{ view: 'ask', query: 'hello from a link' }])
    off()
  })

  it('queues several early links and replays them in arrival order', async () => {
    electronMocks.emit(HOME_CHANNELS.deepLink, { view: 'ask', query: 'first' })
    electronMocks.emit(HOME_CHANNELS.deepLink, { view: 'ask', query: 'second' })
    electronMocks.emit(HOME_CHANNELS.deepLink, { view: 'ask' })
    const seen: DeepLinkPayload[] = []
    const off = homeApi.onDeepLink((payload) => seen.push(payload))
    await flushMicrotasks()
    expect(seen).toEqual([
      { view: 'ask', query: 'first' },
      { view: 'ask', query: 'second' },
      { view: 'ask' },
    ])
    off()
  })

  it('drops non-ask and malformed payloads instead of buffering them', async () => {
    electronMocks.emit(HOME_CHANNELS.deepLink, { view: 'recent' })
    electronMocks.emit(HOME_CHANNELS.deepLink, 'ask')
    electronMocks.emit(HOME_CHANNELS.deepLink, null)
    const seen: DeepLinkPayload[] = []
    const off = homeApi.onDeepLink((payload) => seen.push(payload))
    await flushMicrotasks()
    expect(seen).toEqual([])
    off()
  })

  it('delivers directly once a handler is attached, and buffers again after detach', async () => {
    const seen: DeepLinkPayload[] = []
    const off = homeApi.onDeepLink((payload) => seen.push(payload))
    await flushMicrotasks()
    // attached: the event reaches the handler synchronously, unbuffered
    electronMocks.emit(HOME_CHANNELS.deepLink, { view: 'ask', query: 'direct' })
    expect(seen).toEqual([{ view: 'ask', query: 'direct' }])
    // only the renderer handler unsubscribes; the module-scope one stays
    off()
    expect(electronMocks.listenerCount(HOME_CHANNELS.deepLink)).toBe(1)
    electronMocks.emit(HOME_CHANNELS.deepLink, { view: 'ask', query: 'while away' })
    const seen2: DeepLinkPayload[] = []
    const off2 = homeApi.onDeepLink((payload) => seen2.push(payload))
    await flushMicrotasks()
    expect(seen2).toEqual([{ view: 'ask', query: 'while away' }])
    expect(seen).toEqual([{ view: 'ask', query: 'direct' }]) // no double delivery
    off2()
  })
})

describe('onPaletteOpen buffering', () => {
  it('replays a summon that arrived before attach, once, however many arrived', async () => {
    electronMocks.emit(HOME_CHANNELS.paletteOpen)
    electronMocks.emit(HOME_CHANNELS.paletteOpen)
    let opens = 0
    const off = homeApi.onPaletteOpen(() => {
      opens += 1
    })
    expect(opens).toBe(0) // asynchronous, never inside the attach call stack
    await flushMicrotasks()
    expect(opens).toBe(1)
    off()
  })

  it('delivers directly while attached and buffers again after detach', async () => {
    let opens = 0
    const off = homeApi.onPaletteOpen(() => {
      opens += 1
    })
    await flushMicrotasks()
    electronMocks.emit(HOME_CHANNELS.paletteOpen)
    expect(opens).toBe(1)
    off()
    expect(electronMocks.listenerCount(HOME_CHANNELS.paletteOpen)).toBe(1)
    electronMocks.emit(HOME_CHANNELS.paletteOpen)
    const seen: number[] = []
    const off2 = homeApi.onPaletteOpen(() => seen.push(1))
    await flushMicrotasks()
    expect(seen).toHaveLength(1)
    expect(opens).toBe(1) // no double delivery
    off2()
  })
})
