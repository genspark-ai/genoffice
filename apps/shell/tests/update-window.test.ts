import { beforeEach, describe, expect, it, vi } from 'vitest'
import type { UpdateUiState } from '../src/shared/update-api'

/**
 * The settings-facing update surface (src/main/update-window.ts): the dialog
 * is non-modal, its state is broadcast to the shell window for the About row,
 * and open-for-update re-surfaces a minimized dialog, starting a download
 * that never began.
 */

type Handler = (event?: unknown, input?: unknown) => unknown

const handlers = new Map<string, Handler>()
const sent: Array<{ channel: string; state: UpdateUiState | null }> = []
const winOptions: Array<Record<string, unknown>> = []
const shown = vi.fn()
const focused = vi.fn()

class FakeWebContents {
  send(channel: string, state: UpdateUiState | null): void {
    sent.push({ channel, state })
  }
}

class FakeBrowserWindow {
  webContents = new FakeWebContents()
  constructor(options: Record<string, unknown>) {
    winOptions.push(options)
  }
  isDestroyed(): boolean {
    return false
  }
  show(): void {
    shown()
  }
  focus(): void {
    focused()
  }
  once(): void {}
  on(): void {}
  close(): void {}
  loadURL(): Promise<void> {
    return Promise.resolve()
  }
  loadFile(): Promise<void> {
    return Promise.resolve()
  }
}

vi.mock('electron', () => ({
  BrowserWindow: FakeBrowserWindow,
  ipcMain: {
    handle: (channel: string, handler: Handler) => handlers.set(channel, handler),
  },
}))

const state = (phase: UpdateUiState['phase']): UpdateUiState =>
  ({
    phase,
    version: '1.2.3',
    currentVersion: '1.2.2',
    percent: 0,
    lang: 'zh',
    strings: {
      title: 't',
      headline: 'h',
      desc: 'd',
      download: 'dl',
      later: 'min',
      install: 'in',
      downloading: 'ing',
      failed: 'f',
      retry: 'r',
      manualDesc: 'm',
      openDownload: 'o',
    },
  }) as UpdateUiState

const actions = () => ({
  onDownload: vi.fn(),
  onInstall: vi.fn(),
  onLater: vi.fn(),
  onOpenDownload: vi.fn(),
})

const parent = new FakeBrowserWindow() as unknown as Electron.BrowserWindow

beforeEach(() => {
  vi.resetModules()
  handlers.clear()
  sent.length = 0
  winOptions.length = 0
  shown.mockClear()
  focused.mockClear()
})

async function loadModule() {
  return import('../src/main/update-window')
}

describe('update window — settings-facing surface', () => {
  it('opens non-modal (never blocks the shell window)', async () => {
    const mod = await loadModule()
    mod.showUpdateWindow(parent, state('available'), actions())
    expect(winOptions[0]!.modal).toBeUndefined()
    expect(winOptions[0]!.parent).toBe(parent)
  })

  it('broadcasts every state push to the shell window', async () => {
    const mod = await loadModule()
    mod.showUpdateWindow(parent, state('downloading'), actions())
    mod.pushUpdateState({ percent: 42 })
    const pushes = sent.filter((s) => s.channel === 'update:state-changed')
    expect(pushes).toHaveLength(2)
    expect(pushes[1]!.state!.percent).toBe(42)
  })

  it('open-for-update surfaces the dialog and starts a download that never began', async () => {
    const mod = await loadModule()
    const act = actions()
    mod.showUpdateWindow(parent, state('available'), act)
    // the dialog was minimized (window gone); About re-opens it
    mod.closeUpdateWindow()
    sent.length = 0
    const ok = handlers.get('update:open-for-update')!()
    expect(ok).toBe(true)
    expect(act.onDownload).toHaveBeenCalledTimes(1)
    // re-opening also re-broadcasts for the About row
    expect(sent.some((s) => s.channel === 'update:state-changed')).toBe(true)
  })

  it('open-for-update never restarts an in-flight or finished download', async () => {
    const mod = await loadModule()
    for (const phase of ['downloading', 'downloaded'] as const) {
      const act = actions()
      mod.showUpdateWindow(parent, state(phase), act)
      handlers.get('update:open-for-update')!()
      expect(act.onDownload).not.toHaveBeenCalled()
      mod.closeUpdateWindow()
    }
  })

  it('open-for-update reports false while no update is known', async () => {
    // Settings → About invokes these on mount, so the handlers have to answer
    // on a fresh module with no dialog ever shown — that is the common case
    // (a fresh launch with nothing known). Asserting `handlers.has(...)` only
    // proves a registration call ran, not that the invoke resolves, which is
    // how this shipped broken: the channel existed, but only after a dialog
    // had opened, so the real invoke rejected with "No handler registered".
    const mod = await loadModule()
    expect(mod.currentUpdateUiState()).toBeNull()
    expect(handlers.get('update:get-state')!()).toBeNull()
    expect(handlers.get('update:open-for-update')!()).toBe(false)
    // the pre-fix bug threw here: handlers.get(...) was undefined
    expect(handlers.get('update:open-for-update')).toBeTypeOf('function')
  })

  it('resolves the shell window through the getter, not the cached parent', async () => {
    const mod = await loadModule()
    mod.showUpdateWindow(parent, state('available'), actions())
    mod.closeUpdateWindow()
    const recreated = new FakeBrowserWindow({}) as unknown as Electron.BrowserWindow
    mod.setUpdateParentWindow(() => recreated)
    winOptions.length = 0
    expect(handlers.get('update:open-for-update')!()).toBe(true)
    expect(winOptions[0]!.parent).toBe(recreated)
  })

  it('a channel switch clears the known update and tells the About row', async () => {
    const mod = await loadModule()
    mod.showUpdateWindow(parent, state('downloaded'), actions())
    sent.length = 0
    mod.clearUpdateState()
    expect(mod.currentUpdateUiState()).toBeNull()
    expect(handlers.get('update:get-state')!()).toBeNull()
    expect(handlers.get('update:open-for-update')!()).toBe(false)
    const last = sent.filter((s) => s.channel === 'update:state-changed').at(-1)
    expect(last?.state).toBeNull()
  })

  it('exposes the freshest state for the settings row', async () => {
    const mod = await loadModule()
    expect(mod.currentUpdateUiState()).toBeNull()
    const s = state('available')
    mod.showUpdateWindow(parent, s, actions())
    expect(mod.currentUpdateUiState()).toBe(s)
  })
})
