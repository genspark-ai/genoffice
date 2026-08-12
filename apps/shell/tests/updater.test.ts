import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { UpdateUiState } from '../src/shared/update-api'

/**
 * Main-process auto-updater flow (src/main/updater.ts): manual-download policy
 * driven by the strong-guidance update window — the renderer's Download /
 * Install / Later actions come back through update-window.ts callbacks.
 */

const appState = { isPackaged: true }

const openExternal = vi.hoisted(() => vi.fn())
const readFileSyncMock = vi.hoisted(() => vi.fn<(...args: unknown[]) => string>())

vi.mock('electron', () => ({
  app: {
    get isPackaged() {
      return appState.isPackaged
    },
    getVersion: () => '0.1.0',
  },
  shell: {
    openExternal: (url: string) => openExternal(url),
  },
}))

vi.mock('node:fs', () => ({
  readFileSync: (...args: unknown[]) => readFileSyncMock(...args),
}))

type Listener = (...args: unknown[]) => unknown

const updaterState = {
  listeners: new Map<string, Listener>(),
  autoDownload: true,
  autoInstallOnAppQuit: false,
  disableDifferentialDownload: false,
  channel: null as string | null,
  allowDowngrade: false,
}
const checkForUpdates = vi.fn(() => Promise.resolve(null))
const downloadUpdate = vi.fn<() => Promise<unknown>>(() => Promise.resolve([]))
const quitAndInstall = vi.fn()

vi.mock('electron-updater', () => ({
  autoUpdater: {
    on: (event: string, listener: Listener) => {
      updaterState.listeners.set(event, listener)
    },
    checkForUpdates: () => checkForUpdates(),
    downloadUpdate: () => downloadUpdate(),
    quitAndInstall: (...args: unknown[]) => quitAndInstall(...(args as [boolean, boolean])),
    set autoDownload(v: boolean) {
      updaterState.autoDownload = v
    },
    set autoInstallOnAppQuit(v: boolean) {
      updaterState.autoInstallOnAppQuit = v
    },
    set disableDifferentialDownload(v: boolean) {
      updaterState.disableDifferentialDownload = v
    },
    set channel(v: string | null) {
      updaterState.channel = v
      // mirrors electron-updater's real setter side effect: assigning
      // `channel` unconditionally flips allowDowngrade to true
      updaterState.allowDowngrade = true
    },
    get channel() {
      return updaterState.channel
    },
    set allowDowngrade(v: boolean) {
      updaterState.allowDowngrade = v
    },
    get allowDowngrade() {
      return updaterState.allowDowngrade
    },
  },
}))

interface UpdateActions {
  onDownload: () => void
  onInstall: () => void
  onLater: () => void
  onOpenDownload: () => void
}

const showUpdateWindow =
  vi.fn<(parent: unknown, state: UpdateUiState, actions: UpdateActions) => void>()
const pushUpdateState = vi.fn<(patch: Partial<UpdateUiState>) => void>()
const closeUpdateWindow = vi.fn()

vi.mock('../src/main/update-window', () => ({
  showUpdateWindow: (...args: [unknown, UpdateUiState, UpdateActions]) => showUpdateWindow(...args),
  pushUpdateState: (patch: Partial<UpdateUiState>) => pushUpdateState(patch),
  closeUpdateWindow: () => closeUpdateWindow(),
}))

const FIRST_CHECK_DELAY_MS = 15_000
const RECHECK_INTERVAL_MS = 4 * 60 * 60 * 1000

let platformSpy: { restore: () => void } | null = null

function setPlatform(platform: string): void {
  const original = Object.getOwnPropertyDescriptor(process, 'platform')!
  Object.defineProperty(process, 'platform', { value: platform })
  platformSpy = {
    restore: () => Object.defineProperty(process, 'platform', original),
  }
}

async function loadUpdater() {
  return import('../src/main/updater')
}

async function flushAsync(): Promise<void> {
  // drain promise chains queued by download callbacks
  for (let i = 0; i < 10; i++) await Promise.resolve()
}

function lastShownState(): UpdateUiState {
  return showUpdateWindow.mock.calls.at(-1)![1]
}

function lastShownActions(): UpdateActions {
  return showUpdateWindow.mock.calls.at(-1)![2]
}

beforeEach(() => {
  vi.resetModules()
  vi.useFakeTimers()
  appState.isPackaged = true
  delete process.env.GENOFFICE_FAKE_UPDATE
  updaterState.listeners.clear()
  updaterState.autoDownload = true
  updaterState.autoInstallOnAppQuit = false
  updaterState.disableDifferentialDownload = false
  updaterState.channel = null
  updaterState.allowDowngrade = false
  checkForUpdates.mockClear()
  downloadUpdate.mockReset()
  downloadUpdate.mockImplementation(() => Promise.resolve([]))
  quitAndInstall.mockClear()
  showUpdateWindow.mockClear()
  pushUpdateState.mockClear()
  closeUpdateWindow.mockClear()
  openExternal.mockClear()
  readFileSyncMock.mockReset()
  readFileSyncMock.mockImplementation(() => {
    throw new Error('no app-update.yml')
  })
  setPlatform('darwin')
})

afterEach(() => {
  vi.useRealTimers()
  platformSpy?.restore()
  platformSpy = null
  delete process.env.GENOFFICE_FAKE_UPDATE
})

describe('initAutoUpdater', () => {
  it('does nothing in unpacked (dev) runs without the fake-update env', async () => {
    appState.isPackaged = false
    const { initAutoUpdater } = await loadUpdater()
    initAutoUpdater(() => null)
    vi.advanceTimersByTime(FIRST_CHECK_DELAY_MS)
    expect(updaterState.listeners.size).toBe(0)
    expect(checkForUpdates).not.toHaveBeenCalled()
    expect(showUpdateWindow).not.toHaveBeenCalled()
  })

  it('does nothing on unsupported platforms', async () => {
    platformSpy?.restore()
    setPlatform('linux')
    const { initAutoUpdater } = await loadUpdater()
    initAutoUpdater(() => null)
    vi.advanceTimersByTime(FIRST_CHECK_DELAY_MS)
    expect(updaterState.listeners.size).toBe(0)
    expect(checkForUpdates).not.toHaveBeenCalled()
  })

  it('configures manual full-package download and checks after the initial delay', async () => {
    const { initAutoUpdater } = await loadUpdater()
    initAutoUpdater(() => null)
    expect(updaterState.autoDownload).toBe(false)
    expect(updaterState.autoInstallOnAppQuit).toBe(true)
    expect(updaterState.disableDifferentialDownload).toBe(true)
    expect(checkForUpdates).not.toHaveBeenCalled()
    vi.advanceTimersByTime(FIRST_CHECK_DELAY_MS)
    expect(checkForUpdates).toHaveBeenCalledTimes(1)
  })

  it('re-checks on the periodic interval', async () => {
    const { initAutoUpdater } = await loadUpdater()
    initAutoUpdater(() => null)
    vi.advanceTimersByTime(FIRST_CHECK_DELAY_MS)
    expect(checkForUpdates).toHaveBeenCalledTimes(1)
    vi.advanceTimersByTime(RECHECK_INTERVAL_MS)
    expect(checkForUpdates).toHaveBeenCalledTimes(2)
  })

  it('is idempotent across repeated init calls', async () => {
    const { initAutoUpdater } = await loadUpdater()
    initAutoUpdater(() => null)
    initAutoUpdater(() => null)
    vi.advanceTimersByTime(FIRST_CHECK_DELAY_MS)
    expect(checkForUpdates).toHaveBeenCalledTimes(1)
  })

  it('opens the update window with the available state when an update is found', async () => {
    const { initAutoUpdater } = await loadUpdater()
    initAutoUpdater(() => null)
    updaterState.listeners.get('update-available')!({ version: '0.2.0' })
    expect(showUpdateWindow).toHaveBeenCalledTimes(1)
    const state = lastShownState()
    expect(state.phase).toBe('available')
    expect(state.version).toBe('0.2.0')
    expect(state.currentVersion).toBe('0.1.0')
    expect(state.percent).toBe(0)
    // strings are localized main-side and pushed into the window state
    expect(state.strings.title.length).toBeGreaterThan(0)
    expect(state.strings.install.length).toBeGreaterThan(0)
  })

  it('starts the download and pushes progress when the user clicks download', async () => {
    const { initAutoUpdater } = await loadUpdater()
    initAutoUpdater(() => null)
    updaterState.listeners.get('update-available')!({ version: '0.2.0' })
    lastShownActions().onDownload()
    expect(pushUpdateState).toHaveBeenCalledWith({ phase: 'downloading', percent: 0 })
    expect(downloadUpdate).toHaveBeenCalledTimes(1)

    updaterState.listeners.get('download-progress')!({ percent: 42 })
    expect(pushUpdateState).toHaveBeenCalledWith({ phase: 'downloading', percent: 42 })

    updaterState.listeners.get('update-downloaded')!({ version: '0.2.0' })
    expect(pushUpdateState).toHaveBeenCalledWith({ phase: 'downloaded', percent: 100 })
  })

  it('surfaces a failed download as the error phase', async () => {
    downloadUpdate.mockImplementation(() => Promise.reject(new Error('offline')))
    const { initAutoUpdater } = await loadUpdater()
    initAutoUpdater(() => null)
    updaterState.listeners.get('update-available')!({ version: '0.2.0' })
    lastShownActions().onDownload()
    await flushAsync()
    expect(pushUpdateState).toHaveBeenCalledWith({ phase: 'error' })
  })

  it('closes the window and installs on restart when the user confirms', async () => {
    const { initAutoUpdater } = await loadUpdater()
    initAutoUpdater(() => null)
    updaterState.listeners.get('update-available')!({ version: '0.2.0' })
    lastShownActions().onInstall()
    expect(closeUpdateWindow).toHaveBeenCalledTimes(1)
    // quitAndInstall is deferred so the window can finish closing first
    expect(quitAndInstall).not.toHaveBeenCalled()
    vi.advanceTimersByTime(0)
    expect(quitAndInstall).toHaveBeenCalledWith(true, true)
  })

  it('does not re-prompt for a version the user dismissed this session', async () => {
    const { initAutoUpdater } = await loadUpdater()
    initAutoUpdater(() => null)
    const available = updaterState.listeners.get('update-available')!
    available({ version: '0.2.0' })
    expect(showUpdateWindow).toHaveBeenCalledTimes(1)
    lastShownActions().onLater()
    expect(closeUpdateWindow).toHaveBeenCalledTimes(1)

    available({ version: '0.2.0' })
    expect(showUpdateWindow).toHaveBeenCalledTimes(1)

    // a newer version prompts again
    available({ version: '0.3.0' })
    expect(showUpdateWindow).toHaveBeenCalledTimes(2)
  })

  it('swallows background check failures', async () => {
    checkForUpdates.mockImplementation(() => Promise.reject(new Error('offline')))
    const { initAutoUpdater } = await loadUpdater()
    initAutoUpdater(() => null)
    vi.advanceTimersByTime(FIRST_CHECK_DELAY_MS)
    await flushAsync()
    expect(showUpdateWindow).not.toHaveBeenCalled()
  })

  it('defaults to the stable (latest) feed channel', async () => {
    const { initAutoUpdater } = await loadUpdater()
    initAutoUpdater(() => null)
    expect(updaterState.channel).toBe('latest')
    // must never allow a downgrade, despite electron-updater's channel setter
    // side effect
    expect(updaterState.allowDowngrade).toBe(false)
  })

  it('starts on the beta feed when the beta channel is passed in', async () => {
    const { initAutoUpdater } = await loadUpdater()
    initAutoUpdater(() => null, 'beta')
    expect(updaterState.channel).toBe('beta')
    // must never allow a downgrade, despite electron-updater's channel setter
    // side effect
    expect(updaterState.allowDowngrade).toBe(false)
  })

  it('applyUpdateChannel switches the feed and re-checks immediately', async () => {
    const { initAutoUpdater, applyUpdateChannel } = await loadUpdater()
    initAutoUpdater(() => null)
    expect(checkForUpdates).not.toHaveBeenCalled()
    applyUpdateChannel('beta')
    expect(updaterState.channel).toBe('beta')
    expect(updaterState.allowDowngrade).toBe(false)
    expect(checkForUpdates).toHaveBeenCalledTimes(1)
    applyUpdateChannel('stable')
    expect(updaterState.channel).toBe('latest')
    expect(updaterState.allowDowngrade).toBe(false)
    expect(checkForUpdates).toHaveBeenCalledTimes(2)
  })

  it('applyUpdateChannel is a no-op before the real updater is active', async () => {
    appState.isPackaged = false
    const { initAutoUpdater, applyUpdateChannel } = await loadUpdater()
    initAutoUpdater(() => null)
    applyUpdateChannel('beta')
    expect(updaterState.channel).toBeNull()
    expect(checkForUpdates).not.toHaveBeenCalled()
  })
})

describe('manual download fallback', () => {
  const macFiles = [
    { url: 'GenOffice-0.2.0-arm64.zip' },
    { url: 'GenOffice-0.2.0.zip' },
    { url: 'GenOffice-0.2.0-arm64.dmg' },
    { url: 'GenOffice-0.2.0.dmg' },
  ]

  function setArch(arch: string): () => void {
    const original = Object.getOwnPropertyDescriptor(process, 'arch')!
    Object.defineProperty(process, 'arch', { value: arch })
    return () => Object.defineProperty(process, 'arch', original)
  }

  async function failTwiceIntoManual(files: { url: string }[]): Promise<UpdateActions> {
    downloadUpdate.mockImplementation(() => Promise.reject(new Error('team id changed')))
    const { initAutoUpdater } = await loadUpdater()
    initAutoUpdater(() => null)
    updaterState.listeners.get('update-available')!({ version: '0.2.0', files })
    const actions = lastShownActions()
    actions.onDownload()
    await flushAsync()
    expect(pushUpdateState).toHaveBeenCalledWith({ phase: 'error' })
    actions.onDownload()
    await flushAsync()
    expect(pushUpdateState).toHaveBeenCalledWith({ phase: 'manual' })
    return actions
  }

  it('opens the feed-derived installer matching channel base URL and arm64 arch', async () => {
    Object.defineProperty(process, 'resourcesPath', { value: '/res', configurable: true })
    readFileSyncMock.mockReturnValue('provider: generic\nurl: https://cdn.example.com/mac/\n')
    const restoreArch = setArch('arm64')
    try {
      const actions = await failTwiceIntoManual(macFiles)
      actions.onOpenDownload()
      expect(openExternal).toHaveBeenCalledWith(
        'https://cdn.example.com/mac/GenOffice-0.2.0-arm64.dmg',
      )
    } finally {
      restoreArch()
    }
  })

  it('picks the arch-less dmg for x64 installs', async () => {
    Object.defineProperty(process, 'resourcesPath', { value: '/res', configurable: true })
    readFileSyncMock.mockReturnValue('url: https://cdn.example.com/mac\n')
    const restoreArch = setArch('x64')
    try {
      const actions = await failTwiceIntoManual(macFiles)
      actions.onOpenDownload()
      expect(openExternal).toHaveBeenCalledWith('https://cdn.example.com/mac/GenOffice-0.2.0.dmg')
    } finally {
      restoreArch()
    }
  })

  it('rebuilds absolute metadata URLs against the trusted feed base', async () => {
    Object.defineProperty(process, 'resourcesPath', { value: '/res', configurable: true })
    readFileSyncMock.mockReturnValue('url: https://cdn.example.com/mac\n')
    const restoreArch = setArch('arm64')
    try {
      const actions = await failTwiceIntoManual([
        { url: 'https://attacker.example/GenOffice-0.2.0-arm64.zip' },
        { url: 'https://attacker.example/GenOffice-0.2.0-arm64.dmg' },
        { url: 'https://attacker.example/GenOffice-0.2.0.dmg' },
      ])
      actions.onOpenDownload()
      expect(openExternal).toHaveBeenCalledWith(
        'https://cdn.example.com/mac/GenOffice-0.2.0-arm64.dmg',
      )
    } finally {
      restoreArch()
    }
  })

  it('rejects a non-HTTPS baked feed URL', async () => {
    Object.defineProperty(process, 'resourcesPath', { value: '/res', configurable: true })
    readFileSyncMock.mockReturnValue('url: http://cdn.example.com/mac\n')
    const actions = await failTwiceIntoManual(macFiles)
    actions.onOpenDownload()
    expect(openExternal).toHaveBeenCalledWith(
      'https://github.com/genspark-ai/genoffice/releases/latest',
    )
  })

  it('falls back to the generic download page when the feed base cannot be read', async () => {
    // readFileSyncMock throws by default (no app-update.yml)
    const actions = await failTwiceIntoManual([
      { url: 'https://attacker.example/GenOffice-0.2.0-arm64.dmg' },
    ])
    actions.onOpenDownload()
    expect(openExternal).toHaveBeenCalledWith(
      'https://github.com/genspark-ai/genoffice/releases/latest',
    )
  })
})

describe('initAutoUpdater (fake update preview)', () => {
  it('runs a simulated download to completion in unpacked runs', async () => {
    appState.isPackaged = false
    process.env.GENOFFICE_FAKE_UPDATE = '9.9.9'
    const { initAutoUpdater } = await loadUpdater()
    initAutoUpdater(() => null)

    // the fake flow never touches the real updater
    expect(updaterState.listeners.size).toBe(0)

    vi.advanceTimersByTime(1500)
    expect(showUpdateWindow).toHaveBeenCalledTimes(1)
    expect(lastShownState().version).toBe('9.9.9')

    lastShownActions().onDownload()
    expect(pushUpdateState).toHaveBeenCalledWith({ phase: 'downloading', percent: 0 })
    vi.advanceTimersByTime(3000)
    expect(pushUpdateState).toHaveBeenCalledWith({ phase: 'downloaded', percent: 100 })
    expect(checkForUpdates).not.toHaveBeenCalled()
    expect(downloadUpdate).not.toHaveBeenCalled()
  })

  it('closes the window on later and install without touching electron-updater', async () => {
    appState.isPackaged = false
    process.env.GENOFFICE_FAKE_UPDATE = '9.9.9'
    const { initAutoUpdater } = await loadUpdater()
    initAutoUpdater(() => null)
    vi.advanceTimersByTime(1500)

    lastShownActions().onLater()
    expect(closeUpdateWindow).toHaveBeenCalledTimes(1)
    lastShownActions().onInstall()
    expect(closeUpdateWindow).toHaveBeenCalledTimes(2)
    expect(quitAndInstall).not.toHaveBeenCalled()
  })
})
