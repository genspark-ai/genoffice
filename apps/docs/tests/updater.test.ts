import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

/**
 * Main-process auto-updater flow (src/main/updater.ts): silent background
 * download over the generic CDN feed, one native dialog once the update is
 * downloaded, install-on-restart or defer until quit.
 */

const appState = { isPackaged: true }
const showMessageBox = vi.fn()

vi.mock('electron', () => ({
  app: {
    get isPackaged() {
      return appState.isPackaged
    },
    getVersion: () => '0.1.0',
  },
  dialog: {
    showMessageBox: (...args: unknown[]) => showMessageBox(...args),
  },
}))

type Listener = (...args: unknown[]) => unknown

const updaterState = {
  listeners: new Map<string, Listener>(),
  autoDownload: true,
  autoInstallOnAppQuit: false,
  disableDifferentialDownload: false,
}
const checkForUpdates = vi.fn(() => Promise.resolve(null))
const quitAndInstall = vi.fn()

vi.mock('electron-updater', () => ({
  autoUpdater: {
    on: (event: string, listener: Listener) => {
      updaterState.listeners.set(event, listener)
    },
    checkForUpdates: () => checkForUpdates(),
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
  },
}))

const FIRST_CHECK_DELAY_MS = 15_000

let platformSpy: { restore: () => void } | null = null

function setPlatform(platform: string): void {
  const original = Object.getOwnPropertyDescriptor(process, 'platform')!
  Object.defineProperty(process, 'platform', { value: platform })
  platformSpy = {
    restore: () => Object.defineProperty(process, 'platform', original),
  }
}

async function loadUpdater() {
  const mod = await import('../src/main/updater')
  return mod
}

async function flushAsync(): Promise<void> {
  // drain the dialog promise chain queued by the update-downloaded listener
  for (let i = 0; i < 10; i++) await Promise.resolve()
}

beforeEach(() => {
  vi.resetModules()
  vi.useFakeTimers()
  appState.isPackaged = true
  updaterState.listeners.clear()
  updaterState.autoDownload = true
  updaterState.autoInstallOnAppQuit = false
  updaterState.disableDifferentialDownload = false
  showMessageBox.mockReset()
  checkForUpdates.mockClear()
  quitAndInstall.mockClear()
  setPlatform('darwin')
})

afterEach(() => {
  vi.useRealTimers()
  platformSpy?.restore()
  platformSpy = null
})

describe('initDocsAutoUpdater', () => {
  it('does nothing in unpacked (dev) runs', async () => {
    appState.isPackaged = false
    const { initDocsAutoUpdater } = await loadUpdater()
    initDocsAutoUpdater(() => null)
    vi.advanceTimersByTime(FIRST_CHECK_DELAY_MS)
    expect(updaterState.listeners.size).toBe(0)
    expect(checkForUpdates).not.toHaveBeenCalled()
  })

  it('does nothing on unsupported platforms', async () => {
    platformSpy?.restore()
    setPlatform('linux')
    const { initDocsAutoUpdater } = await loadUpdater()
    initDocsAutoUpdater(() => null)
    vi.advanceTimersByTime(FIRST_CHECK_DELAY_MS)
    expect(updaterState.listeners.size).toBe(0)
    expect(checkForUpdates).not.toHaveBeenCalled()
  })

  it('configures silent full-package download and checks after the initial delay', async () => {
    const { initDocsAutoUpdater } = await loadUpdater()
    initDocsAutoUpdater(() => null)
    expect(updaterState.autoDownload).toBe(true)
    expect(updaterState.autoInstallOnAppQuit).toBe(true)
    expect(updaterState.disableDifferentialDownload).toBe(true)
    expect(checkForUpdates).not.toHaveBeenCalled()
    vi.advanceTimersByTime(FIRST_CHECK_DELAY_MS)
    expect(checkForUpdates).toHaveBeenCalledTimes(1)
  })

  it('installs immediately when the user confirms the downloaded update', async () => {
    showMessageBox.mockResolvedValue({ response: 0 })
    const { initDocsAutoUpdater } = await loadUpdater()
    initDocsAutoUpdater(() => null)
    updaterState.listeners.get('update-downloaded')!({ version: '0.2.0' })
    await flushAsync()
    expect(showMessageBox).toHaveBeenCalledTimes(1)
    expect(quitAndInstall).toHaveBeenCalledWith(true, true)
  })

  it('does not re-prompt for a version the user dismissed', async () => {
    showMessageBox.mockResolvedValue({ response: 1 })
    const { initDocsAutoUpdater } = await loadUpdater()
    initDocsAutoUpdater(() => null)
    const downloaded = updaterState.listeners.get('update-downloaded')!
    downloaded({ version: '0.2.0' })
    await flushAsync()
    expect(showMessageBox).toHaveBeenCalledTimes(1)
    expect(quitAndInstall).not.toHaveBeenCalled()

    downloaded({ version: '0.2.0' })
    await flushAsync()
    expect(showMessageBox).toHaveBeenCalledTimes(1)

    // a newer version prompts again
    downloaded({ version: '0.3.0' })
    await flushAsync()
    expect(showMessageBox).toHaveBeenCalledTimes(2)
  })

  it('is idempotent across repeated init calls', async () => {
    const { initDocsAutoUpdater } = await loadUpdater()
    initDocsAutoUpdater(() => null)
    initDocsAutoUpdater(() => null)
    vi.advanceTimersByTime(FIRST_CHECK_DELAY_MS)
    expect(checkForUpdates).toHaveBeenCalledTimes(1)
  })
})
