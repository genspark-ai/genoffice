/**
 * @vitest-environment jsdom
 */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import type { Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { HomeApi, SystemPresenceSettings } from '../src/shared/home-api'
import { LocaleProvider } from '../src/renderer/src/locale'
import { SettingsModal } from '../src/renderer/src/SettingsModal'

/**
 * The settings center's system-presence rows: the login-item and hotkey
 * switches and the accelerator input talk to the home IPC surface
 * (getSystemPresence / setLoginItemEnabled / setGlobalHotkey) and re-render
 * from the returned snapshot.
 */

const actEnvironment = globalThis as typeof globalThis & {
  IS_REACT_ACT_ENVIRONMENT?: boolean
}
actEnvironment.IS_REACT_ACT_ENVIRONMENT = true

let host: HTMLDivElement
let root: Root

beforeEach(() => {
  host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
})

afterEach(() => {
  act(() => root.unmount())
  host.remove()
})

const PRESENCE: SystemPresenceSettings = {
  loginItem: false,
  hotkey: { enabled: true, accelerator: 'Alt+Space' },
  hotkeyActive: true,
  loginItemSupported: true,
}

function stubApi(overrides: Partial<HomeApi> = {}): HomeApi {
  return {
    getTheme: async () => 'system',
    getDefaultSaveDir: async () => '',
    getAnalyticsEnabled: async () => true,
    setAnalyticsEnabled: async () => true,
    getUpdateChannel: async () => 'stable',
    getAppVersion: async () => '1.0.0',
    githubStars: async () => null,
    getSystemPresence: async () => PRESENCE,
    ...overrides,
  } as unknown as HomeApi
}

async function openGeneral(): Promise<void> {
  await act(async () => {
    root.render(
      createElement(
        LocaleProvider,
        { initial: 'en' },
        createElement(SettingsModal, {
          status: null,
          loggingOut: false,
          loginWaiting: false,
          loginUrl: null,
          urlCopied: false,
          onOpenLoginUrl: vi.fn(),
          onCopyLoginUrl: vi.fn(),
          onClose: vi.fn(),
          onLogin: vi.fn(),
          onLogout: vi.fn(),
        }),
      ),
    )
    await Promise.resolve()
  })
  const general = Array.from(host.querySelectorAll<HTMLButtonElement>('.set-nav-item')).find(
    (button) => button.textContent?.includes('General'),
  )
  expect(general).toBeDefined()
  await act(async () => {
    general!.click()
    await Promise.resolve()
  })
}

async function click(element: HTMLElement): Promise<void> {
  await act(async () => {
    element.click()
    await Promise.resolve()
  })
}

describe('Settings system presence', () => {
  it('renders both switches from getSystemPresence and toggles the login item', async () => {
    const setLoginItem = vi
      .fn<(enabled: boolean) => Promise<SystemPresenceSettings>>()
      .mockResolvedValue({ ...PRESENCE, loginItem: true })
    window.aiOffice = stubApi({ setLoginItemEnabled: setLoginItem })
    await openGeneral()

    const login = host.querySelector<HTMLButtonElement>('.set-switch[aria-label="Launch at login"]')
    const hotkey = host.querySelector<HTMLButtonElement>(
      '.set-switch[aria-label="Global shortcut"]',
    )
    expect(login?.getAttribute('aria-checked')).toBe('false')
    expect(hotkey?.getAttribute('aria-checked')).toBe('true')

    await click(login!)
    expect(setLoginItem).toHaveBeenCalledWith(true)
    // the row re-renders from the returned snapshot
    expect(login?.getAttribute('aria-checked')).toBe('true')
  })

  it('commits a new accelerator on blur and reverts when the main process refuses', async () => {
    const setHotkey = vi
      .fn<(patch: { enabled?: boolean; accelerator?: string }) => Promise<SystemPresenceSettings>>()
      .mockResolvedValueOnce({
        ...PRESENCE,
        hotkey: { enabled: true, accelerator: 'Ctrl+Shift+K' },
      })
      .mockRejectedValueOnce(new Error('invalid'))
    window.aiOffice = stubApi({ setGlobalHotkey: setHotkey })
    await openGeneral()

    const input = host.querySelector<HTMLInputElement>('#set-hotkey-accelerator')
    expect(input?.value).toBe('Alt+Space')

    // a valid edit persists and the input shows the returned value
    await act(async () => {
      input!.focus()
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!
      setter.call(input!, 'Ctrl+Shift+K')
      input!.dispatchEvent(new Event('input', { bubbles: true }))
      await Promise.resolve()
    })
    await act(async () => {
      input!.blur()
      await Promise.resolve()
    })
    expect(setHotkey).toHaveBeenCalledWith({ accelerator: 'Ctrl+Shift+K' })
    expect(input?.value).toBe('Ctrl+Shift+K')

    // a refused edit reverts to the persisted accelerator
    await act(async () => {
      input!.focus()
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!
      setter.call(input!, 'junk')
      input!.dispatchEvent(new Event('input', { bubbles: true }))
      await Promise.resolve()
    })
    await act(async () => {
      input!.blur()
      await Promise.resolve()
    })
    expect(setHotkey).toHaveBeenCalledWith({ accelerator: 'junk' })
    expect(input?.value).toBe('Ctrl+Shift+K')
  })

  it('disables the login-item switch on platforms that cannot drive it', async () => {
    window.aiOffice = stubApi({
      getSystemPresence: async () => ({ ...PRESENCE, loginItemSupported: false }),
    })
    await openGeneral()
    const login = host.querySelector<HTMLButtonElement>('.set-switch[aria-label="Launch at login"]')
    expect(login?.disabled).toBe(true)
  })

  it('hides the accelerator input while the hotkey is off', async () => {
    window.aiOffice = stubApi({
      getSystemPresence: async () => ({
        ...PRESENCE,
        hotkey: { enabled: false, accelerator: 'Alt+Space' },
      }),
    })
    await openGeneral()
    expect(host.querySelector('#set-hotkey-accelerator')).toBeNull()
  })
})
