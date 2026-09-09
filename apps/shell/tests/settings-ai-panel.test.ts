/**
 * @vitest-environment jsdom
 */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import type { Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { AiPanelPrefs } from '@genoffice/ui'
import type { HomeApi } from '../src/shared/home-api'
import { LocaleProvider } from '../src/renderer/src/locale'
import { SettingsModal } from '../src/renderer/src/SettingsModal'

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

async function click(button: HTMLButtonElement): Promise<void> {
  await act(async () => {
    button.click()
    await Promise.resolve()
  })
}

describe('Settings AI panel preferences', () => {
  it('shows the saved spellcheck state and persists the toggle as a patch', async () => {
    let saved: AiPanelPrefs = { fontSize: 'large', customFontSize: 14, spellcheck: false }
    const setAiPanelPrefs = vi.fn(async (patch: Partial<AiPanelPrefs>) => {
      saved = { ...saved, ...patch }
      return saved
    })
    window.aiOffice = {
      getTheme: async () => 'system',
      getDefaultSaveDir: async () => '',
      getAnalyticsEnabled: async () => true,
      setAnalyticsEnabled: async () => true,
      getAiPanelPrefs: async () => saved,
      setAiPanelPrefs,
      getUpdateChannel: async () => 'stable',
      getAppVersion: async () => '1.0.0',
      githubStars: async () => null,
    } as unknown as HomeApi

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
    await click(general!)

    const spellcheck = host.querySelector<HTMLButtonElement>(
      '.set-switch[aria-label="Spell check in AI chat"]',
    )
    expect(spellcheck?.getAttribute('aria-checked')).toBe('false')
    expect(host.textContent).toContain('Large')

    await click(spellcheck!)
    expect(setAiPanelPrefs).toHaveBeenLastCalledWith({ spellcheck: true })
    expect(spellcheck?.getAttribute('aria-checked')).toBe('true')
    expect(host.querySelector('.set-num-input')).toBeNull()
  })

  it('shows a px input for the custom size and persists in-range values as typed', async () => {
    let saved: AiPanelPrefs = { fontSize: 'custom', customFontSize: 20, spellcheck: true }
    const setAiPanelPrefs = vi.fn(async (patch: Partial<AiPanelPrefs>) => {
      saved = { ...saved, ...patch }
      return saved
    })
    window.aiOffice = {
      getTheme: async () => 'system',
      getDefaultSaveDir: async () => '',
      getAnalyticsEnabled: async () => true,
      setAnalyticsEnabled: async () => true,
      getAiPanelPrefs: async () => saved,
      setAiPanelPrefs,
      getUpdateChannel: async () => 'stable',
      getAppVersion: async () => '1.0.0',
      githubStars: async () => null,
    } as unknown as HomeApi

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
    await click(general!)

    const input = host.querySelector<HTMLInputElement>('.set-num-input')
    expect(input?.value).toBe('20')
    expect(host.textContent).toContain('Custom')

    const setValue = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!
    const type = async (text: string) => {
      await act(async () => {
        setValue.call(input, text)
        input!.dispatchEvent(new Event('input', { bubbles: true }))
        await Promise.resolve()
      })
    }
    await act(async () => {
      input!.dispatchEvent(new FocusEvent('focusin', { bubbles: true }))
    })
    await type('1')
    expect(setAiPanelPrefs).not.toHaveBeenCalled()
    await type('18')
    expect(setAiPanelPrefs).toHaveBeenLastCalledWith({ customFontSize: 18 })
    await type('99')
    expect(setAiPanelPrefs).toHaveBeenCalledTimes(1)
    await act(async () => {
      input!.dispatchEvent(new FocusEvent('focusout', { bubbles: true }))
      await Promise.resolve()
    })
    expect(setAiPanelPrefs).toHaveBeenLastCalledWith({ customFontSize: 32 })
    expect(input?.value).toBe('32')
  })
})
