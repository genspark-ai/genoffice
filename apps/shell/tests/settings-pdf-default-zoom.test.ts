/**
 * @vitest-environment jsdom
 */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import type { Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { HomeApi, PdfDefaultZoom } from '../src/shared/home-api'
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

async function openGeneral(): Promise<void> {
  const general = Array.from(host.querySelectorAll<HTMLButtonElement>('.set-nav-item')).find(
    (button) => button.textContent?.includes('General'),
  )
  expect(general, 'General nav item').toBeDefined()
  await click(general!)
}

describe('Settings default PDF zoom (genoffice#2454)', () => {
  it('shows the saved zoom and persists a change', async () => {
    let saved: PdfDefaultZoom = 'width'
    const setPdfDefaultZoom = vi.fn(async (zoom: PdfDefaultZoom) => {
      saved = zoom
      return saved
    })
    window.aiOffice = {
      getTheme: async () => 'system',
      getDefaultSaveDir: async () => '',
      getAnalyticsEnabled: async () => true,
      getPdfDefaultZoom: async () => saved,
      setPdfDefaultZoom,
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
    await openGeneral()

    const zoomDropdown = host.querySelector<HTMLElement>('[aria-label="Default PDF zoom"]')
    expect(zoomDropdown, 'Default PDF zoom dropdown').not.toBeNull()
    expect(zoomDropdown?.textContent).toContain('Fit width')

    await click(zoomDropdown!)
    const fitPage = Array.from(
      document.querySelectorAll<HTMLButtonElement>('[role="option"]'),
    ).find((option) => option.textContent === 'Fit page')
    expect(fitPage, 'Fit page option').toBeDefined()
    await click(fitPage!)

    expect(setPdfDefaultZoom).toHaveBeenLastCalledWith('page')
    expect(saved).toBe('page')
    expect(zoomDropdown?.textContent).toContain('Fit page')
  })

  it('shows last used and can switch to it', async () => {
    let saved: PdfDefaultZoom = 'last'
    const setPdfDefaultZoom = vi.fn(async (zoom: PdfDefaultZoom) => {
      saved = zoom
      return saved
    })
    window.aiOffice = {
      getTheme: async () => 'system',
      getDefaultSaveDir: async () => '',
      getAnalyticsEnabled: async () => true,
      getPdfDefaultZoom: async () => saved,
      setPdfDefaultZoom,
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
    await openGeneral()

    const zoomDropdown = host.querySelector<HTMLElement>('[aria-label="Default PDF zoom"]')
    expect(zoomDropdown?.textContent).toContain('Last used zoom')

    await click(zoomDropdown!)
    const fitWidth = Array.from(
      document.querySelectorAll<HTMLButtonElement>('[role="option"]'),
    ).find((option) => option.textContent === 'Fit width')
    await click(fitWidth!)
    expect(setPdfDefaultZoom).toHaveBeenLastCalledWith('width')
  })
})
