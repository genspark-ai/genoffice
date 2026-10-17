import { act } from 'react'
import type { ReactElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { LocaleProvider } from '../src/renderer/i18n/locale'
import App from '../src/renderer/App'
import { FakeMailApi } from '../src/renderer/devFakes'

/**
 * P4 shell-mount port (UI lane step one; source commit 703c0f16 was lost
 * with its workspace, so this rebuilds the mount from the O-lane
 * components in this stack): OutboxBanner replaces the P3 notice div,
 * OutboxPanel opens as a dialog, the sidebar carries the offline
 * indicator plus per-account sync status, all over seeded fakes.
 */

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

let containers: Array<{ container: HTMLElement; root: Root }> = []

function render(node: ReactElement): HTMLElement {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  containers.push({ container, root })
  act(() => {
    root.render(<LocaleProvider initial="en">{node}</LocaleProvider>)
  })
  return container
}

async function settle(rounds = 20): Promise<void> {
  await act(async () => {
    for (let i = 0; i < rounds; i += 1) {
      await Promise.resolve()
    }
  })
}

function buttonByText(container: HTMLElement, text: string): HTMLButtonElement | undefined {
  return [...container.querySelectorAll('button')].find((button) => button.textContent === text) as
    HTMLButtonElement | undefined
}

function rowBySubject(container: HTMLElement, subject: string): HTMLElement | undefined {
  return [...container.querySelectorAll('.mail-outbox-item')].find((row) =>
    row.textContent?.includes(subject),
  ) as HTMLElement | undefined
}

function mountSeeded(): { api: FakeMailApi; container: HTMLElement } {
  const api = new FakeMailApi()
  api.getAiSettings = (async () => ({})) as FakeMailApi['getAiSettings']
  ;(window as unknown as { mailApi: FakeMailApi }).mailApi = api
  api.seedOutboxDemo('demo')
  const container = render(<App api={api} />)
  return { api, container }
}

beforeEach(() => {
  const api = new FakeMailApi()
  api.getAiSettings = (async () => ({})) as FakeMailApi['getAiSettings']
  ;(window as unknown as { mailApi: FakeMailApi }).mailApi = api
})

afterEach(() => {
  for (const { container, root } of containers) {
    act(() => {
      root.unmount()
    })
    container.remove()
  }
  containers = []
})

describe('mail shell outbox mount', () => {
  it('shows the seeded banner, online chip, and per-account caps', async () => {
    const { container } = mountSeeded()
    await settle()
    expect(buttonByText(container, 'View outbox (3)')).toBeDefined()
    expect(container.querySelector('.mail-offline-chip.is-online')?.textContent).toContain('Online')
    expect(container.querySelector('.mail-sync-caps-detail')?.textContent).toContain('IMAP')
    expect(container.querySelector('.mail-sync-caps-detail')?.textContent).toContain('50 MB')
  })

  it('opens the panel dialog with retry/cancel affordances and closes it', async () => {
    const { container } = mountSeeded()
    await settle()
    act(() => {
      buttonByText(container, 'View outbox (3)')?.click()
    })
    await settle()
    expect(container.querySelector('[role="dialog"]')).not.toBeNull()
    for (const badge of ['Queued', 'Sending', 'Failed']) {
      expect(container.querySelector('[role="dialog"]')?.textContent).toContain(badge)
    }
    // Retry moves the failed row back to queued.
    const failedRow = rowBySubject(container, 'Failed note')
    const retryButton = failedRow ? buttonByText(failedRow, 'Retry') : undefined
    expect(retryButton).toBeDefined()
    act(() => {
      retryButton?.click()
    })
    await settle()
    expect(rowBySubject(container, 'Failed note')?.textContent).toContain('Queued')
    // Cancel moves a queued row to cancelled.
    const queuedRow = rowBySubject(container, 'Queued note')
    const cancelButton = queuedRow ? buttonByText(queuedRow, 'Cancel') : undefined
    expect(cancelButton).toBeDefined()
    act(() => {
      cancelButton?.click()
    })
    await settle()
    expect(rowBySubject(container, 'Queued note')?.textContent).toContain('Cancelled')
    // The dialog closes through the shell close affordance.
    act(() => {
      buttonByText(container, 'Close')?.click()
    })
    await settle()
    expect(container.querySelector('[role="dialog"]')).toBeNull()
  })

  it('flips the offline chip and banner text when connectivity drops', async () => {
    const { api, container } = mountSeeded()
    await settle()
    act(() => {
      api.setOnline(false)
    })
    await settle()
    expect(container.querySelector('.mail-offline-chip.is-offline')).not.toBeNull()
    expect(container.querySelector('.mail-notice')?.textContent).toContain('Offline')
  })

  it('renders no notice surface when empty and online', async () => {
    const api = new FakeMailApi()
    ;(window as unknown as { mailApi: FakeMailApi }).mailApi = api
    const container = render(<App api={api} />)
    await settle()
    expect(container.querySelector('.mail-notice')).toBeNull()
  })
})
