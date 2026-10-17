import { act } from 'react'
import type { ReactElement } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { LocaleProvider } from '../src/renderer/i18n/locale'
import { FakeMailApi } from '../src/renderer/devFakes'
import { FakeOutboxSyncApi, resolveOutboxSyncApi } from '../src/renderer/components/OutboxApi'
import type { FakeOutboxSeed } from '../src/renderer/components/OutboxApi'
import { OutboxPanel } from '../src/renderer/components/OutboxPanel'
import { OutboxBanner } from '../src/renderer/components/OutboxBanner'
import { OfflineIndicator } from '../src/renderer/components/OfflineIndicator'
import { AccountSyncStatus } from '../src/renderer/components/AccountSyncStatus'

/**
 * Lane O component tests over the outbox/offline fakes: the five spec
 * states, retry/cancel affordances, the offline indicator, the extended
 * P3 banner (same .mail-notice root, never a second banner), per-account
 * sync status, and per-driver send caps. Example addresses only.
 */

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

let containers: HTMLElement[] = []

function render(node: ReactElement): HTMLElement {
  const container = document.createElement('div')
  document.body.appendChild(container)
  containers.push(container)
  const root = createRoot(container)
  act(() => {
    root.render(<LocaleProvider initial="en">{node}</LocaleProvider>)
  })
  return container
}

async function settle(rounds = 10): Promise<void> {
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

const ALL_STATES: FakeOutboxSeed[] = [
  { id: 'o1', subject: 'Queued note', to: ['a@example.com'], state: 'queued' },
  { id: 'o2', subject: 'Sending note', to: ['b@example.com'], state: 'sending' },
  {
    id: 'o3',
    subject: 'Failed note',
    to: ['c@example.com'],
    state: 'failed',
    lastError: 'smtp: connection refused',
  },
  { id: 'o4', subject: 'Sent note', to: ['d@example.com'], state: 'sent' },
  { id: 'o5', subject: 'Cancelled note', to: ['e@example.com'], state: 'cancelled' },
]

beforeEach(() => {
  vi.restoreAllMocks()
  // LocaleProvider listens for language changes on the global bridge;
  // the P3 fake serves that while components take their own api prop.
  ;(window as unknown as { mailApi: FakeMailApi }).mailApi = new FakeMailApi()
})

afterEach(() => {
  for (const container of containers) container.remove()
  containers = []
})

describe('OutboxPanel', () => {
  it('renders all five spec states with translated badges', async () => {
    const api = new FakeOutboxSyncApi('demo', ALL_STATES)
    const container = render(<OutboxPanel api={api} accountId="demo" />)
    await settle()
    for (const badge of ['Queued', 'Sending', 'Failed', 'Sent', 'Cancelled']) {
      expect(container.textContent).toContain(badge)
    }
    expect(container.textContent).toContain('Queued note')
  })

  it('shows the empty state when the outbox is clear', async () => {
    const api = new FakeOutboxSyncApi('demo', [])
    const container = render(<OutboxPanel api={api} accountId="demo" />)
    await settle()
    expect(container.textContent).toContain('Outbox is empty.')
  })

  it('retries a failed entry back to queued', async () => {
    const api = new FakeOutboxSyncApi('demo', [ALL_STATES[2]])
    const container = render(<OutboxPanel api={api} accountId="demo" />)
    await settle()
    expect(container.textContent).toContain('smtp: connection refused')
    const retry = buttonByText(container, 'Retry')
    expect(retry).toBeDefined()
    await act(async () => {
      retry?.click()
    })
    await settle()
    const records = await api.outboxList('demo')
    expect(records[0].state).toBe('queued')
    expect(container.textContent).toContain('Queued')
  })

  it('cancels a queued entry', async () => {
    const api = new FakeOutboxSyncApi('demo', [ALL_STATES[0]])
    const container = render(<OutboxPanel api={api} accountId="demo" />)
    await settle()
    const cancel = buttonByText(container, 'Cancel')
    expect(cancel).toBeDefined()
    await act(async () => {
      cancel?.click()
    })
    await settle()
    const records = await api.outboxList('demo')
    expect(records[0].state).toBe('cancelled')
    expect(container.textContent).toContain('Cancelled')
  })

  it('offers no retry/cancel on a sending entry', async () => {
    const api = new FakeOutboxSyncApi('demo', [ALL_STATES[1]])
    const container = render(<OutboxPanel api={api} accountId="demo" />)
    await settle()
    expect(buttonByText(container, 'Retry')).toBeUndefined()
    expect(buttonByText(container, 'Cancel')).toBeUndefined()
  })

  it('sync-now invokes the bridge and tracks syncing events', async () => {
    const api = new FakeOutboxSyncApi('demo', [])
    const container = render(<OutboxPanel api={api} accountId="demo" />)
    await settle()
    const sync = buttonByText(container, 'Sync now')
    expect(sync).toBeDefined()
    await act(async () => {
      sync?.click()
    })
    await settle()
    expect(api.syncNowCalls).toEqual([{ accountId: 'demo', folderId: undefined }])
    await act(async () => {
      api.emitSync({ kind: 'sync-state', accountId: 'demo', folderId: 'inbox', state: 'syncing' })
    })
    await settle()
    const busy = buttonByText(container, 'Syncing…')
    expect(busy?.disabled).toBe(true)
    await act(async () => {
      api.emitSync({ kind: 'sync-state', accountId: 'demo', folderId: 'inbox', state: 'idle' })
    })
    await settle()
    expect(buttonByText(container, 'Sync now')).toBeDefined()
  })
})

describe('OfflineIndicator', () => {
  it('replays online on subscribe and flips with status events', async () => {
    const api = new FakeOutboxSyncApi('demo', [])
    const container = render(<OfflineIndicator api={api} accountId="demo" />)
    await settle()
    expect(container.textContent).toContain('Online')
    await act(async () => {
      api.setOnline(false)
    })
    await settle()
    expect(container.textContent).toContain('Offline')
    await act(async () => {
      api.setOnline(true)
    })
    await settle()
    expect(container.textContent).toContain('Online')
  })

  it('ignores other accounts when scoped', async () => {
    const api = new FakeOutboxSyncApi('demo', [])
    const container = render(<OfflineIndicator api={api} accountId="other" />)
    await settle()
    // Replay carries demo, so the other-scoped chip stays hidden.
    expect(container.textContent).toBe('')
  })
})

describe('OutboxBanner', () => {
  it('extends the P3 banner root and opens the panel', async () => {
    const api = new FakeOutboxSyncApi('demo', [ALL_STATES[0], ALL_STATES[2]])
    const onOpenPanel = vi.fn()
    const container = render(
      <OutboxBanner
        api={api}
        accountId="demo"
        notice="Queued in the outbox."
        onOpenPanel={onOpenPanel}
      />,
    )
    await settle()
    const banner = container.querySelector('.mail-notice')
    expect(banner).not.toBeNull()
    // One banner only: no second notice root is forked alongside it.
    expect(container.querySelectorAll('.mail-notice')).toHaveLength(1)
    expect(container.textContent).toContain('Queued in the outbox.')
    const open = buttonByText(container, 'View outbox (2)')
    expect(open).toBeDefined()
    act(() => {
      open?.click()
    })
    expect(onOpenPanel).toHaveBeenCalledTimes(1)
  })

  it('renders nothing when clean, online, and notice-free', async () => {
    const api = new FakeOutboxSyncApi('demo', [])
    const container = render(
      <OutboxBanner api={api} accountId="demo" notice={null} onOpenPanel={() => undefined} />,
    )
    await settle()
    expect(container.querySelector('.mail-notice')).toBeNull()
  })

  it('surfaces the offline chip inside the same banner', async () => {
    const api = new FakeOutboxSyncApi('demo', [])
    const container = render(
      <OutboxBanner api={api} accountId="demo" notice={null} onOpenPanel={() => undefined} />,
    )
    await settle()
    await act(async () => {
      api.setOnline(false)
    })
    await settle()
    expect(container.querySelector('.mail-notice')).not.toBeNull()
    expect(container.textContent).toContain('Offline')
  })
})

describe('AccountSyncStatus', () => {
  it('shows per-driver send caps', async () => {
    const api = new FakeOutboxSyncApi('demo', [])
    const container = render(
      <AccountSyncStatus api={api} accountId="demo" label="Demo" driver="imap" />,
    )
    await settle()
    expect(container.textContent).toContain('Up to date')
    expect(container.textContent).toContain('IMAP: 50 MB total, 25 MB per file, 100 recipients')
  })

  it('shows gmail and graph caps', async () => {
    const gmail = render(
      <AccountSyncStatus
        api={new FakeOutboxSyncApi('demo', [])}
        accountId="demo"
        label="Demo"
        driver="gmail"
      />,
    )
    await settle()
    expect(gmail.textContent).toContain('Gmail: 35 MB total, 35 MB per file, 500 recipients')
    const graph = render(
      <AccountSyncStatus
        api={new FakeOutboxSyncApi('demo', [])}
        accountId="demo"
        label="Demo"
        driver="graph"
      />,
    )
    await settle()
    expect(graph.textContent).toContain(
      'Microsoft Graph: 4 MB total, 3 MB per file, 500 recipients',
    )
  })

  it('reflects resync and error events for its account only', async () => {
    const api = new FakeOutboxSyncApi('demo', [])
    const container = render(<AccountSyncStatus api={api} accountId="demo" label="Demo" />)
    await settle()
    await act(async () => {
      api.emitSync({ kind: 'resync', accountId: 'demo', folderId: 'inbox', reason: 'manual' })
    })
    await settle()
    expect(container.textContent).toContain('Mailbox resynced from the server.')
    await act(async () => {
      api.emitSync({
        kind: 'resync',
        accountId: 'someone-else',
        folderId: 'inbox',
        reason: 'manual',
      })
    })
    await settle()
    // Other-account events never overwrite this account's note.
    expect(container.textContent).toContain('Mailbox resynced from the server.')
    await act(async () => {
      api.emitSync({
        kind: 'sync-state',
        accountId: 'demo',
        folderId: 'inbox',
        state: 'error',
      })
    })
    await settle()
    expect(container.textContent).toContain('Sync failed; will retry.')
  })
})

describe('resolveOutboxSyncApi', () => {
  it('returns null when the bridge lacks the P4 methods', () => {
    const previous = (window as unknown as { mailApi?: unknown }).mailApi
    ;(window as unknown as { mailApi?: unknown }).mailApi = {}
    try {
      expect(resolveOutboxSyncApi()).toBeNull()
    } finally {
      ;(window as unknown as { mailApi?: unknown }).mailApi = previous
    }
  })

  it('binds to a full bridge', () => {
    const previous = (window as unknown as { mailApi?: unknown }).mailApi
    const fake = new FakeOutboxSyncApi('demo', [])
    ;(window as unknown as { mailApi?: unknown }).mailApi = fake
    try {
      expect(resolveOutboxSyncApi()).toBe(fake)
    } finally {
      ;(window as unknown as { mailApi?: unknown }).mailApi = previous
    }
  })
})
