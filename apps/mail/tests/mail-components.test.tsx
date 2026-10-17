import { act } from 'react'
import type { ReactElement } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import App from '../src/renderer/App'
import { LocaleProvider } from '../src/renderer/i18n/locale'
import { FolderTree } from '../src/renderer/folders/FolderTree'
import { MessageList } from '../src/renderer/list/MessageList'
import { MessageReader } from '../src/renderer/reader/MessageReader'
import { FakeMailApi } from '../src/renderer/devFakes'

/**
 * Lane R component tests over IPC fakes: folder/list render, the
 * remote-blocked default in the reader (plus the show-images escape
 * hatch), phishing flags, and an App smoke pass.
 */

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

function clearTestStorage(): void {
  try {
    if (typeof localStorage !== 'undefined') localStorage.clear()
  } catch {
    // jsdom here has no origin storage; the allowlist falls back to memory.
  }
}

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

async function settle(rounds = 8): Promise<void> {
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

beforeEach(() => {
  clearTestStorage()
  ;(window as unknown as { mailApi: FakeMailApi }).mailApi = new FakeMailApi()
})

afterEach(() => {
  for (const container of containers) container.remove()
  containers = []
})

describe('FolderTree', () => {
  it('renders folders with unread badges over fakes', async () => {
    const api = new FakeMailApi()
    const accounts = await api.listAccounts()
    const folders = await api.listFolders(accounts[0].accountId)
    const container = render(
      <FolderTree
        accounts={accounts}
        foldersByAccount={{ [accounts[0].accountId]: folders }}
        selected={{ accountId: accounts[0].accountId, folderId: 'inbox' }}
        onSelect={() => undefined}
        onAddAccount={() => undefined}
        onRemoveAccount={() => undefined}
      />,
    )
    expect(container.textContent).toContain('Inbox')
    expect(container.textContent).toContain('Sent')
    expect(container.textContent).toContain('Demo Account')
    const badge = container.querySelector('.mail-folder-unread')
    expect(badge?.textContent).toBe('3')
    expect(container.querySelector('.mail-folder.is-active')).not.toBeNull()
  })
})

describe('MessageList', () => {
  it('renders rows with a virtualized count footer', async () => {
    const api = new FakeMailApi()
    const page = await api.listMessages('demo', 'inbox', 50)
    const container = render(
      <MessageList
        messages={page.envelopes}
        selectedUid={null}
        loading={false}
        hasMore={false}
        onSelect={() => undefined}
        onLoadMore={() => undefined}
      />,
    )
    expect(container.textContent).toContain('Flagged and unread sample')
    expect(container.textContent).toContain(`${page.envelopes.length} messages`)
    expect(container.querySelector('[role="listbox"]')).not.toBeNull()
    expect(container.querySelectorAll('[role="option"]').length).toBe(page.envelopes.length)
  })

  it('shows the empty state when no messages match', () => {
    const container = render(
      <MessageList
        messages={[]}
        selectedUid={null}
        loading={false}
        hasMore={false}
        onSelect={() => undefined}
        onLoadMore={() => undefined}
      />,
    )
    expect(container.textContent).toContain('No messages here.')
  })
})

describe('MessageReader', () => {
  it('blocks remote images by default and reveals them on Show images', async () => {
    const api = new FakeMailApi()
    const page = await api.listMessages('demo', 'inbox', 50)
    const banner = page.envelopes.find((env) => env.ref.uid === 'm2')
    expect(banner).toBeDefined()
    const container = render(<MessageReader api={api} envelope={banner ?? null} />)
    await settle()
    expect(container.querySelector('.mq-blocked')).not.toBeNull()
    expect(container.querySelector('.mq-body img[src^="http"]')).toBeNull()
    expect(container.textContent).toContain('Remote images are blocked by default.')
    const show = buttonByText(container, 'Show images')
    expect(show).toBeDefined()
    act(() => {
      show?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })
    expect(container.querySelector('.mq-body img[src^="http"]')).not.toBeNull()
    expect(container.querySelector('.mq-blocked')).toBeNull()
  })

  it('renders authored colors without chrome token references', async () => {
    const api = new FakeMailApi()
    const page = await api.listMessages('demo', 'inbox', 50)
    const invoice = page.envelopes.find((env) => env.ref.uid === 'm1')
    const container = render(<MessageReader api={api} envelope={invoice ?? null} />)
    await settle()
    const body = container.querySelector('.mq-body')
    expect(body?.innerHTML).toContain('#107c41')
    expect(body?.innerHTML).not.toContain('var(--')
  })

  it('surfaces Reply-To divergence and lookalike senders as flags', async () => {
    const api = new FakeMailApi()
    const page = await api.listMessages('demo', 'inbox', 50)
    const odd = page.envelopes.find((env) => env.ref.uid === 'm3')
    const spoof = page.envelopes.find((env) => env.ref.uid === 'm4')
    const oddBox = render(
      <MessageReader
        api={api}
        envelope={odd ?? null}
        security={{ replyTo: 'outside@example.org' }}
      />,
    )
    await settle()
    expect(oddBox.textContent).toContain('Reply-To differs from the sender.')
    const spoofBox = render(<MessageReader api={api} envelope={spoof ?? null} />)
    await settle()
    expect(spoofBox.textContent).toContain('This domain looks like a known sender.')
  })

  it('asks for confirmation before opening links from untrusted senders', async () => {
    const api = new FakeMailApi()
    const page = await api.listMessages('demo', 'inbox', 50)
    const invoice = page.envelopes.find((env) => env.ref.uid === 'm1')
    const container = render(<MessageReader api={api} envelope={invoice ?? null} />)
    await settle()
    expect(container.querySelector('[role="alertdialog"]')).toBeNull()
    expect(container.textContent).not.toContain('This link goes outside the allowlist.')
  })
})

describe('App smoke', () => {
  it('composes tree, search, list, reader, and the compose mount over fakes', async () => {
    const container = render(<App api={new FakeMailApi()} />)
    await settle(12)
    expect(container.querySelector('.mail-app')).not.toBeNull()
    expect(container.textContent).toContain('Inbox')
    const searchInput = container.querySelector('#mail-search-input')
    expect(searchInput?.getAttribute('placeholder')).toBe('Search this mailbox')
    expect(container.textContent).toContain('Flagged and unread sample')
    expect(container.querySelector('[data-compose-mount="true"]')).not.toBeNull()
    expect(container.textContent).toContain('Ask about this mail')
  })
})
