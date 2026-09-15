// @vitest-environment jsdom
/**
 * Tests for the React UI components (PiDialogHost, NotificationToaster, PiStatusBar).
 *
 * Strategy: mount the components against a real ReactUIAdapter inside a
 * PiSessionProvider, drive the adapter from "below" (push dialogs / notify),
 * and assert against the rendered DOM. No external testing-library — we use
 * react-dom/client + container.querySelector + vi.waitFor.
 *
 * jsdom is required: see vitest.config.ts environmentMatchGlobs.
 */

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { createRoot, type Root } from 'react-dom/client'
import { act } from 'react'
import {
  PiDialogHost,
  NotificationToaster,
  PiStatusBar,
} from '../src/components'
import { OfficeSessionContext } from '../src/provider'
import type { ReactUIAdapter } from '../src/ui-adapter'

// ---------------------------------------------------------------------------
// Lightweight helpers
// ---------------------------------------------------------------------------

/** Mount a tree inside a fresh container. Returns the root and container. */
function mount(children: React.ReactNode): { root: Root; container: HTMLDivElement } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => {
    root.render(children)
  })
  return { root, container }
}

/** Unmount and clean up. */
function unmount(root: Root, container: HTMLDivElement) {
  act(() => {
    root.unmount()
  })
  container.remove()
}

/** Find the first element with this data-testid inside the container. */
function byTestId(container: HTMLElement, id: string): HTMLElement | null {
  return container.querySelector(`[data-testid="${id}"]`)
}

/** Find all elements with this data-testid. */
function allByTestId(container: HTMLElement, id: string): HTMLElement[] {
  return Array.from(container.querySelectorAll(`[data-testid="${id}"]`))
}

/**
 * Build a tree that wires a hand-built ReactUIAdapter into the provider
 * context, without spinning up a real pi session. Uses the bare
 * OfficeSessionContext (exported by the provider module for exactly this
 * purpose) instead of the heavy PiSessionProvider.
 */
function ProviderWithAdapter(props: {
  adapter: ReactUIAdapter
  children: React.ReactNode
}): React.ReactNode {
  // Fake session — we never touch session.* in component tests, only uiAdapter.
  const fakeSession = {
    session: {} as never,
    uiAdapter: props.adapter,
    reloadResources: async () => ({ skills: 0, extensions: 0 }),
    dispose: () => {},
  }
  return (
    <OfficeSessionContext.Provider value={fakeSession}>
      {props.children}
    </OfficeSessionContext.Provider>
  )
}

// ---------------------------------------------------------------------------
// PiDialogHost
// ---------------------------------------------------------------------------

describe('PiDialogHost', () => {
  let root: Root | null = null
  let container: HTMLDivElement | null = null

  afterEach(() => {
    if (root && container) unmount(root, container)
    root = null
    container = null
    document.body.innerHTML = ''
  })

  it('renders nothing when no dialogs are pending', async () => {
    const { ReactUIAdapter } = await import('../src/ui-adapter')
    const adapter = new ReactUIAdapter()
    const m = mount(
      <ProviderWithAdapter adapter={adapter}>
        <PiDialogHost />
      </ProviderWithAdapter>
    )
    root = m.root
    container = m.container
    expect(byTestId(m.container, 'pi-dialog-backdrop')).toBeNull()
  })

  it('shows a confirm dialog and resolves with the chosen value', async () => {
    const { ReactUIAdapter } = await import('../src/ui-adapter')
    const adapter = new ReactUIAdapter()
    const m = mount(
      <ProviderWithAdapter adapter={adapter}>
        <PiDialogHost />
      </ProviderWithAdapter>
    )
    root = m.root
    container = m.container

    // Drive the adapter: push a confirm dialog
    let resolved: boolean | undefined = undefined
    await act(async () => {
      const p = adapter.confirm('Save changes?', 'You have unsaved work.')
      // Subscribe to resolution
      p.then((v) => { resolved = v })
    })
    // After act flush, the dialog should appear
    const backdrop = byTestId(m.container, 'pi-dialog-backdrop')
    expect(backdrop).not.toBeNull()
    expect(backdrop?.querySelector('.pi-dialog-title')?.textContent).toBe('Save changes?')

    // Click OK
    const okBtn = byTestId(m.container, 'pi-dialog-ok')
    expect(okBtn).not.toBeNull()
    await act(async () => {
      okBtn!.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })
    // Wait for the promise to resolve
    await act(async () => {
      await new Promise((r) => setTimeout(r, 0))
    })
    expect(resolved).toBe(true)
    expect(adapter.dialogs.length).toBe(0)
    expect(byTestId(m.container, 'pi-dialog-backdrop')).toBeNull()
  })

  it('shows an input dialog and resolves with the typed value', async () => {
    const { ReactUIAdapter } = await import('../src/ui-adapter')
    const adapter = new ReactUIAdapter()
    const m = mount(
      <ProviderWithAdapter adapter={adapter}>
        <PiDialogHost />
      </ProviderWithAdapter>
    )
    root = m.root
    container = m.container

    let resolved: string | undefined = 'unset'
    await act(async () => {
      adapter.input('Rename', 'new name').then((v) => { resolved = v })
    })
    const input = byTestId(m.container, 'pi-dialog-input') as HTMLInputElement | null
    expect(input).not.toBeNull()
    await act(async () => {
      // Simulate user typing
      const nativeSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value')!.set!
      nativeSetter.call(input, 'Document_v2.docx')
      input!.dispatchEvent(new Event('input', { bubbles: true }))
    })
    const okBtn = byTestId(m.container, 'pi-dialog-ok')
    await act(async () => {
      okBtn!.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })
    await act(async () => {
      await new Promise((r) => setTimeout(r, 0))
    })
    expect(resolved).toBe('Document_v2.docx')
  })

  it('shows a select dialog and resolves with the chosen option', async () => {
    const { ReactUIAdapter } = await import('../src/ui-adapter')
    const adapter = new ReactUIAdapter()
    const m = mount(
      <ProviderWithAdapter adapter={adapter}>
        <PiDialogHost />
      </ProviderWithAdapter>
    )
    root = m.root
    container = m.container

    let resolved: string | undefined = 'unset'
    await act(async () => {
      adapter.select('Pick format', ['docx', 'pdf', 'md']).then((v) => { resolved = v })
    })
    const pdfBtn = byTestId(m.container, 'pi-dialog-option-pdf')
    expect(pdfBtn).not.toBeNull()
    await act(async () => {
      pdfBtn!.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })
    await act(async () => {
      await new Promise((r) => setTimeout(r, 0))
    })
    expect(resolved).toBe('pdf')
  })

  it('cancels a confirm dialog when Cancel is clicked', async () => {
    const { ReactUIAdapter } = await import('../src/ui-adapter')
    const adapter = new ReactUIAdapter()
    const m = mount(
      <ProviderWithAdapter adapter={adapter}>
        <PiDialogHost />
      </ProviderWithAdapter>
    )
    root = m.root
    container = m.container

    let resolved: boolean | undefined = undefined
    await act(async () => {
      adapter.confirm('Delete?', 'Cannot undo').then((v) => { resolved = v })
    })
    const cancelBtn = byTestId(m.container, 'pi-dialog-cancel')
    await act(async () => {
      cancelBtn!.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })
    await act(async () => {
      await new Promise((r) => setTimeout(r, 0))
    })
    expect(resolved).toBe(false)
  })

  it('removes the dialog automatically after timeout (no countdown UI by default)', async () => {
    const { ReactUIAdapter } = await import('../src/ui-adapter')
    const adapter = new ReactUIAdapter()
    const m = mount(
      <ProviderWithAdapter adapter={adapter}>
        <PiDialogHost showCountdown={false} />
      </ProviderWithAdapter>
    )
    root = m.root
    container = m.container

    let resolved: boolean | undefined = undefined
    await act(async () => {
      adapter.confirm('Slow?', 'Patient', { timeout: 30 }).then((v) => { resolved = v })
    })
    expect(byTestId(m.container, 'pi-dialog-backdrop')).not.toBeNull()
    // Wait past the timeout
    await new Promise((r) => setTimeout(r, 100))
    expect(resolved).toBe(false)
    expect(byTestId(m.container, 'pi-dialog-backdrop')).toBeNull()
  })
})

// ---------------------------------------------------------------------------
// NotificationToaster
// ---------------------------------------------------------------------------

describe('NotificationToaster', () => {
  let root: Root | null = null
  let container: HTMLDivElement | null = null

  afterEach(() => {
    if (root && container) unmount(root, container)
    root = null
    container = null
    document.body.innerHTML = ''
  })

  it('renders nothing when there are no notifications', async () => {
    const { ReactUIAdapter } = await import('../src/ui-adapter')
    const adapter = new ReactUIAdapter()
    const m = mount(
      <ProviderWithAdapter adapter={adapter}>
        <NotificationToaster />
      </ProviderWithAdapter>
    )
    root = m.root
    container = m.container
    expect(byTestId(m.container, 'pi-toast-container')).toBeNull()
  })

  it('shows notifications by severity and dismisses them on click', async () => {
    const { ReactUIAdapter } = await import('../src/ui-adapter')
    const adapter = new ReactUIAdapter()
    const m = mount(
      <ProviderWithAdapter adapter={adapter}>
        <NotificationToaster />
      </ProviderWithAdapter>
    )
    root = m.root
    container = m.container

    await act(async () => {
      adapter.notify('Saved', 'info')
      adapter.notify('Connection slow', 'warning')
      adapter.notify('Save failed', 'error')
    })

    expect(allByTestId(m.container, 'pi-toast-info').length).toBe(1)
    expect(allByTestId(m.container, 'pi-toast-warning').length).toBe(1)
    expect(allByTestId(m.container, 'pi-toast-error').length).toBe(1)
    expect(adapter.notifications.length).toBe(3)

    // Click × on the warning toast
    const warning = byTestId(m.container, 'pi-toast-warning')!
    const dismiss = warning.querySelector('button')!
    await act(async () => {
      dismiss.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(adapter.notifications.length).toBe(2)
    expect(allByTestId(m.container, 'pi-toast-warning').length).toBe(0)
    expect(allByTestId(m.container, 'pi-toast-info').length).toBe(1)
    expect(allByTestId(m.container, 'pi-toast-error').length).toBe(1)
  })
})

// ---------------------------------------------------------------------------
// PiStatusBar
// ---------------------------------------------------------------------------

describe('PiStatusBar', () => {
  let root: Root | null = null
  let container: HTMLDivElement | null = null

  afterEach(() => {
    if (root && container) unmount(root, container)
    root = null
    container = null
    document.body.innerHTML = ''
  })

  it('renders status entries that have text and hides cleared ones', async () => {
    const { ReactUIAdapter } = await import('../src/ui-adapter')
    const adapter = new ReactUIAdapter()
    const m = mount(
      <ProviderWithAdapter adapter={adapter}>
        <PiStatusBar />
      </ProviderWithAdapter>
    )
    root = m.root
    container = m.container

    await act(async () => {
      adapter.setStatus('model', 'claude-sonnet-4-5')
      adapter.setStatus('tokens', '12.3k')
    })

    expect(byTestId(m.container, 'pi-status-model')?.textContent).toBe('claude-sonnet-4-5')
    expect(byTestId(m.container, 'pi-status-tokens')?.textContent).toBe('12.3k')

    // Clear tokens
    await act(async () => {
      adapter.setStatus('tokens', undefined)
    })
    expect(byTestId(m.container, 'pi-status-tokens')).toBeNull()
    // model still there
    expect(byTestId(m.container, 'pi-status-model')).not.toBeNull()
  })
})

// ---------------------------------------------------------------------------
// Cleanup sanity check (no leaked listeners across mounts)
// ---------------------------------------------------------------------------

describe('listener cleanup', () => {
  it('removes all adapter listeners when the provider unmounts', async () => {
    const { ReactUIAdapter } = await import('../src/ui-adapter')
    const adapter = new ReactUIAdapter()
    // Inspect internal listener set via private field (TS escape hatch)
    expect((adapter as unknown as { dialogListeners: Set<unknown> }).dialogListeners.size).toBe(0)

    const m = mount(
      <ProviderWithAdapter adapter={adapter}>
        <PiDialogHost />
        <NotificationToaster />
        <PiStatusBar />
      </ProviderWithAdapter>
    )

    // After mount, there should be a listener from each subscription
    expect((adapter as unknown as { dialogListeners: Set<unknown> }).dialogListeners.size).toBe(1)
    expect((adapter as unknown as { notificationListeners: Set<unknown> }).notificationListeners.size).toBe(1)
    expect((adapter as unknown as { statusListeners: Set<unknown> }).statusListeners.size).toBe(1)

    unmount(m.root, m.container)

    // After unmount, all listeners should be cleaned up
    expect((adapter as unknown as { dialogListeners: Set<unknown> }).dialogListeners.size).toBe(0)
    expect((adapter as unknown as { notificationListeners: Set<unknown> }).notificationListeners.size).toBe(0)
    expect((adapter as unknown as { statusListeners: Set<unknown> }).statusListeners.size).toBe(0)
  })
})
