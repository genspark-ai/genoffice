// @vitest-environment jsdom
import { afterEach, beforeAll, describe, expect, it } from 'vitest'
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import type { Root } from 'react-dom/client'

/**
 * Renders the real Home screen against the browser mock and drives it with
 * real DOM events: guards against hook-order / runtime errors and covers the
 * list interactions (shortcuts, undo-delete, context menu, sort, search, menus, drop).
 */

let root: Root | null = null
let host: HTMLElement | null = null

const tick = (ms = 0) => act(async () => void (await new Promise((r) => setTimeout(r, ms))))
const q = <T extends Element = HTMLElement>(sel: string) => document.querySelector<T>(sel)
const qa = (sel: string) => Array.from(document.querySelectorAll<HTMLElement>(sel))
const names = () => qa('.recent-row .recent-name').map((el) => el.textContent)
const key = (k: string, init: KeyboardEventInit = {}) =>
  act(async () => {
    window.dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true, ...init }))
  })

async function mount() {
  const { LocaleProvider } = await import('../src/renderer/src/locale')
  const { ToastProvider } = await import('../src/renderer/src/toast')
  const { Home } = await import('../src/renderer/src/Home')
  host = document.createElement('div')
  document.body.appendChild(host)
  root = createRoot(host)
  await act(async () => {
    root!.render(
      createElement(
        LocaleProvider,
        { initial: 'en' },
        createElement(ToastProvider, null, createElement(Home)),
      ),
    )
  })
}

beforeAll(async () => {
  ;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
  class IO {
    observe() {}
    unobserve() {}
    disconnect() {}
  }
  ;(window as unknown as { IntersectionObserver: unknown }).IntersectionObserver = IO
  await import('../src/renderer/src/browser-mock')
})

afterEach(async () => {
  await act(async () => root?.unmount())
  host?.remove()
  document.body.innerHTML = ''
  root = null
})

describe('Home screen (jsdom smoke)', () => {
  it('renders the file rows once the first page has loaded', async () => {
    await mount()
    await tick()
    await tick()
    expect(q('.recent-table.skeleton')).toBeNull()
    expect(qa('.recent-row')).toHaveLength(5)
  })

  it('Ctrl+A selects all, Delete hides them with an Undo toast, Undo restores them', async () => {
    await mount()
    await tick()
    await tick()
    await key('a', { ctrlKey: true })
    expect(q('.selection-bar')).not.toBeNull()
    await key('Delete')
    expect(qa('.recent-row')).toHaveLength(0)
    const undo = q('.toast .toast-action')
    expect(undo?.textContent).toBe('Undo')
    await act(async () => undo!.click())
    await tick()
    expect(qa('.recent-row')).toHaveLength(5)
  })

  it('right-click opens the row menu', async () => {
    await mount()
    await tick()
    await tick()
    const row = qa('.recent-row')[1]!
    await act(async () => {
      row.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true }))
    })
    expect(row.querySelector('[role="menu"], .row-menu')).not.toBeNull()
  })

  it('column header sorts by name, and again to flip', async () => {
    await mount()
    await tick()
    await tick()
    const nameSort = q('.col-name .col-sort')!
    await act(async () => nameSort.click())
    expect(names()[0]).toBe('Design_System_Spec.md')
    await act(async () => q('.col-name .col-sort')!.click())
    expect(names()[0]).toBe('Q4_Security_Audit.pdf')
  })

  it('search filters the list by file name', async () => {
    await mount()
    await tick()
    await tick()
    const input = q<HTMLInputElement>('.file-search input')!
    await act(async () => {
      const set = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!
      set.call(input, 'keynote')
      input.dispatchEvent(new Event('input', { bubbles: true }))
    })
    await tick(300)
    expect(names()).toEqual(['Product_Keynote_v3.pptx'])
  })

  it('New menu lists the file types; account menu opens from the rail', async () => {
    await mount()
    await tick()
    await act(async () => q('.new-btn')!.click())
    expect(qa('.new-menu [role="menuitem"]').length).toBeGreaterThanOrEqual(5)
    await act(async () => q('.rail .account-btn')!.click())
    expect(q('.acct-menu')).not.toBeNull()
  })

  it('dragging files over the window shows the drop overlay', async () => {
    await mount()
    await tick()
    await act(async () => {
      const ev = new Event('dragenter', { bubbles: true, cancelable: true })
      Object.defineProperty(ev, 'dataTransfer', { value: { types: ['Files'] } })
      window.dispatchEvent(ev)
    })
    expect(q('.drop-overlay')).not.toBeNull()
  })

  it('arrow keys move focus between rows', async () => {
    await mount()
    await tick()
    await tick()
    await key('ArrowDown')
    expect(document.activeElement?.classList.contains('recent-item')).toBe(true)
  })
})
