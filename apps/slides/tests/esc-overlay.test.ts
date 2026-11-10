import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { useEscOverlay, useEscOverlayOpen } from '../src/renderer/esc-overlay'

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

function Overlay({ open, onEscape }: { open: boolean; onEscape?: () => void }) {
  useEscOverlay(open, onEscape)
  return null
}

const seen = { open: false }
function Reader() {
  seen.open = useEscOverlayOpen()
  return null
}

const readerRoot = createRoot(document.createElement('div'))
act(() => readerRoot.render(createElement(Reader)))

let root: Root | null = null
function mount(open: boolean, onEscape?: () => void): Root {
  root ??= createRoot(document.createElement('div'))
  const r = root
  act(() => r.render(createElement(Overlay, { open, onEscape })))
  return r
}

function unmount(): void {
  if (root) act(() => root!.unmount())
  root = null
}

function pressEscape(key = 'Escape'): void {
  act(() => void window.dispatchEvent(new KeyboardEvent('keydown', { key })))
}

afterEach(unmount)

describe('useEscOverlay', () => {
  it('reports an open overlay until it unmounts', () => {
    mount(true)
    expect(seen.open).toBe(true)

    unmount()
    expect(seen.open).toBe(false)
  })

  it('does not count a closed overlay', () => {
    const onEscape = vi.fn()
    mount(false, onEscape)

    pressEscape()

    expect(seen.open).toBe(false)
    expect(onEscape).not.toHaveBeenCalled()
  })

  it('invokes the latest onEscape on Escape only', () => {
    const first = vi.fn()
    const second = vi.fn()
    mount(true, first)
    pressEscape('Enter')
    expect(first).not.toHaveBeenCalled()

    mount(true, second)
    pressEscape()

    expect(first).not.toHaveBeenCalled()
    expect(second).toHaveBeenCalledTimes(1)
  })

  it('stops listening once the overlay closes', () => {
    const onEscape = vi.fn()
    mount(true, onEscape)
    mount(false, onEscape)

    pressEscape()

    expect(onEscape).not.toHaveBeenCalled()
    expect(seen.open).toBe(false)
  })
})
