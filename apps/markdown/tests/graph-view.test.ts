import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { GraphView } from '../src/renderer/components/GraphView'

/**
 * The graph overlay covers the ribbon strip, whose native drag region swallows
 * the presses the close button needs (the popover-open row flip exists for the
 * same reason). While the overlay is up <html> must carry the
 * `markdown-graph-open` marker (rows yield to presses through it), and a press
 * on the ✕ — or Escape, the way any other full-screen view dismisses — has to
 * reach onClose. genoffice#2134
 */

const onClose = vi.fn()

async function renderGraph(): Promise<void> {
  const container = document.createElement('div')
  document.body.appendChild(container)
  root = createRoot(container)
  await act(async () =>
    root!.render(createElement(GraphView, { current: null, onClose, onOpen: () => {} })),
  )
}

let root: Root | null = null

beforeEach(() => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
  onClose.mockClear()
})

afterEach(() => {
  if (root) act(() => root!.unmount())
  root = null
  document.body.innerHTML = ''
  vi.unstubAllGlobals()
})

describe('GraphView close paths (genoffice#2134)', () => {
  it('a press on the close button calls onClose', async () => {
    await renderGraph()
    const button = document.querySelector<HTMLButtonElement>('.graph-close')
    expect(button).not.toBeNull()
    await act(async () => {
      button!.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })
    expect(onClose).toHaveBeenCalledTimes(1)
  })

  it('Escape closes', async () => {
    await renderGraph()
    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }))
    })
    expect(onClose).toHaveBeenCalledTimes(1)
    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }))
    })
    expect(onClose).toHaveBeenCalledTimes(1)
  })

  it('marks the graph open on <html> while mounted, and clears it on unmount', async () => {
    await renderGraph()
    expect(document.documentElement.classList.contains('markdown-graph-open')).toBe(true)
    act(() => root!.unmount())
    root = null
    expect(document.documentElement.classList.contains('markdown-graph-open')).toBe(false)
  })
})
