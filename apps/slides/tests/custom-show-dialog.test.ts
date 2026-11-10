import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'

import { CustomShowDialog, isShowSlideIndex } from '../src/renderer/components/CustomShowDialog'
import type { CustomShow } from '../src/renderer/slideshow-utils'

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

const roots: Array<{ root: Root; container: HTMLElement }> = []
afterEach(() => {
  for (const { root, container } of roots.splice(0)) {
    act(() => root.unmount())
    container.remove()
  }
})

const shows: CustomShow[] = [
  { id: 'a', name: 'Alpha', slideIndices: [0, 1, 99, -1] },
  { id: 'b', name: 'Beta', slideIndices: [2] },
]

async function mount() {
  const onChange = vi.fn()
  const onPlay = vi.fn()
  const onClose = vi.fn()
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  await act(async () => {
    root.render(
      createElement(CustomShowDialog, { shows, slideCount: 3, onChange, onPlay, onClose }),
    )
  })
  roots.push({ root, container })
  return { container, onChange, onPlay, onClose }
}

describe('isShowSlideIndex', () => {
  it('accepts only integers inside the live page range', () => {
    expect(isShowSlideIndex(0, 3)).toBe(true)
    expect(isShowSlideIndex(2, 3)).toBe(true)
    expect(isShowSlideIndex(-1, 3)).toBe(false)
    expect(isShowSlideIndex(3, 3)).toBe(false)
    expect(isShowSlideIndex(99, 3)).toBe(false)
    expect(isShowSlideIndex(1.5, 3)).toBe(false)
    expect(isShowSlideIndex('1', 3)).toBe(false)
  })
})

describe('CustomShowDialog semantics', () => {
  it('exposes a dialog with listbox options and selected state', async () => {
    const { container } = await mount()
    const dlg = container.querySelector('[role="dialog"]')
    expect(dlg?.getAttribute('aria-modal')).toBe('true')
    expect(dlg?.getAttribute('aria-labelledby')).toBeTruthy()
    const options = Array.from(container.querySelectorAll('[role="option"]'))
    expect(options).toHaveLength(2)
    expect(options[0]?.getAttribute('aria-selected')).toBe('true')
    expect(options[1]?.getAttribute('aria-selected')).toBe('false')
  })

  it('counts only in-range pages and moves selection with arrow keys', async () => {
    const { container } = await mount()
    // Alpha holds [0,1,99,-1]: badge counts 2, order list shows 2 rows
    const badge = container.querySelector('.csd-item-count')
    expect(badge?.textContent).toMatch(/2/)
    const listbox = container.querySelector('[role="listbox"]') as HTMLElement
    act(() => {
      listbox.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }))
    })
    const options = Array.from(container.querySelectorAll('[role="option"]'))
    expect(options[1]?.getAttribute('aria-selected')).toBe('true')
  })
})
