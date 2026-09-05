import type { TabStop } from '@genoffice/docx-engine'
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { beforeAll, describe, expect, it } from 'vitest'

import { isRenderableTabStop, Ruler } from '../src/renderer/components/Ruler'

beforeAll(() => {
  Element.prototype.scrollTo ??= () => {}
})

const section = {
  pageWidth: 12240,
  marginLeft: 1440,
  marginRight: 1440,
} as never

function mountRuler(): { container: HTMLElement; cleanup: () => void } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => {
    root.render(createElement(Ruler, { section, editor: null, onTabStopsChange: () => {} }))
  })
  return {
    container,
    cleanup: () => {
      act(() => root.unmount())
      container.remove()
    },
  }
}

describe('ruler tab type cycle', () => {
  it('cycles L/C/R/Decimal/Bar like Word, then wraps', () => {
    const { container, cleanup } = mountRuler()
    try {
      const button = container.querySelector<HTMLButtonElement>('.ruler-tab-type')
      expect(button).not.toBeNull()
      // Glyph legend: L left, ⊥ center, ⌐ right, . decimal, | bar.
      const seen = [button!.textContent]
      for (let i = 0; i < 4; i++) {
        act(() => button!.click())
        seen.push(button!.textContent)
      }
      expect(seen).toEqual(['L', '⊥', '⌐', '.', '|'])
      act(() => button!.click())
      expect(button!.textContent).toBe('L')
    } finally {
      cleanup()
    }
  })
})

describe('isRenderableTabStop', () => {
  it('hides clear stops (they cancel inheritance, mark no position)', () => {
    const stop = (val: TabStop['val']): TabStop => ({ pos: 100, val })
    expect(isRenderableTabStop(stop('clear'))).toBe(false)
    expect(isRenderableTabStop(stop('left'))).toBe(true)
    expect(isRenderableTabStop(stop('bar'))).toBe(true)
  })
})
