import type { Editor } from '@tiptap/core'
import type { TabStop } from '@genoffice/docx-engine'
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { beforeAll, describe, expect, it } from 'vitest'

import { directTabStops, isRenderableTabStop, Ruler } from '../src/renderer/components/Ruler'

beforeAll(() => {
  Element.prototype.scrollTo ??= () => {}
})

const section = {
  pageWidth: 12240,
  marginLeft: 1440,
  marginRight: 1440,
} as never

/** editor stub: every paragraph query reports the given tab stops */
function editorWith(stops: TabStop[]): Editor {
  return {
    isActive: () => false,
    getAttributes: () => ({ tabStops: JSON.stringify(stops) }),
  } as unknown as Editor
}

function mountRuler(editor: Editor | null = null): {
  container: HTMLElement
  cleanup: () => void
} {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => {
    root.render(createElement(Ruler, { section, editor, onTabStopsChange: () => {} }))
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

describe('default tab guides', () => {
  const guides = (stops: TabStop[]) => {
    const { container, cleanup } = mountRuler(editorWith(stops))
    try {
      return {
        defaults: container.querySelectorAll('.ruler-tab-default').length,
        custom: container.querySelectorAll('.ruler-tab').length,
      }
    } finally {
      cleanup()
    }
  }

  it('a clear-only set (last inherited stop deleted) still shows the default grid', () => {
    expect(guides([{ pos: 709, val: 'clear' }])).toEqual({ defaults: 12, custom: 0 })
  })

  it('a real stop hides the default grid', () => {
    expect(
      guides([
        { pos: 709, val: 'clear' },
        { pos: 1440, val: 'left' },
      ]),
    ).toEqual({
      defaults: 0,
      custom: 1,
    })
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

describe('directTabStops', () => {
  const inherited = (pos: number): TabStop => ({ pos, val: 'left', inherited: true })

  it('strips inherited from stops that stay, so the set is written out as direct', () => {
    expect(directTabStops([inherited(709), { pos: 1440, val: 'right' }], [inherited(709)])).toEqual(
      [{ pos: 709, val: 'left' }],
    )
  })

  it('records a clear where an inherited stop was deleted', () => {
    expect(directTabStops([inherited(709), inherited(5752)], [inherited(5752)])).toEqual([
      { pos: 709, val: 'clear' },
      { pos: 5752, val: 'left' },
    ])
  })

  it('a moved inherited stop keeps a clear at its old position beside the new direct stop', () => {
    expect(directTabStops([inherited(709)], [{ pos: 1500, val: 'left', inherited: true }])).toEqual(
      [
        { pos: 709, val: 'clear' },
        { pos: 1500, val: 'left' },
      ],
    )
  })

  it('a direct stop landing on the inherited position needs no clear', () => {
    expect(directTabStops([inherited(709)], [{ pos: 709, val: 'center' }])).toEqual([
      { pos: 709, val: 'center' },
    ])
  })

  it('deleting a direct stop records nothing', () => {
    expect(directTabStops([{ pos: 709, val: 'left' }], [])).toEqual([])
  })
})
