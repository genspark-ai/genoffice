import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import type { MindMapTreeNode } from '@genoffice/agent-core'
import { MindMapCard } from '../src/renderer/ai/MindMapCard'

const tree: MindMapTreeNode = {
  label: 'Doc',
  children: [{ label: 'A', children: [{ label: 'A1' }] }, { label: 'B' }],
}

const roots: Root[] = []
afterEach(() => {
  for (const root of roots.splice(0)) act(() => root.unmount())
  document.body.innerHTML = ''
})

function renderCard(onInsert?: (svg: string, size: { width: number; height: number }) => void): {
  viewport: HTMLDivElement
  canvas: HTMLDivElement
} {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  roots.push(root)
  act(() => {
    root.render(
      createElement(MindMapCard, { tree, onAsk: () => {}, onJumpBlock: () => {}, onInsert }),
    )
  })
  const viewport = container.querySelector('.ai-mindmap-viewport')
  const canvas = container.querySelector('.ai-mindmap-canvas')
  if (!(viewport instanceof HTMLDivElement) || !(canvas instanceof HTMLDivElement)) {
    throw new Error('mind map card did not render a viewport')
  }
  return { viewport, canvas }
}

/** Current zoom, read back off the canvas transform. */
function zoomOf(canvas: HTMLDivElement): number {
  const m = /scale\(([\d.]+)\)/.exec(canvas.style.transform)
  return m ? Number(m[1]) : 1
}

function wheel(el: HTMLElement, init: WheelEventInit): WheelEvent {
  const event = new WheelEvent('wheel', { bubbles: true, cancelable: true, ...init })
  act(() => {
    el.dispatchEvent(event)
  })
  return event
}

describe('MindMapCard wheel gestures', () => {
  it('leaves a plain wheel to the chat panel: no zoom, no preventDefault', () => {
    const { viewport, canvas } = renderCard()

    const event = wheel(viewport, { deltaY: -240, clientX: 40, clientY: 40 })

    // Not preventDefault -> the conversation keeps scrolling past the card.
    expect(event.defaultPrevented).toBe(false)
    expect(zoomOf(canvas)).toBe(1)
  })

  it('zooms on a ctrl-wheel (trackpad pinch) gesture', () => {
    const { viewport, canvas } = renderCard()
    expect(zoomOf(canvas)).toBe(1)

    const event = wheel(viewport, { deltaY: -240, ctrlKey: true, clientX: 40, clientY: 40 })

    expect(event.defaultPrevented).toBe(true)
    expect(zoomOf(canvas)).toBeGreaterThan(1)
  })

  it('zooms on a meta-wheel gesture too, in both directions', () => {
    const { viewport, canvas } = renderCard()

    wheel(viewport, { deltaY: 240, metaKey: true })
    expect(zoomOf(canvas)).toBeLessThan(1)

    wheel(viewport, { deltaY: -240, metaKey: true })
    expect(zoomOf(canvas)).toBeCloseTo(1, 5)
  })
})

describe('MindMapCard exports', () => {
  it('the SVG button downloads the baked standalone SVG', () => {
    renderCard(() => {})
    const serializer = vi.spyOn(XMLSerializer.prototype, 'serializeToString')
    const created: Blob[] = []
    const downloads: string[] = []
    const URLRecord = URL as unknown as Record<string, unknown>
    const hadCreate = typeof URLRecord.createObjectURL === 'function'
    const hadRevoke = typeof URLRecord.revokeObjectURL === 'function'
    if (!hadCreate) URLRecord.createObjectURL = () => ''
    if (!hadRevoke) URLRecord.revokeObjectURL = () => {}
    const createObjectURL = vi
      .spyOn(URL, 'createObjectURL')
      .mockImplementation((blob: Blob | MediaSource) => {
        created.push(blob as Blob)
        return 'blob:export'
      })
    const revokeObjectURL = vi.spyOn(URL, 'revokeObjectURL').mockImplementation(() => {})
    const anchorClick = vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(function (
      this: HTMLAnchorElement,
    ) {
      downloads.push(this.download)
    })
    try {
      const buttons = document.querySelectorAll('.ai-mindmap-actions button')
      expect(buttons.length).toBe(4)
      act(() => {
        ;(buttons[2] as HTMLButtonElement).click()
      })

      expect(downloads).toEqual(['mindmap.svg'])
      expect(created.length).toBe(1)
      expect(created[0].type).toBe('image/svg+xml;charset=utf-8')
      expect(created[0].size).toBeGreaterThan(0)
      // the file is the palette-baked standalone SVG the PNG rasterizes from
      expect(serializer).toHaveBeenCalledTimes(1)
      const raw = String(serializer.mock.results[0].value)
      expect(raw).toContain('<svg')
      expect(raw).toContain('fill="#ffffff"')
      expect(raw).toContain('stroke="#d0d3d8"')
    } finally {
      createObjectURL.mockRestore()
      revokeObjectURL.mockRestore()
      anchorClick.mockRestore()
      serializer.mockRestore()
      if (!hadCreate) delete URLRecord.createObjectURL
      if (!hadRevoke) delete URLRecord.revokeObjectURL
    }
  })

  it('the insert button hands over the baked SVG and its layout size', () => {
    const onInsert = vi.fn()
    renderCard(onInsert)
    const buttons = document.querySelectorAll('.ai-mindmap-actions button')
    expect(buttons.length).toBe(4)
    act(() => {
      ;(buttons[3] as HTMLButtonElement).click()
    })

    expect(onInsert).toHaveBeenCalledTimes(1)
    const [svg, size] = onInsert.mock.calls[0]
    // the receiver gets the same palette-baked file the exports build from
    expect(String(svg)).toContain('<svg')
    expect(String(svg)).toContain('fill="#ffffff"')
    expect(size.width).toBeGreaterThan(0)
    expect(size.height).toBeGreaterThan(0)
  })
})
