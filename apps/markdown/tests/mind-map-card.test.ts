import { afterEach, describe, expect, it } from 'vitest'
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

function renderCard(): { viewport: HTMLDivElement; canvas: HTMLDivElement } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  roots.push(root)
  act(() => {
    root.render(createElement(MindMapCard, { tree, onAsk: () => {}, onJumpBlock: () => {} }))
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
