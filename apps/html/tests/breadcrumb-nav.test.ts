import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { Breadcrumb } from '../src/renderer/components/Breadcrumb'
import { buildParseMap } from '../src/renderer/document/parse-map'

beforeEach(() => vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true))
const cleanups: Array<() => void> = []
afterEach(() => {
  cleanups.splice(0).forEach((cleanup) => cleanup())
  vi.unstubAllGlobals()
})

function mount(element: React.ReactElement): HTMLElement {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(element))
  cleanups.push(() => {
    act(() => root.unmount())
    container.remove()
  })
  return container
}

describe('Breadcrumb landmark and current crumb', () => {
  const text = '<div><main><p>hi</p></main></div>'
  const map = buildParseMap(text, 1)
  const target = map.elements.find((e) => e.tag === 'p') ?? map.elements[0]!

  function mountCrumbs(onSelect: (sid: number) => void = () => {}): HTMLElement {
    return mount(
      createElement(Breadcrumb, { text, map, sid: target.sid, state: 'static', onSelect }),
    )
  }

  it('renders a nav landmark carrying the translated label', () => {
    const container = mountCrumbs()
    const nav = container.querySelector('nav.crumbs')
    expect(nav).not.toBeNull()
    expect(nav!.getAttribute('aria-label')).toBeTruthy()
  })

  it('marks exactly the selected crumb as current', () => {
    const container = mountCrumbs()
    const current = container.querySelectorAll('button.crumb[aria-current="true"]')
    expect(current.length).toBe(1)
    expect(current[0]!.textContent).toContain('p')
  })

  it('clicking a crumb selects that ancestor', () => {
    const seen: number[] = []
    const container = mountCrumbs((sid) => seen.push(sid))
    const buttons = container.querySelectorAll('button.crumb')
    expect(buttons.length).toBeGreaterThan(1)
    ;(buttons[0] as HTMLButtonElement).click()
    expect(seen.length).toBe(1)
  })

  it('hides the decorative separators from assistive tech', () => {
    const container = mountCrumbs()
    for (const sep of container.querySelectorAll('.crumb-sep')) {
      expect(sep.getAttribute('aria-hidden')).toBe('true')
    }
  })
})
