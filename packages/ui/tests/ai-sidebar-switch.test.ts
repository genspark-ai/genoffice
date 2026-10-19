/** @vitest-environment jsdom */
/**
 * Settings › "show the AI sidebar" is a soft hide, the way Word treats the
 * Copilot pane: an explicit Ask-AI action (ribbon, context menu, shortcut)
 * still opens the panel for the current document.
 *
 * Turning the switch off is the one thing that closes it. Without that, a panel
 * that was open when the switch went off stayed open with its rail hidden, and
 * every entry point — which only ever asks to *open* — became a no-op, so only
 * a reload recovered. All six editors share this rule through the hook; the six
 * call sites are pinned in apps/shell/tests/ai-sidebar-switch.test.ts.
 */
import { act, createElement, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { applyAiPanelPrefs, useAiSidebarEnabled } from '../src/ai-panel-prefs-store'

/** Stands in for an editor's dock: `showAi` is the panel's open state. */
function Dock() {
  const [showAi, setShowAi] = useState(true)
  const enabled = useAiSidebarEnabled(setShowAi)
  return createElement(
    'div',
    { 'data-open': String(showAi), 'data-enabled': String(enabled) },
    // the Ask-AI entry points: the only way `showAi` ever turns true
    createElement('button', { className: 'ask-ai', onClick: () => setShowAi(true) }),
    createElement('button', { className: 'collapse', onClick: () => setShowAi(false) }),
  )
}

let host: HTMLDivElement
let root: ReturnType<typeof createRoot>

beforeEach(() => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
  applyAiPanelPrefs({})
  host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
})

afterEach(() => {
  act(() => root.unmount())
  host.remove()
  applyAiPanelPrefs({})
  vi.unstubAllGlobals()
})

function open(): boolean {
  return host.firstElementChild!.getAttribute('data-open') === 'true'
}

function click(selector: string): void {
  act(() => host.querySelector<HTMLButtonElement>(selector)!.click())
}

describe('useAiSidebarEnabled', () => {
  it('closes an open panel when the switch turns off', () => {
    act(() => root.render(createElement(Dock)))
    expect(open()).toBe(true)

    act(() => applyAiPanelPrefs({ showAiSidebar: false }))
    expect(open()).toBe(false)
    expect(host.firstElementChild!.getAttribute('data-enabled')).toBe('false')
  })

  it('closes a panel that starts open while the switch is already off', () => {
    applyAiPanelPrefs({ showAiSidebar: false })
    act(() => root.render(createElement(Dock)))
    expect(open()).toBe(false)
  })

  it('still opens on an explicit AI action while the switch is off, and closes again on the rail', () => {
    act(() => root.render(createElement(Dock)))
    act(() => applyAiPanelPrefs({ showAiSidebar: false }))
    expect(open()).toBe(false)

    // soft hide: the ribbon / context-menu Ask-AI entry point opens the panel
    click('.ask-ai')
    expect(open()).toBe(true)

    click('.collapse')
    expect(open()).toBe(false)
  })

  it('leaves the panel closed when the switch comes back on', () => {
    act(() => root.render(createElement(Dock)))
    act(() => applyAiPanelPrefs({ showAiSidebar: false }))
    act(() => applyAiPanelPrefs({ showAiSidebar: true }))
    expect(open()).toBe(false)
    expect(host.firstElementChild!.getAttribute('data-enabled')).toBe('true')
  })

  it('leaves the panel alone while the switch is on', () => {
    act(() => root.render(createElement(Dock)))
    click('.collapse')
    expect(open()).toBe(false)
    act(() => applyAiPanelPrefs({ showAiSidebar: true }))
    expect(open()).toBe(false)
    click('.ask-ai')
    expect(open()).toBe(true)
  })
})
