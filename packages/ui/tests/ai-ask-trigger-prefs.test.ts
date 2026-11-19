/** @vitest-environment jsdom */
/**
 * The "show the AI-edit chip on selection" switch is read by every editor window
 * through useAiPanelPrefs(). Pin that the shell's broadcast reaches the hook,
 * and that flipping only this switch does not disturb the other preferences.
 */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { applyAiPanelPrefs, useAiPanelPrefs } from '../src/ai-panel-prefs-store'

function Probe() {
  const prefs = useAiPanelPrefs()
  return createElement('div', {
    'data-show-trigger': String(prefs.showAiAskTrigger),
    'data-show-sidebar': String(prefs.showAiSidebar),
  })
}

let host: HTMLDivElement
let root: ReturnType<typeof createRoot>

const flag = (name: string): string | null => host.firstElementChild!.getAttribute(`data-${name}`)

beforeEach(() => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
  applyAiPanelPrefs({})
  host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
  act(() => root.render(createElement(Probe)))
})

afterEach(() => {
  act(() => root.unmount())
  host.remove()
  applyAiPanelPrefs({})
  vi.unstubAllGlobals()
})

describe('showAiAskTrigger in the renderer', () => {
  it('is on by default', () => {
    expect(flag('show-trigger')).toBe('true')
  })

  it('reaches the hook when shell prefs are applied', () => {
    act(() => applyAiPanelPrefs({ showAiAskTrigger: false }))
    expect(flag('show-trigger')).toBe('false')
    expect(flag('show-sidebar')).toBe('true')
  })

  it('can be turned back on without a reload', () => {
    act(() => applyAiPanelPrefs({ showAiAskTrigger: false }))
    expect(flag('show-trigger')).toBe('false')
    act(() => applyAiPanelPrefs({ showAiAskTrigger: true }))
    expect(flag('show-trigger')).toBe('true')
  })
})
