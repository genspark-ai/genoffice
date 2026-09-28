/** @vitest-environment jsdom */
import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  aiPanelInitiallyOpen,
  rememberAiPanelOpen,
  applyAiPanelPrefs,
  aiPanelWidthAtPointer,
} from '../src/ai-panel-prefs-store'

afterEach(() => {
  applyAiPanelPrefs({})
  vi.unstubAllGlobals()
})

describe('AI panel side in the renderer', () => {
  it('moves the layout on a side-only update and restores the default', () => {
    applyAiPanelPrefs({ side: 'right' })
    expect(document.documentElement.dataset.aiPanelSide).toBe('right')
    applyAiPanelPrefs({ side: 'left' })
    expect(document.documentElement.dataset.aiPanelSide).not.toBe('right')
  })

  it('measures width inward from the selected window edge', () => {
    vi.stubGlobal('innerWidth', 1200)
    applyAiPanelPrefs({ side: 'left' })
    expect(aiPanelWidthAtPointer(350)).toBe(350)
    applyAiPanelPrefs({ side: 'right' })
    expect(aiPanelWidthAtPointer(850)).toBe(350)
    expect(aiPanelWidthAtPointer(750)).toBe(450)
  })
})

describe('AI panel initial open state', () => {
  afterEach(() => localStorage.clear())

  it('follows the remembered per-app state while the setting is on', () => {
    applyAiPanelPrefs({ openInNewDocs: true })
    expect(aiPanelInitiallyOpen('k')).toBe(true)
    localStorage.setItem('k', '0')
    expect(aiPanelInitiallyOpen('k')).toBe(false)
  })

  it('starts collapsed regardless of the remembered state when off', () => {
    applyAiPanelPrefs({ openInNewDocs: false })
    localStorage.setItem('k', '1')
    expect(aiPanelInitiallyOpen('k')).toBe(false)
  })

  it('leaves the remembered state alone while off so turning it back on restores it', () => {
    applyAiPanelPrefs({ openInNewDocs: true })
    rememberAiPanelOpen('k', true)
    applyAiPanelPrefs({ openInNewDocs: false })
    rememberAiPanelOpen('k', false)
    applyAiPanelPrefs({ openInNewDocs: true })
    expect(aiPanelInitiallyOpen('k')).toBe(true)
    rememberAiPanelOpen('k', false)
    expect(aiPanelInitiallyOpen('k')).toBe(false)
  })
})
