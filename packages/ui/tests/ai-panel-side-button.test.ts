/** @vitest-environment jsdom */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, expect, it, vi } from 'vitest'
import { AiPanelSideButton } from '../src/AiPanelSideButton'
import { AI_PANEL_SIDE_LABELS } from '../src/strings-ai-panel-side'
import { applyAiPanelPrefs } from '../src/ai-panel-prefs-store'
import { DEFAULT_AI_PANEL_PREFS } from '../src/ai-panel-prefs'

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
  vi.restoreAllMocks()
})

it('switches sides, updates the label, and persists each click', async () => {
  const onMove = vi.fn(async (side: 'left' | 'right') => ({ ...DEFAULT_AI_PANEL_PREFS, side }))
  act(() => root.render(createElement(AiPanelSideButton, { lang: 'en', onMove })))
  const button = host.querySelector('button')!
  expect(button.getAttribute('aria-label')).toBe('Move AI panel to the right')
  await act(async () => button.click())
  expect(onMove).toHaveBeenLastCalledWith('right')
  expect(document.documentElement.dataset.aiPanelSide).toBe('right')
  expect(button.getAttribute('aria-label')).toBe('Move AI panel to the left')
  await act(async () => button.click())
  expect(onMove).toHaveBeenLastCalledWith('left')
  expect(document.documentElement.dataset.aiPanelSide).toBe('left')
})

it('keeps the existing side when saving the preference fails', async () => {
  const onMove = vi.fn(async () => {
    throw new Error('write failed')
  })
  vi.spyOn(console, 'error').mockImplementation(() => {})
  act(() => root.render(createElement(AiPanelSideButton, { lang: 'zh', onMove })))
  const button = host.querySelector('button')!
  await act(async () => button.click())
  expect(button.getAttribute('aria-label')).toBe(AI_PANEL_SIDE_LABELS.zh.right)
  expect(button.disabled).toBe(false)
  expect(document.documentElement.dataset.aiPanelSide).not.toBe('right')
})
