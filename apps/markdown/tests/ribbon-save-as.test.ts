import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { Ribbon } from '../src/renderer/components/Ribbon'

// The assistant is unrelated to quick-access file actions.
vi.mock('../src/renderer/ai/AiPanel', () => ({ GensparkMark: () => null }))

beforeEach(() => vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true))
const cleanups: Array<() => void> = []
afterEach(() => {
  cleanups.splice(0).forEach((cleanup) => cleanup())
  vi.unstubAllGlobals()
})

function renderRibbon(disabled = false) {
  const editor = new Editor({
    extensions: buildExtensions({
      slashController: { onOpen() {}, onUpdate() {}, onKeyDown: () => false, onClose() {} },
      slashItems: () => [],
    }),
    content: '<p>Saved document</p>',
  })
  const onSave = vi.fn()
  const onSaveAs = vi.fn()
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  const props = {
    disabled,
    dirty: false,
    onSave,
    onSaveAs,
    onFind: vi.fn(),
    autoSave: false,
    onToggleAutoSave: vi.fn(),
    aiOpen: false,
    onToggleAi: vi.fn(),
    onAiPreset: vi.fn(),
    editor,
    imageEnabled: true,
    onInsertImage: vi.fn(),
    frontmatterOpen: false,
    onToggleFrontmatter: vi.fn(),
    outlineOpen: false,
    onToggleOutline: vi.fn(),
    hasOutline: false,
  }
  act(() => root.render(createElement(Ribbon, props)))
  cleanups.push(() => {
    act(() => root.unmount())
    editor.destroy()
    container.remove()
  })
  return { container, onSave, onSaveAs }
}

function saveAsButton(container: HTMLElement): HTMLButtonElement {
  const button = container.querySelector<HTMLButtonElement>(
    '.ribbon-tabs button[aria-label="\u53e6\u5b58\u4e3a\u2026"]',
  )
  expect(button, 'Save As must be available in the top-left quick-access row').not.toBeNull()
  return button!
}

describe('Save As quick-access button', () => {
  it('can save a copy of an unchanged document without triggering normal save', () => {
    const { container, onSave, onSaveAs } = renderRibbon()
    const button = saveAsButton(container)
    expect(button.textContent).toBe('\u53e6\u5b58\u4e3a\u2026')
    expect(button.disabled).toBe(false)
    act(() => button.click())
    expect(onSaveAs).toHaveBeenCalledOnce()
    expect(onSave).not.toHaveBeenCalled()
  })

  it('does not allow Save As while the document is unavailable', () => {
    const { container, onSaveAs } = renderRibbon(true)
    const button = saveAsButton(container)
    expect(button.disabled).toBe(true)
    act(() => button.click())
    expect(onSaveAs).not.toHaveBeenCalled()
  })
})
