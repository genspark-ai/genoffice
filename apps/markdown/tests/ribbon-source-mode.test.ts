import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { Ribbon } from '../src/renderer/components/Ribbon'

// The assistant is unrelated to the formatting buttons under test.
vi.mock('../src/renderer/ai/AiPanel', () => ({ GensparkMark: () => null }))

beforeEach(() => vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true))
const cleanups: Array<() => void> = []
afterEach(() => {
  cleanups.splice(0).forEach((cleanup) => cleanup())
  vi.unstubAllGlobals()
})

function renderRibbon(sourceMode: boolean) {
  const editor = new Editor({
    extensions: buildExtensions({
      slashController: { onOpen() {}, onUpdate() {}, onKeyDown: () => false, onClose() {} },
      slashItems: () => [],
    }),
    content: '<p>body</p>',
  })
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => {
    root.render(
      createElement(Ribbon, {
        editor,
        disabled: false,
        dirty: false,
        onSave: vi.fn(),
        onSaveAs: vi.fn(),
        onFind: vi.fn(),
        autoSave: false,
        onToggleAutoSave: vi.fn(),
        imageEnabled: true,
        onInsertImage: vi.fn(),
        frontmatterOpen: false,
        onToggleFrontmatter: vi.fn(),
        outlineOpen: false,
        onToggleOutline: vi.fn(),
        hasOutline: false,
        spellcheck: true,
        onToggleSpellcheck: vi.fn(),
        aiOpen: false,
        onToggleAi: vi.fn(),
        onAiPreset: vi.fn(),
        sourceMode,
      }),
    )
  })
  cleanups.push(() => {
    act(() => root.unmount())
    container.remove()
    editor.destroy()
  })
  return container
}

/**
 * The style dropdown is the one formatting control with its own class, and it
 * only makes sense for a block document. The remaining formatting buttons are
 * plain icon buttons, so they are counted in the toolbar body instead — a
 * source-mode ribbon keeps only the spellcheck button down there.
 */
function counts(container: HTMLElement) {
  return {
    styleDropdown: container.querySelectorAll('.rb-style').length,
    // every icon button in the toolbar body, minus the one the source ribbon keeps
    iconButtons: container.querySelectorAll('.ribbon-body button.rb-btn').length,
  }
}

describe('Ribbon source mode', () => {
  it('drops the block-formatting controls for a .txt/.json file', () => {
    const source = counts(renderRibbon(true))
    const markdown = counts(renderRibbon(false))
    expect(markdown.styleDropdown).toBe(1)
    expect(source.styleDropdown).toBe(0)
    // 14, plus this feature's withhold toggle — which stays in source mode,
    // because it gates the right-click gesture and that now works there too.
    expect(markdown.iconButtons).toBe(15)
    // source mode keeps the spellcheck toggle and the withhold toggle
    expect(source.iconButtons).toBe(2)
  })

  it('keeps the actions that still apply to source text', () => {
    const container = renderRibbon(true)
    // the quick-access row is untouched: save, save-as, undo, redo, find
    expect(container.querySelectorAll('.ribbon-tabs .qa-btn').length).toBe(5)
    expect(container.querySelector('.autosave-toggle')).not.toBeNull()
  })

  it('shows the formatting controls again for a markdown file', () => {
    const { styleDropdown, iconButtons } = counts(renderRibbon(false))
    expect(styleDropdown).toBe(1)
    expect(iconButtons).toBe(15)
  })
})
