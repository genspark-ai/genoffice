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

/**
 * @param sourceMode the .txt/.json "this file is source text" flag from genoffice#1848
 * @param sourceViewOpen this PR's markdown source view, kept a separate flag
 *   because the two mean different things and can never both be on
 */
function renderRibbon(sourceMode: boolean, sourceViewOpen = false) {
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
        sourceViewOpen,
        onToggleSource: vi.fn(),
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
    // 14 block-formatting controls, plus this PR's withhold toggle, plus the
    // markdown-only source-view toggle (absent for a .txt/.json, which has no
    // use for it), plus the image-host button, which is not a formatting
    // control and so renders in either mode.
    expect(markdown.iconButtons).toBe(17)
    // source mode keeps exactly one: the spellcheck toggle (the withhold
    // toggle gates the block editor's marks, which source mode does not carry)
    expect(source.iconButtons).toBe(1)
  })

  it('keeps the actions that still apply to source text', () => {
    const container = renderRibbon(true)
    // the quick-access row is untouched: save, save-as, undo, redo, find
    expect(container.querySelectorAll('.ribbon-tabs .qa-btn').length).toBe(5)
    expect(container.querySelector('.autosave-toggle')).not.toBeNull()
  })

  it('disables the block-formatting controls in the markdown source view', () => {
    // A different situation from a .txt/.json, and so a different answer. Here
    // the document is still live behind the pane — a keystroke in it is parsed
    // straight back into the editor — so the controls are disabled rather than
    // removed: they act on a selection the reader cannot see right now, and
    // removing them would make the ribbon jump on every toggle.
    const open = renderRibbon(false, true)
    const plain = renderRibbon(false)
    // Label-free on purpose: this suite renders in the default locale, so an
    // aria-label lookup would be asserting on a translation. What matters is
    // how many controls are live, and whether the ribbon lost any of them.
    const disabledIn = (container: HTMLElement) =>
      container.querySelectorAll('button.rb-btn:disabled').length
    // Relative, not absolute: undo and redo are already disabled on a fresh
    // document, so the claim is that opening the view quiets MORE of them.
    expect(disabledIn(open)).toBeGreaterThan(disabledIn(plain))
    // nothing was removed: the reader's ribbon does not jump on toggle, and
    // the controls come back enabled when the view closes
    expect(counts(open).iconButtons).toBe(counts(plain).iconButtons)
  })

  it('keeps Save As live in both modes and Find live only outside the source view', () => {
    const quick = (container: HTMLElement) =>
      Array.from(container.querySelectorAll<HTMLButtonElement>('.ribbon-tabs .qa-btn'))
    const [, saveAsPlain, , , findPlain] = quick(renderRibbon(false))
    const [, saveAsSource, , , findSource] = quick(renderRibbon(true))
    const [, saveAsView, , , findView] = quick(renderRibbon(false, true))
    expect(saveAsPlain!.disabled).toBe(false)
    expect(saveAsSource!.disabled).toBe(false)
    expect(saveAsView!.disabled).toBe(false)
    expect(findPlain!.disabled).toBe(false)
    expect(findSource!.disabled).toBe(false)
    expect(findView!.disabled).toBe(true)
  })

  it('shows the formatting controls again for a markdown file', () => {
    const { styleDropdown, iconButtons } = counts(renderRibbon(false))
    expect(styleDropdown).toBe(1)
    // 17 here too: the source-view toggle is present here and absent above,
    // the withhold toggle and the image-host button are present in both
    expect(iconButtons).toBe(17)
  })
})
