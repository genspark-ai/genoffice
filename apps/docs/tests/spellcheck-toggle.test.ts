// Review → Spelling shows or hides the red lines under misspelled words: the
// ribbon button flips the persisted pref.
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { computeFormatState } from '../src/renderer/components/ribbon-format-state'
import { Ribbon } from '../src/renderer/components/Ribbon'
import { SPELLCHECK_KEY, spellcheckEnabled } from '../src/renderer/spellcheck-pref'
import { t } from '../src/renderer/i18n/locale'
import { ribbonProps } from './helpers/ribbon-props'

function makeEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'hello world' }] }],
    },
  })
}

describe('spellcheck preference', () => {
  afterEach(() => localStorage.removeItem(SPELLCHECK_KEY))

  it('defaults to enabled and follows the stored flag', () => {
    expect(spellcheckEnabled()).toBe(true)
    localStorage.setItem(SPELLCHECK_KEY, '0')
    expect(spellcheckEnabled()).toBe(false)
    localStorage.setItem(SPELLCHECK_KEY, '1')
    expect(spellcheckEnabled()).toBe(true)
  })
})

describe('Review → Spelling toggle', () => {
  let editor: Editor
  let root: Root
  let container: HTMLElement
  const onCommand = vi.fn()

  function renderRibbon(spellcheck: boolean) {
    act(() =>
      root.render(
        createElement(Ribbon, {
          ...ribbonProps(editor, computeFormatState(editor)),
          spellcheck,
          onCommand,
        }),
      ),
    )
  }

  const spellingButton = () =>
    [...container.querySelectorAll<HTMLButtonElement>('.rb-small')].find(
      // the icon's SVG text glyph ("abc") is part of textContent, so match its tail
      (b) => b.textContent?.endsWith(t('ribbonSpellcheckBtn')),
    )

  beforeEach(() => {
    editor = makeEditor()
    container = document.createElement('div')
    document.body.appendChild(container)
    root = createRoot(container)
  })

  afterEach(() => {
    act(() => root.unmount())
    container.remove()
    editor.destroy()
    onCommand.mockClear()
  })

  it('shows an active toggle that runs the spelling command on click', () => {
    renderRibbon(true)
    const reviewTab = [...container.querySelectorAll<HTMLButtonElement>('.ribbon-tab')].find(
      (b) => b.textContent === t('ribbonTabReview'),
    )!
    act(() => reviewTab.click())
    const btn = spellingButton()!
    expect(btn.className).toContain('active')
    act(() => btn.click())
    expect(onCommand).toHaveBeenCalledWith('toggle-spellcheck')
  })

  it('renders inactive when spellcheck is off and re-enables on click', () => {
    renderRibbon(false)
    const reviewTab = [...container.querySelectorAll<HTMLButtonElement>('.ribbon-tab')].find(
      (b) => b.textContent === t('ribbonTabReview'),
    )!
    act(() => reviewTab.click())
    const btn = spellingButton()!
    expect(btn.className).not.toContain('active')
    act(() => btn.click())
    expect(onCommand).toHaveBeenCalledWith('toggle-spellcheck')
  })
})
