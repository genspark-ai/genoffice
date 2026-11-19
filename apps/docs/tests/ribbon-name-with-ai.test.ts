import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { computeFormatState } from '../src/renderer/components/ribbon-format-state'
import { Ribbon } from '../src/renderer/components/Ribbon'
import { LocaleProvider, setModuleLang, t } from '../src/renderer/i18n/locale'
import { ribbonProps } from './helpers/ribbon-props'

/**
 * The manual trigger's only entry point.
 *
 * This button is the whole feature from the reader's side, and it shipped
 * detached once already: the naming function existed, the projection existed,
 * and nothing on screen reached either. A test of the naming function passes
 * just as happily when no button is wired to it, so the wiring needs its own.
 */

let container: HTMLDivElement
let root: Root
let editor: Editor
const onNameWithAi = vi.fn()

function makeEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'Quarterly notes' }] }],
    },
  })
}

function openFileMenu(): void {
  act(() => {
    root.render(
      createElement(LocaleProvider, {
        initial: 'en',
        children: createElement(Ribbon, {
          ...ribbonProps(editor, computeFormatState(editor)),
          onNameWithAi,
        }),
      }),
    )
  })
  const tab = container.querySelector<HTMLButtonElement>('.ribbon-tab-file')
  act(() => {
    tab!.dispatchEvent(new MouseEvent('click', { bubbles: true }))
  })
}

function nameButton(): HTMLButtonElement | null {
  return container.querySelector<HTMLButtonElement>('.qa-name-with-ai')
}

beforeEach(() => {
  onNameWithAi.mockReset()
  ;(window as unknown as { desktop: unknown }).desktop = { onLanguageChanged: () => () => {} }
  container = document.createElement('div')
  document.body.appendChild(container)
  root = createRoot(container)
  editor = makeEditor()
  setModuleLang('en')
  openFileMenu()
})

afterEach(() => {
  act(() => root.unmount())
  container.remove()
  editor.destroy()
  setModuleLang('zh')
})

describe('File ▸ the manual naming trigger', () => {
  it('is there, labelled, and enabled once a document is open', () => {
    const button = nameButton()
    expect(button).not.toBeNull()
    expect(button!.textContent).toBe(t('ribbonNameWithAi'))
    expect(button!.disabled).toBe(false)
  })

  it('asks the model when pressed', () => {
    const button = nameButton()
    expect(button).not.toBeNull()
    act(() => {
      button!.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })
    expect(onNameWithAi).toHaveBeenCalledTimes(1)
  })

  it('closes the menu, so the dialog that follows is not behind it', () => {
    act(() => {
      nameButton()!.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })
    expect(nameButton()).toBeNull()
  })

  it('is disabled with no document, rather than asking about nothing', () => {
    act(() => {
      root.render(
        createElement(LocaleProvider, {
          initial: 'en',
          children: createElement(Ribbon, {
            ...ribbonProps(editor, computeFormatState(editor)),
            hasDoc: false,
            onNameWithAi,
          }),
        }),
      )
    })
    // the File menu is open already, so this only re-reads the button
    expect(nameButton()?.disabled).toBe(true)
    act(() => {
      nameButton()!.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })
    expect(onNameWithAi).not.toHaveBeenCalled()
  })
})
