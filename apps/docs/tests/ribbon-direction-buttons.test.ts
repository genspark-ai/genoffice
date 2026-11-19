// Word shows the LTR/RTL paragraph buttons only with an RTL editing language;
// here that means an RTL UI language or an RTL paragraph under the selection.
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import type { Lang } from '@genoffice/i18n'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { computeFormatState } from '../src/renderer/components/ribbon-format-state'
import { Ribbon } from '../src/renderer/components/Ribbon'
import { LocaleProvider, setModuleLang, t } from '../src/renderer/i18n/locale'
import { ribbonProps } from './helpers/ribbon-props'

function makeEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        { type: 'docParagraph', content: [{ type: 'text', text: 'hello world' }] },
        {
          type: 'docParagraph',
          attrs: { bidi: true },
          content: [{ type: 'text', text: 'שלום' }],
        },
      ],
    },
  })
}

describe('paragraph direction buttons', () => {
  let editor: Editor
  let root: Root
  let container: HTMLElement

  beforeEach(() => {
    editor = makeEditor()
    container = document.createElement('div')
    document.body.appendChild(container)
    root = createRoot(container)
    ;(window as unknown as { desktop: unknown }).desktop = { onLanguageChanged: () => () => {} }
  })

  afterEach(() => {
    act(() => root.unmount())
    container.remove()
    editor.destroy()
    setModuleLang('zh')
  })

  function render(lang: Lang) {
    setModuleLang(lang)
    act(() =>
      root.render(
        createElement(LocaleProvider, {
          initial: lang,
          children: createElement(Ribbon, ribbonProps(editor, computeFormatState(editor))),
        }),
      ),
    )
    return container.querySelector(`[aria-label="${t('ribbonDirRtlTip')}"]`)
  }

  function placeCursor(pos: number) {
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, pos)))
  }

  it('hides them in an LTR paragraph under an LTR UI language', () => {
    placeCursor(2)
    expect(render('en')).toBeNull()
  })

  it('shows them when the cursor sits in an RTL paragraph', () => {
    placeCursor(editor.state.doc.content.size - 2)
    expect(render('en')).not.toBeNull()
  })

  it('always shows them under an RTL UI language', () => {
    placeCursor(2)
    expect(render('he')).not.toBeNull()
  })
})
