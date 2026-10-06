import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'

import { EditorView } from '@codemirror/view'

import { PlainTextEditor } from '../src/renderer/source/PlainTextEditor'

/**
 * The right-click gesture on the source-text surface.
 *
 * The block editor's own menu cannot serve here: it hangs off
 * `editor.view.dom`, and this editor is mounted *instead* of that one, so the
 * listener would never fire. That is the whole reason this handler exists, and
 * the reason it has its own test — a source-text file with no way to mark it is
 * exactly the case the gesture is for.
 *
 * The browser's own menu is suppressed only when there is something to offer,
 * so a right-click on empty space still gives the reader Cut/Copy/Paste.
 */

let root: Root | null = null
let host: HTMLElement | null = null

beforeEach(() => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
  host = document.createElement('div')
  document.body.appendChild(host)
})
afterEach(() => {
  act(() => root?.unmount())
  root = null
  host?.remove()
  host = null
  vi.unstubAllGlobals()
})

/// Mount the editor over `text` and select `selected` in it.
function mount(
  text: string,
  selected: string | null,
  onSelectionRequest?: (value: string) => void,
) {
  root = createRoot(host!)
  act(() => {
    root!.render(
      createElement(PlainTextEditor, {
        initialText: text,
        mode: 'plain',
        spellcheck: false,
        onChange: () => {},
        onSelectionRequest,
      }),
    )
  })
  const cm = host!.querySelector('.cm-content') as HTMLElement
  expect(cm, 'the editor did not mount').not.toBeNull()
  const view = EditorView.findFromDOM(cm)
  expect(view, 'no EditorView on the mounted content').not.toBeNull()
  if (selected !== null) {
    const at = text.indexOf(selected)
    act(() => {
      view!.dispatch({
        selection: { anchor: at, head: at + selected.length },
      })
    })
  }
  return { cm, view: view! }
}

/** Right-click the editor, reporting whether the browser menu was suppressed. */
function rightClick(cm: HTMLElement): boolean {
  let prevented = false
  act(() => {
    const event = new MouseEvent('contextmenu', { bubbles: true, cancelable: true })
    cm.dispatchEvent(event)
    prevented = event.defaultPrevented
  })
  return prevented
}

const ENV = 'DB_HOST=localhost\nDB_PASSWORD=hunter2\n'

describe('right-clicking a selection in the source-text surface', () => {
  it('reports the selected text', () => {
    const onRequest = vi.fn()
    const { cm } = mount(ENV, 'hunter2', onRequest)
    rightClick(cm)
    expect(onRequest).toHaveBeenCalledWith('hunter2')
  })

  it('takes the browser menu over, so the two do not both appear', () => {
    const { cm } = mount(ENV, 'hunter2', vi.fn())
    expect(rightClick(cm)).toBe(true)
  })

  it('leaves the browser menu alone when nothing is selected', () => {
    // With no selection there is nothing to withhold, and suppressing the
    // browser menu there would take away Cut/Copy/Paste for no reason.
    const onRequest = vi.fn()
    const { cm } = mount(ENV, null, onRequest)
    expect(rightClick(cm)).toBe(false)
    expect(onRequest).not.toHaveBeenCalled()
  })

  it('does nothing at all when the app supplies no handler', () => {
    const { cm } = mount(ENV, 'hunter2', undefined)
    expect(rightClick(cm)).toBe(false)
  })

  it('ignores a selection that is only whitespace', () => {
    // Marking a blank run would write a mark line over nothing.
    const onRequest = vi.fn()
    const { cm } = mount(ENV + '\n   \n', '   ', onRequest)
    rightClick(cm)
    expect(onRequest).not.toHaveBeenCalled()
  })
})
