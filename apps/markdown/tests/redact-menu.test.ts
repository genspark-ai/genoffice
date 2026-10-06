import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { Redaction } from '../src/renderer/editor/Redaction'
import { RedactMenu } from '../src/renderer/components/RedactMenu'

/**
 * The menu item has to survive its own dismissal.
 *
 * The menu listens on `window` for `mousedown` so a click elsewhere closes it.
 * mousedown fires *before* click, so choosing an item used to close the menu
 * first — and because the dialog was rendered inside the same
 * `if (!point) return null` branch, it was unmounted in the same tick that
 * opened it. The command did nothing at all.
 */

const SECRET = '13800138000'

let root: Root | null = null
let editor: Editor | null = null

beforeEach(() => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
  editor = new Editor({
    element: document.createElement('div'),
    extensions: [
      ...buildExtensions({
        slashController: { onOpen() {}, onUpdate() {}, onKeyDown: () => false, onClose() {} },
        slashItems: () => [],
      }),
      Redaction,
    ],
    content: `<p>Call ${SECRET} now</p>`,
  })
  const container = document.createElement('div')
  document.body.appendChild(container)
  root = createRoot(container)
  act(() => root!.render(createElement(RedactMenu, { editor, enabled: true })))
})

afterEach(() => {
  if (root) act(() => root!.unmount())
  root = null
  editor?.destroy()
  editor = null
  document.body.innerHTML = ''
  vi.unstubAllGlobals()
})

/** select the number the way a user's drag would */
function selectSecret(): void {
  let from = -1
  editor!.state.doc.descendants((node, pos) => {
    if (from >= 0 || !node.isText) return
    const at = node.text?.indexOf(SECRET) ?? -1
    if (at !== -1) from = pos + at
  })
  expect(from).toBeGreaterThan(0)
  act(() => {
    editor!.commands.setTextSelection({ from, to: from + SECRET.length })
  })
}

const rightClick = () =>
  act(() => {
    editor!.view.dom.dispatchEvent(
      new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 40, clientY: 40 }),
    )
  })

/** the window listener is attached on the next tick after the menu opens */
const nextTick = () => new Promise((r) => setTimeout(r, 5))

describe('the right-click menu that withholds a selection', () => {
  it('stays out of the way with no selection, leaving the browser menu alone', () => {
    act(() => editor!.commands.focus('end'))
    let prevented = true
    act(() => {
      const event = new MouseEvent('contextmenu', {
        bubbles: true,
        cancelable: true,
        clientX: 40,
        clientY: 40,
      })
      editor!.view.dom.dispatchEvent(event)
      prevented = event.defaultPrevented
    })
    expect(prevented).toBe(false)
    expect(document.querySelector('.redact-menu')).toBeNull()
  })

  it('offers the command over a selection', () => {
    selectSecret()
    rightClick()
    expect(document.querySelector('.redact-menu-item')).not.toBeNull()
  })

  it('opens the dialog even though the menu dismisses on the same mousedown', async () => {
    selectSecret()
    rightClick()
    const item = document.querySelector('.redact-menu-item') as HTMLButtonElement
    expect(item).not.toBeNull()
    await nextTick()

    // the real sequence: mousedown reaches the window listener and closes the
    // menu, then click runs the item's handler
    act(() => {
      item.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }))
    })
    act(() => {
      item.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }))
    })

    expect(document.querySelector('.redact-menu')).toBeNull()
    expect(
      document.querySelector('.redact-dialog'),
      'the dialog must outlive the menu that opened it',
    ).not.toBeNull()
  })

  it('marks the selected words, keeping them in the document', async () => {
    selectSecret()
    rightClick()
    const item = document.querySelector('.redact-menu-item') as HTMLButtonElement
    await nextTick()
    act(() => {
      item.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }))
      item.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }))
    })
    const input = document.querySelector('.redact-dialog-input') as HTMLInputElement
    expect(input).not.toBeNull()
    act(() => {
      const setter = Object.getOwnPropertyDescriptor(
        window.HTMLInputElement.prototype,
        'value',
      )!.set!
      setter.call(input, 'client phone')
      input.dispatchEvent(new Event('input', { bubbles: true }))
    })
    const confirm = document.querySelector('.redact-dialog-btn.primary') as HTMLButtonElement
    act(() => {
      confirm.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }))
    })

    // the words are still there — a mark over them, never a deletion
    expect(editor!.state.doc.textContent).toContain(SECRET)
    let marked = ''
    editor!.state.doc.descendants((node) => {
      if (node.isText && node.marks.some((m) => m.type.name === 'redaction'))
        marked += node.text ?? ''
    })
    expect(marked).toBe(SECRET)
  })
})
