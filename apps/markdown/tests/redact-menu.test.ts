import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { hasRedactionIn, Redaction } from '../src/renderer/editor/Redaction'
import { RedactMenu } from '../src/renderer/components/RedactMenu'
import { t } from '../src/renderer/i18n/locale'

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

/** where the number starts, found by walking the doc like a search would */
function secretFrom(): number {
  let from = -1
  editor!.state.doc.descendants((node, pos) => {
    if (from >= 0 || !node.isText) return
    const at = node.text?.indexOf(SECRET) ?? -1
    if (at !== -1) from = pos + at
  })
  expect(from).toBeGreaterThan(0)
  return from
}

/** select the number the way a user's drag would */
function selectSecret(): void {
  const from = secretFrom()
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

/**
 * Labels are read through the app's own translator rather than pasted in, so
 * this file carries no translated text of its own to drift out of step.
 */
const HIDE_LABEL = t('redactMenuLabel')
const SHOW_LABEL = t('redactShowLabel')

const menuItems = () => [...document.querySelectorAll<HTMLButtonElement>('.redact-menu-item')]

const menuItem = (label: string) =>
  menuItems().find((b) => b.textContent === label) as HTMLButtonElement | undefined

/** put the redaction mark over the number, the way the menu would */
function withholdSecret(): void {
  selectSecret()
  act(() => {
    editor!.commands.setRedaction('client phone')
  })
}

/** every character currently carrying the redaction mark */
function markedText(): string {
  let marked = ''
  editor!.state.doc.descendants((node) => {
    if (node.isText && node.marks.some((m) => m.type.name === 'redaction'))
      marked += node.text ?? ''
  })
  return marked
}

/** the real sequence a click produces: mousedown, then click */
function choose(item: HTMLButtonElement): void {
  act(() => {
    item.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }))
    item.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }))
  })
}

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

describe('hasRedactionIn — is there something to stop withholding here', () => {
  it('sees a range that lies wholly inside a withheld span', () => {
    withholdSecret()
    expect(hasRedactionIn(editor!, secretFrom() + 1, secretFrom() + 4)).toBe(true)
  })

  it('sees a range that only clips the edge of a span', () => {
    // `unsetRedaction` clears the whole selection, so a partial selection is
    // still a reason to offer the entry rather than a reason to hide it
    withholdSecret()
    expect(hasRedactionIn(editor!, secretFrom() - 4, secretFrom() + 2)).toBe(true)
  })

  it('reports plain text as withheld by nothing', () => {
    selectSecret()
    expect(hasRedactionIn(editor!, secretFrom(), secretFrom() + SECRET.length)).toBe(false)
  })

  it('reports an empty range as nothing, so a bare caret is not an offer', () => {
    // the mark is `inclusive: false`, so a caret inside the span carries no
    // mark of its own and clearing "here" would be a control that does nothing
    withholdSecret()
    expect(hasRedactionIn(editor!, secretFrom() + 2, secretFrom() + 2)).toBe(false)
  })
})

describe('the right-click menu that stops withholding a selection', () => {
  it('adds the mirror item under the hide item, and only on a withheld span', () => {
    selectSecret()
    rightClick()
    // plain prose gets the one item that has something to do
    expect(menuItems().length).toBe(1)
    expect(menuItem(HIDE_LABEL)).toBeTruthy()
    expect(menuItem(SHOW_LABEL)).toBeUndefined()

    withholdSecret()
    rightClick()
    expect(menuItems().length).toBe(2)
    expect(menuItem(SHOW_LABEL)).toBeTruthy()
    // right below the item it mirrors, so the pair reads as a pair
    const names = menuItems().map((b) => b.textContent)
    expect(names.indexOf(SHOW_LABEL)).toBe(names.indexOf(HIDE_LABEL) + 1)
  })

  it('gives the words back when chosen, which is the point of the pair', async () => {
    // before this existed, the only way to stop withholding was to delete the
    // text — losing the reader's own data to undo a decision made in the editor
    withholdSecret()
    expect(markedText()).toBe(SECRET)
    rightClick()
    const entry = menuItem(SHOW_LABEL)!
    expect(entry).toBeTruthy()
    await nextTick()

    choose(entry)

    // the mark is gone, and the words were never touched: this is a mark
    // over the real text, never a replacement for it
    expect(markedText()).toBe('')
    expect(editor!.state.doc.textContent).toContain(SECRET)
    // no dialog either: there was no label to ask about
    expect(document.querySelector('.redact-dialog')).toBeNull()
  })

  it('never appears with no selection, where the browser menu is left alone', () => {
    withholdSecret()
    act(() => editor!.commands.focus('end'))
    rightClick()
    expect(document.querySelector('.redact-menu')).toBeNull()
    // and nothing was cleared behind the reader's back
    expect(markedText()).toBe(SECRET)
  })
})
