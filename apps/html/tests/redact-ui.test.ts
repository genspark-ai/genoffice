import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { SourceEditor, type SourceEditorHandle } from '../src/renderer/source/SourceEditor'
import { RedactDialog } from '../src/renderer/components/RedactDialog'
import { t } from '../src/renderer/i18n/locale'

/**
 * The reader-facing surface: a right-click menu in the source pane, and the dialog
 * that asks what a withheld span stands for.
 *
 * Both are the half of the feature a reader touches, and both have one job that
 * cannot be faked by "the component rendered" — the menu must hand over the exact
 * range and the exact text that were selected, and the dialog must hand back the
 * label the reader typed. A menu wired to a stale selection, or a dialog that
 * confirms the seed instead of the edit, passes a render check and fails here.
 */

beforeEach(() => vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true))
const cleanups: Array<() => void> = []
afterEach(() => {
  cleanups.splice(0).forEach((cleanup) => cleanup())
  vi.unstubAllGlobals()
})

function mount(element: React.ReactElement): HTMLElement {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root: Root = createRoot(container)
  act(() => root.render(element))
  cleanups.push(() => {
    act(() => root.unmount())
    container.remove()
  })
  return container
}

/** type into a controlled input the way a person does: the value setter, then the event */
function type(input: HTMLInputElement, value: string) {
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set
  setter?.call(input, value)
  act(() => {
    input.dispatchEvent(new Event('input', { bubbles: true }))
  })
}

const SECRET = '13800138000'
const HTML = `<p>Call ${SECRET} now</p>`
const FROM = HTML.indexOf(SECRET)
const TO = FROM + SECRET.length

/** the source pane with a real CodeMirror behind it, and the handle that reads it */
function mountEditor(onRedact: (range: { from: number; to: number; text: string }) => void) {
  const ref = { current: null as SourceEditorHandle | null }
  const container = mount(
    createElement(SourceEditor, {
      ref,
      className: 'source-editor',
      initialText: HTML,
      onChange: () => {},
      onCursor: () => {},
      onRedact,
    }),
  )
  return { container, ref, editor: container.querySelector('.source-editor') as HTMLElement }
}

/** right-click the editor; returns the event so a test can see whether we suppressed it */
function rightClick(editor: HTMLElement) {
  const event = new MouseEvent('contextmenu', {
    bubbles: true,
    cancelable: true,
    clientX: 12,
    clientY: 20,
  })
  act(() => {
    editor.dispatchEvent(event)
  })
  return event
}

const item = (container: HTMLElement) => container.querySelector('.source-ctx-item')

describe('the source pane context menu', () => {
  it('offers the command over a selection and hands over that exact range and text', () => {
    const onRedact = vi.fn()
    const { container, ref, editor } = mountEditor(onRedact)
    act(() => ref.current?.revealRange(FROM, TO, false))
    const event = rightClick(editor)

    // the browser's own menu must not also open over ours
    expect(event.defaultPrevented).toBe(true)
    const button = item(container) as HTMLButtonElement | null
    expect(button).not.toBeNull()
    expect(button?.textContent).toBe(t('redactMenuLabel'))
    act(() => button?.click())
    expect(onRedact).toHaveBeenCalledWith({ from: FROM, to: TO, text: SECRET })
    // and the menu is gone, so the next right-click starts from a clean state
    expect(item(container)).toBeNull()
  })

  it('says nothing on a bare caret, and leaves the native menu alone', () => {
    const onRedact = vi.fn()
    const { container, editor } = mountEditor(onRedact)
    // the editor is really there: a component that failed to mount would also
    // have no item, and this assertion is what tells the two apart
    expect(container.querySelector('.cm-editor')).not.toBeNull()
    expect(ref0SelectionIsEmpty(container)).toBe(true)

    const event = rightClick(editor)
    expect(item(container)).toBeNull()
    expect(event.defaultPrevented).toBe(false)
    expect(onRedact).not.toHaveBeenCalled()
  })

  it('closes on Escape without claiming the selection', () => {
    const onRedact = vi.fn()
    const { container, ref, editor } = mountEditor(onRedact)
    act(() => ref.current?.revealRange(FROM, TO, false))
    rightClick(editor)
    expect(item(container)).not.toBeNull()

    act(() => {
      window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
    })
    expect(item(container)).toBeNull()
    expect(onRedact).not.toHaveBeenCalled()
  })
})

/** the freshly mounted editor has an empty selection — the state the caret case needs */
function ref0SelectionIsEmpty(container: HTMLElement): boolean {
  return container.querySelectorAll('.cm-selectionBackground').length === 0
}

describe('the redact dialog', () => {
  function open(seed = SECRET) {
    const onSubmit = vi.fn()
    const onCancel = vi.fn()
    const container = mount(createElement(RedactDialog, { seed, onSubmit, onCancel }))
    const input = container.querySelector('input') as HTMLInputElement
    return {
      container,
      input,
      onSubmit,
      onCancel,
      preview: () => container.querySelector('.redact-dialog-preview')?.textContent,
      confirm: () => container.querySelector('.redact-dialog-btn.primary') as HTMLButtonElement,
      cancel: () => container.querySelector('.redact-dialog-btn') as HTMLButtonElement,
    }
  }

  it('seeds from the selection and shows the marker that label produces', () => {
    const d = open()
    expect(d.input.value).toBe(SECRET)
    expect(d.preview()).toBe(`{{${SECRET}}}`)
  })

  it('previews the edit, not the seed, and confirms the sanitized label on Enter', () => {
    const d = open()
    // a label the sanitizer has to work on, so confirming the raw field value
    // instead of the sanitized one cannot pass unnoticed
    type(d.input, '  <> client   phone  ')
    expect(d.preview()).toBe('{{client phone}}')
    act(() => {
      d.input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
    })
    expect(d.onSubmit).toHaveBeenCalledWith('client phone')
  })

  it('cancels on Escape and submits nothing', () => {
    const d = open()
    type(d.input, 'client phone')
    act(() => {
      d.input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
    })
    expect(d.onCancel).toHaveBeenCalledTimes(1)
    expect(d.onSubmit).not.toHaveBeenCalled()
  })

  it('refuses a label that sanitizes to nothing, and accepts one that does', () => {
    const d = open()
    // the assertion has to tell the two states apart, so both are checked
    expect(d.confirm().disabled).toBe(false)
    type(d.input, '<>')
    expect(d.confirm().disabled).toBe(true)
    type(d.input, '<>phone')
    expect(d.confirm().disabled).toBe(false)
    expect(d.preview()).toBe('{{phone}}')
  })

  it('keeps a bare button click equivalent to Enter', () => {
    const d = open('client phone')
    act(() => d.confirm().click())
    expect(d.onSubmit).toHaveBeenCalledWith('client phone')
  })

  it('cancels from the cancel button without submitting', () => {
    const d = open()
    act(() => d.cancel().click())
    expect(d.onCancel).toHaveBeenCalledTimes(1)
    expect(d.onSubmit).not.toHaveBeenCalled()
  })
})
