import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { NodeSelection, type Selection } from '@tiptap/pm/state'
import { CellSelection, TableMap } from '@tiptap/pm/tables'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { AiAskPopover } from '../src/renderer/components/AiAskPopover'
import { LocaleProvider, setModuleLang } from '../src/renderer/i18n/locale'
import { createDocsRuntime } from '../src/renderer/runtime/docs-runtime'

Object.assign(globalThis, { IS_REACT_ACT_ENVIRONMENT: true })
const cleanup: Array<() => void> = []
afterEach(() => {
  for (const dispose of cleanup.splice(0).reverse()) dispose()
  vi.restoreAllMocks()
})

function fixture(queueEnabled = false, beforeOpen?: (editor: Editor) => void) {
  Object.assign(window, { desktop: { onLanguageChanged: () => () => {} } })
  setModuleLang('en')
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: '<p>Selected text</p>',
  })
  editor.commands.setTextSelection({ from: 1, to: 9 })
  editor.isFocused = true
  vi.spyOn(editor.view, 'coordsAtPos').mockReturnValue({ left: 10, right: 50, top: 20, bottom: 35 })
  const container = document.createElement('div')
  document.body.append(container)
  const root = createRoot(container)
  const editorDom = editor.view.dom
  const onSendNow = vi.fn<(text: string) => void>()
  const onQueueAdd = vi.fn<(text: string) => void>()
  act(() =>
    root.render(
      createElement(LocaleProvider, {
        initial: 'en',
        children: createElement(AiAskPopover, {
          editor,
          queueEnabled,
          queueFull: false,
          getItem: () => undefined,
          onSendNow,
          onQueueAdd,
          onQueueUpdate: () => {},
          onQueueRemove: () => {},
        }),
      }),
    ),
  )
  cleanup.push(() => {
    act(() => root.unmount())
    editor.destroy()
    container.remove()
    editorDom.remove()
  })
  const trigger = container.querySelector<HTMLButtonElement>('.ai-ask-trigger')!
  expect(trigger).not.toBeNull()
  act(() => {
    beforeOpen?.(editor)
    trigger.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }))
    trigger.click()
  })
  const input = container.querySelector<HTMLInputElement>('.ai-ask-pop-input')!
  const setText = (value: string) =>
    act(() => {
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!.call(input, value)
      input.dispatchEvent(new Event('input', { bubbles: true }))
    })
  return { editor, container, input, setText, onSendNow, onQueueAdd }
}

describe('Ask AI immediate instructions', () => {
  it.each([false, true])(
    'commits the live DOM text range before the prompt takes focus (queue=%s)',
    (queue) => {
      let revisionUnchanged = () => false
      const { editor, container, setText, onSendNow, onQueueAdd } = fixture(queue, (editor) => {
        document.body.append(editor.view.dom)
        editor.commands.setTextSelection({ from: 1, to: 2 })
        const text = editor.view.dom.querySelector('p')!.firstChild!
        // Native selection changes arrive before ProseMirror's next selectionchange flush.
        window.getSelection()!.setBaseAndExtent(text, 8, text, 0)
        expect(window.getSelection()!.toString()).toBe('Selected')
        expect(
          editor.state.doc.textBetween(editor.state.selection.from, editor.state.selection.to),
        ).toBe('S')
        const runtime = createDocsRuntime({
          current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
        })
        const client = runtime.connect({ id: 'selection-transfer' })
        const revision = client.context().revision
        revisionUnchanged = () => client.context().revision === revision
        cleanup.push(() => runtime.dispose())
      })
      expect(container.querySelector('.ai-ask-pop-sub')!.textContent).toBe('Selected')
      expect(editor.view.dom.querySelector('.doc-inactive-selection')!.textContent).toBe('Selected')
      expect(editor.state.selection.anchor).toBe(9)
      expect(editor.state.selection.head).toBe(1)
      const content = editor.getJSON()
      let captured = ''
      const submit = queue ? onQueueAdd : onSendNow
      submit.mockImplementation(() => {
        captured = editor.state.doc.textBetween(
          editor.state.selection.from,
          editor.state.selection.to,
        )
      })
      setText('Edit the selected passage')
      act(() => container.querySelector<HTMLButtonElement>('.ai-ask-confirm')!.click())
      expect(captured).toBe('Selected')
      expect(submit).toHaveBeenCalledExactlyOnceWith('Edit the selected passage')
      expect(editor.getJSON()).toEqual(content)
      expect(revisionUnchanged()).toBe(true)
      expect(editor.commands.undo()).toBe(false)
    },
  )

  it('ignores a DOM range outside the editor', () => {
    const { editor, container } = fixture(false, (editor) => {
      document.body.append(editor.view.dom)
      const outside = document.createElement('span')
      outside.textContent = 'Outside'
      document.body.append(outside)
      cleanup.push(() => outside.remove())
      window.getSelection()!.setBaseAndExtent(outside.firstChild!, 0, outside.firstChild!, 7)
    })
    expect(container.querySelector('.ai-ask-pop-sub')!.textContent).toBe('Selected')
    expect(editor.state.selection.from).toBe(1)
    expect(editor.state.selection.to).toBe(9)
  })

  it.each(['node', 'cell'] as const)(
    'preserves a native %s selection when the DOM text range differs',
    (kind) => {
      let original!: Selection
      const { editor } = fixture(false, (editor) => {
        editor.commands.setContent(
          '<table><tr><td><p>First</p></td><td><p>Second</p></td></tr></table>',
        )
        const table = editor.state.doc.firstChild!
        const selection =
          kind === 'node'
            ? NodeSelection.create(editor.state.doc, 0)
            : CellSelection.create(editor.state.doc, 1 + TableMap.get(table).map[0])
        editor.view.dispatch(editor.state.tr.setSelection(selection))
        original = selection
        document.body.append(editor.view.dom)
        const text = editor.view.dom.querySelector('p')!.firstChild!
        window.getSelection()!.setBaseAndExtent(text, 0, text, 2)
      })
      expect(editor.state.selection.eq(original)).toBe(true)
    },
  )
  it.each(['enter', 'confirm', 'local-send'] as const)(
    'preserves leading and trailing spaces with %s',
    (entry) => {
      const { container, input, setText, onSendNow, onQueueAdd } = fixture(entry === 'local-send')
      const instruction = '  Keep these exact words.  '
      setText(instruction)
      act(() => {
        if (entry === 'enter')
          input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
        else
          container
            .querySelector<HTMLButtonElement>(
              entry === 'confirm' ? '.ai-ask-confirm' : '.ai-ask-cancel',
            )!
            .click()
      })
      expect(onSendNow).toHaveBeenCalledExactlyOnceWith(instruction)
      expect(onQueueAdd).not.toHaveBeenCalled()
    },
  )

  it('continues to reject blank instructions', () => {
    const { container, input, setText, onSendNow, onQueueAdd } = fixture()
    setText('   ')
    expect(container.querySelector<HTMLButtonElement>('.ai-ask-confirm')!.disabled).toBe(true)
    act(() => input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true })))
    expect(onSendNow).not.toHaveBeenCalled()
    expect(onQueueAdd).not.toHaveBeenCalled()
  })
})
