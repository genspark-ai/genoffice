import { afterEach, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { Ribbon } from '../src/renderer/components/Ribbon'
import { computeFormatState } from '../src/renderer/components/ribbon-format-state'
import { LocaleProvider } from '../src/renderer/i18n/locale'
import { MENU_COMMAND_HANDLERS, type CommandHost } from '../src/renderer/commands/menu-commands'
import { printDoc, type FileActionContext } from '../src/renderer/file-actions'
import { ribbonProps } from './helpers/ribbon-props'

afterEach(() => vi.unstubAllGlobals())

it('offers Print in the non-mac File menu and opens the existing preview without printing', () => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
  vi.stubGlobal('desktop', { onLanguageChanged: () => () => {}, fontAliases: async () => ({}) })
  vi.stubGlobal(
    'ResizeObserver',
    class {
      observe() {}
      disconnect() {}
      unobserve() {}
    },
  )
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: '<p>Synthetic print-entry regression.</p>',
  })
  const container = document.createElement('div')
  document.body.append(container)
  const root = createRoot(container)
  const preview = vi.fn()
  const dialog = vi.fn()
  const ctx = {
    doc: { fileName: 'synthetic.docx' },
    printAutoOpenedPreviewRef: { current: false },
    setShowPagePreview: preview,
    setShowPrintDialog: dialog,
  } as unknown as FileActionContext
  const command = vi.fn((id: string) => {
    expect(id).toBe('print')
    MENU_COMMAND_HANDLERS.print({ hasDoc: true, printDoc: () => printDoc(ctx) } as CommandHost)
  })
  const render = (hasDoc: boolean) =>
    act(() =>
      root.render(
        createElement(LocaleProvider, {
          initial: 'en',
          children: createElement(Ribbon, {
            ...ribbonProps(editor, computeFormatState(editor)),
            hasDoc,
            onCommand: command,
          }),
        }),
      ),
    )
  const openFile = () =>
    act(() => container.querySelector<HTMLButtonElement>('.ribbon-tab-file')!.click())
  const printButton = () =>
    [...container.querySelectorAll<HTMLButtonElement>('.file-menu button')].find((button) =>
      button.textContent?.startsWith('Print'),
    )
  try {
    render(true)
    openFile()
    expect(printButton()).toBeDefined()
    expect(printButton()!.textContent).toContain('Ctrl+P')
    expect(printButton()!.disabled).toBe(false)
    act(() => printButton()!.click())
    expect(command).toHaveBeenCalledTimes(1)
    expect(preview).toHaveBeenCalledWith(true)
    expect(dialog).toHaveBeenCalledWith(true)
    expect(container.querySelector('.file-menu')).toBeNull()
    render(false)
    openFile()
    expect(printButton()!.disabled).toBe(true)
    act(() => printButton()!.click())
    expect(command).toHaveBeenCalledTimes(1)
  } finally {
    act(() => root.unmount())
    editor.destroy()
    container.remove()
  }
})
