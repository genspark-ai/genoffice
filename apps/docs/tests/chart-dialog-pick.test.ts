/**
 * Insert > Chart lists the chart types. A pick made while the chart dialog is
 * already open moves the dialog to that type, also when it repeats a type the
 * user has since changed inside the dialog.
 */
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { ChartInsertModal } from '../src/renderer/components/ribbon-insert-tab'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { t } from '../src/renderer/i18n/locale'

describe('chart dialog type pick', () => {
  let editor: Editor
  let root: Root
  let container: HTMLElement

  beforeEach(() => {
    editor = new Editor({ element: document.createElement('div'), extensions: editorExtensions })
    container = document.createElement('div')
    document.body.appendChild(container)
    root = createRoot(container)
  })

  afterEach(() => {
    act(() => root.unmount())
    container.remove()
    editor.destroy()
  })

  const show = (initialType: 'bar' | 'line' | 'pie', pick: number) =>
    act(() =>
      root.render(
        createElement(ChartInsertModal, { editor, initialType, pick, onClose: () => {} }),
      ),
    )
  const selected = () => container.querySelector('.modal-chart .btn-primary')?.textContent
  const typeButton = (label: string) =>
    [...container.querySelectorAll<HTMLButtonElement>('.modal-chart .modal-row button')].find(
      (b) => b.textContent === label,
    )!

  it('follows a new pick and a repeated pick, and keeps the user’s own choice in between', () => {
    show('pie', 1)
    expect(selected()).toBe(t('ribbonChartPie'))
    show('line', 2)
    expect(selected()).toBe(t('ribbonChartLine'))
    act(() => typeButton(t('ribbonChartBar')).click())
    expect(selected()).toBe(t('ribbonChartBar'))
    // the same type again, picked afresh in the menu
    show('line', 3)
    expect(selected()).toBe(t('ribbonChartLine'))
  })
})
