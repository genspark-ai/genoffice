/**
 * Insert > Chart's type list must read as a list: one button per chart type,
 * no two types sharing a label in any language (genoffice#2910 — fr showed
 * "Histogramme" for both the Column chart and the frequency histogram), and
 * the buttons row wraps inside the 460px modal instead of overflowing it.
 */
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { ChartInsertModal } from '../src/renderer/components/ribbon-insert-tab'
import { CHART_TYPES, CHART_TYPE_LABELS } from '../src/renderer/components/chart-types'
import { ribbonStrings } from '../src/renderer/i18n/strings-ribbon'
import { editorExtensions } from '../src/renderer/editor/extensions'

describe('chart dialog type labels', () => {
  it('names every chart type distinctly in every language', () => {
    for (const [lang, strings] of Object.entries(ribbonStrings)) {
      const dict = strings as Record<string, string>
      const labels = CHART_TYPES.map((type) => dict[CHART_TYPE_LABELS[type]])
      expect(labels, `${lang} has a duplicate chart label`).toHaveLength(new Set(labels).size)
      expect(labels.every((label) => typeof label === 'string' && label.length > 0)).toBe(true)
    }
  })
})

describe('chart dialog type row', () => {
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

  it('renders one button per type on the wrapping row', () => {
    act(() => root.render(createElement(ChartInsertModal, { editor, onClose: () => {} })))
    // the CSS wrap rule hangs on this class; its effect itself is e2e territory
    const row = container.querySelector('.modal-chart .modal-row.modal-chart-types')
    expect(row).not.toBeNull()
    expect(row!.querySelectorAll('button')).toHaveLength(CHART_TYPES.length)
  })
})
