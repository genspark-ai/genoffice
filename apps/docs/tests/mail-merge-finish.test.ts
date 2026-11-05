// The mail merge output: the panel picks the output shape, and the finish
// section asks which records to merge and opens the result in the app — the
// 0.12.29 behaviour, no save dialog.
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { MergeFinishDialog } from '../src/renderer/components/MergeFinishDialog'
import { MailMergePanel } from '../src/renderer/components/MailMergePanel'
import { t } from '../src/renderer/i18n/locale'
import type { MailMergeRequest } from '../src/renderer/mail-merge/output'

const source = {
  path: '/tmp/recipients.csv',
  name: 'recipients.csv',
  sheet: '',
  fields: ['FirstName'],
  records: [{ FirstName: 'Ada' }],
}

const records = [{ FirstName: 'Ada' }, { FirstName: 'Grace' }]

describe('the merge output', () => {
  let editor: Editor
  let root: Root
  let container: HTMLElement
  const onMerge = vi.fn(async (_request: MailMergeRequest) => ({ ok: true, count: 1 }))

  beforeEach(() => {
    editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'x' }] }],
      },
    })
    container = document.createElement('div')
    document.body.appendChild(container)
    root = createRoot(container)
    onMerge.mockClear()
  })

  afterEach(() => {
    act(() => root.unmount())
    container.remove()
    editor.destroy()
  })

  function button(label: string): HTMLButtonElement | null {
    return (
      [...container.querySelectorAll('button')].find((b) => b.textContent?.trim() === label) ?? null
    )
  }

  function radio(label: string): HTMLInputElement | null {
    const labels = [...container.querySelectorAll('label.mailmerge-radio')]
    const match = labels.find((l) => l.textContent?.trim() === label)
    return (match?.querySelector('input[type="radio"]') as HTMLInputElement) ?? null
  }

  // the first range radio lives inside the record-range block
  function pickAllRange() {
    const block = container.querySelector('.mailmerge-range-block')
    const first = block?.querySelector('input[type="radio"]') as HTMLInputElement | null
    act(() => first?.click())
  }

  describe('the finish section', () => {
    function renderFinish() {
      act(() =>
        root.render(
          createElement(MergeFinishDialog, {
            source: source as never,
            records,
            currentIndex: 0,
            editor,
            mode: 'single',
            namePattern: 'FirstName',
            onMerge,
          }),
        ),
      )
    }

    it('opens the merged document in the app when finished', async () => {
      renderFinish()
      pickAllRange()
      await act(async () => button(t('ribbonMailingsFinish'))?.click())
      expect(onMerge).toHaveBeenCalledTimes(1)
      expect(onMerge.mock.calls[0]![0]).toMatchObject({ mode: 'single' })
    })

    it('disables finish until a range is picked', () => {
      renderFinish()
      expect(button(t('ribbonMailingsFinish'))?.disabled).toBe(true)
      pickAllRange()
      expect(button(t('ribbonMailingsFinish'))?.disabled).toBe(false)
    })
  })

  describe('the panel', () => {
    function renderPanel() {
      act(() =>
        root.render(
          createElement(MailMergePanel, {
            editor,
            source: source as never,
            records,
            onPickSource: () => {},
            onMerge,
            recordIndex: 0,
            onClose: () => {},
          }),
        ),
      )
    }

    it('fuses shape and range under one output section', () => {
      renderPanel()
      const headings = [
        ...container.querySelectorAll('.mailmerge-section .mailmerge-label'),
      ].map((e) => e.textContent?.trim())
      expect(headings).toContain(t('ribbonMailingsOutput'))
      expect(headings).toContain(t('ribbonMailingsRange'))
      // the finish button lives in the same section as the shape radios
      expect(button(t('ribbonMailingsFinish'))).not.toBeNull()
    })

    it('lets the per-record output finish in the app (not disk-only)', async () => {
      renderPanel()
      const perRecord = radio(t('ribbonMailingsOutputPerRecord'))
      act(() => perRecord?.click())
      pickAllRange()
      expect(button(t('ribbonMailingsFinish'))?.disabled).toBe(false)
      await act(async () => button(t('ribbonMailingsFinish'))?.click())
      expect(onMerge).toHaveBeenCalledTimes(1)
      expect(onMerge.mock.calls[0]![0]).toMatchObject({ mode: 'perRecord' })
    })
  })
})
