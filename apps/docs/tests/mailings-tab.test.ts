// The Mailings tab: the field list comes from the data source, clicking a
// field inserts a real MERGEFIELD, and the panel walks the records.
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { mergeFieldInstruction } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { computeFormatState } from '../src/renderer/components/ribbon-format-state'
import { Ribbon } from '../src/renderer/components/Ribbon'
import { MailMergePanel } from '../src/renderer/components/MailMergePanel'
import { t } from '../src/renderer/i18n/locale'
import { ribbonProps } from './helpers/ribbon-props'
import { MERGE_RANGE_ALL, type MergeRange } from '../src/renderer/mail-merge/range'

function makeEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'hello' }] }],
    },
  })
}

const source = {
  path: '/tmp/recipients.csv',
  name: 'recipients.csv',
  fields: ['FirstName', 'LastName'],
  records: [
    { FirstName: 'Ada', LastName: 'Lovelace' },
    { FirstName: 'Grace', LastName: 'Hopper' },
  ],
}

describe('the Mailings tab', () => {
  let editor: Editor
  let root: Root
  let container: HTMLElement
  const onInsertField = vi.fn()
  const onCommand = vi.fn()

  beforeEach(() => {
    editor = makeEditor()
    container = document.createElement('div')
    document.body.appendChild(container)
    root = createRoot(container)
  })

  afterEach(() => {
    act(() => root.unmount())
    container.remove()
    editor.destroy()
    onInsertField.mockClear()
    onCommand.mockClear()
  })

  function renderMailings(overrides: Record<string, unknown> = {}) {
    act(() =>
      root.render(
        createElement(Ribbon, {
          ...ribbonProps(editor, computeFormatState(editor)),
          mergeSource: source,
          onCommand,
          ...overrides,
        }),
      ),
    )
    const tab = [...container.querySelectorAll<HTMLButtonElement>('.ribbon-tab')].find(
      (b) => b.textContent === t('ribbonTabMailings'),
    )!
    act(() => tab.click())
    return container
  }

  it('offers the fields of the data source', () => {
    const view = renderMailings()
    const fieldButton = [...view.querySelectorAll<HTMLButtonElement>('button.rb-big')].find(
      (b) => b.textContent === t('ribbonMailingsInsertField'),
    )!
    expect(fieldButton.disabled).toBe(false)
    act(() => fieldButton.click())
    const items = [...view.querySelectorAll<HTMLButtonElement>('button.rb-field-item')]
    expect(items.map((b) => b.textContent)).toEqual(['FirstName', 'LastName'])
  })

  it('inserts the merge field of the field the user clicked', () => {
    const view = renderMailings()
    const fieldButton = [...view.querySelectorAll<HTMLButtonElement>('button.rb-big')].find(
      (b) => b.textContent === t('ribbonMailingsInsertField'),
    )!
    act(() => fieldButton.click())
    const item = [...view.querySelectorAll<HTMLButtonElement>('button.rb-field-item')].find(
      (b) => b.textContent === 'FirstName',
    )!
    act(() => item.click())
    expect(onCommand).toHaveBeenCalledWith('insert-field', mergeFieldInstruction('FirstName'))
  })

  it('has no field list without a data source', () => {
    const view = renderMailings({ mergeSource: null })
    const fieldButton = [...view.querySelectorAll<HTMLButtonElement>('button.rb-big')].find(
      (b) => b.textContent === t('ribbonMailingsInsertField'),
    )!
    expect(fieldButton.disabled).toBe(true)
  })

  it('shows the record the preview walks', () => {
    const view = renderMailings({ mergeRecordIndex: 1 })
    expect(view.textContent).toContain(t('ribbonMailingsRecord', { n: 2, total: 2 }))
  })

  const field = (view: HTMLElement, label: string) =>
    view.querySelector<HTMLInputElement>(`input[aria-label="${label}"]`)!
  /** types into a box and leaves it, as Tab or Enter does */
  const commit = (input: HTMLInputElement, text: string) =>
    act(() => {
      input.focus()
      input.value = text
      input.blur()
    })

  it('the Record box shows the previewed record and moves the preview to the number typed', () => {
    const onMergeRecordIndex = vi.fn()
    const view = renderMailings({ mergeRecordIndex: 1, onMergeRecordIndex })
    const record = field(view, t('ribbonMailingsRecordBox'))
    expect(record.value).toBe('2')
    commit(record, '1')
    expect(onMergeRecordIndex).toHaveBeenLastCalledWith(0)
    // past the end stops at the last record, which is the one shown already
    onMergeRecordIndex.mockClear()
    commit(record, '9')
    expect(record.value).toBe('2')
    expect(onMergeRecordIndex).not.toHaveBeenCalled()
    expect(
      field(renderMailings({ mergeSource: null }), t('ribbonMailingsRecordBox')).disabled,
    ).toBe(true)
  })

  it('Merge Range picks All, Current Record or Custom; From and To take numbers only for Custom', () => {
    const onMergeRange = vi.fn()
    const view = renderMailings({ onMergeRange })
    const from = () => field(view, t('ribbonMergeRangeFrom'))
    const to = () => field(view, t('ribbonMergeRangeTo'))
    expect(from().disabled).toBe(true)
    const range = view.querySelector<HTMLButtonElement>(
      `button[aria-label="${t('ribbonMergeRange')}"]`,
    )!
    expect(range.textContent).toBe(t('ribbonMergeRangeAll'))
    act(() => range.click())
    const items = [...view.querySelectorAll<HTMLButtonElement>('.layout-menu button')]
    expect(items.map((b) => b.textContent)).toEqual([
      `${t('ribbonMergeRangeAll')} ✓`,
      t('ribbonMergeRangeCurrent'),
      t('ribbonMergeRangeCustom'),
    ])
    act(() => items[2].click())
    expect(onMergeRange).toHaveBeenLastCalledWith({ kind: 'custom', from: 1, to: null })

    const custom: MergeRange = { kind: 'custom', from: 1, to: null }
    renderMailings({ onMergeRange, mergeRange: custom })
    expect([from().disabled, from().value, to().value]).toEqual([false, '1', '2'])
    commit(from(), '2')
    expect(onMergeRange).toHaveBeenLastCalledWith({ kind: 'custom', from: 2, to: null })
    commit(to(), '1')
    expect(onMergeRange).toHaveBeenLastCalledWith({ kind: 'custom', from: 1, to: 1 })
  })

  it('Update Labels is greyed out while the document is not a sheet of labels', () => {
    const view = renderMailings()
    const update = [...view.querySelectorAll<HTMLButtonElement>('button.rb-big')].find(
      (b) => b.textContent === t('ribbonUpdateLabels'),
    )!
    expect(update.disabled).toBe(true)
  })

  it('Finish & Merge takes the records of the merge range', async () => {
    const onMerge = vi.fn(async () => ({ ok: true }))
    const mergedWith = async (range: MergeRange, recordIndex = 0) => {
      onMerge.mockClear()
      act(() =>
        root.render(
          createElement(MailMergePanel, {
            editor,
            source,
            onPickSource: () => {},
            onInsertField,
            onMerge,
            recordIndex,
            onRecordIndex: () => {},
            range,
            onClose: () => {},
          }),
        ),
      )
      const finish = container.querySelector<HTMLButtonElement>('button.btn-primary')!
      if (finish.disabled) return null
      await act(async () => finish.click())
      return (onMerge.mock.calls[0] as unknown as [{ records: unknown[] }])[0].records
    }
    expect(await mergedWith(MERGE_RANGE_ALL)).toEqual(source.records)
    expect(await mergedWith({ kind: 'current', from: 1, to: null }, 1)).toEqual([source.records[1]])
    expect(await mergedWith({ kind: 'custom', from: 1, to: 1 })).toEqual([source.records[0]])
    expect(await mergedWith({ kind: 'custom', from: 2, to: 1 })).toBeNull()
  })

  it('the merge panel opens on the kind of output asked for, and follows a later pick while open', () => {
    const radios = () => [
      ...container.querySelectorAll<HTMLInputElement>('input[name="mailmerge-output"]'),
    ]
    const checkedOutput = (initialMode?: 'envelopes' | 'labels', pick = 1) => {
      act(() =>
        root.render(
          createElement(MailMergePanel, {
            editor,
            source,
            onPickSource: () => {},
            onInsertField,
            onMerge: async () => ({ ok: true }),
            recordIndex: 0,
            onRecordIndex: () => {},
            range: MERGE_RANGE_ALL,
            initialMode,
            pick,
            onClose: () => {},
          }),
        ),
      )
      return radios().findIndex((radio) => radio.checked)
    }
    expect(checkedOutput()).toBe(0)
    expect(checkedOutput('labels')).toBe(2)
    expect(checkedOutput('envelopes')).toBe(3)
    // the user picks another output, then asks for Envelopes again
    act(() => radios()[1].click())
    expect(radios().findIndex((radio) => radio.checked)).toBe(1)
    expect(checkedOutput('envelopes', 2)).toBe(3)
  })

  it('Start Mail Merge, Envelopes and Labels run their commands', () => {
    const view = renderMailings()
    const big = (label: string) =>
      [...view.querySelectorAll<HTMLButtonElement>('button.rb-big')].find(
        (b) => b.textContent === label,
      )!
    act(() => big(t('ribbonMailingsStart')).click())
    act(() => big(t('ribbonEnvelopes')).click())
    act(() => big(t('ribbonLabels')).click())
    expect(onCommand.mock.calls.map(([id]) => id)).toEqual([
      'mailings-start',
      'tools-envelopes',
      'tools-labels',
    ])
  })
})
