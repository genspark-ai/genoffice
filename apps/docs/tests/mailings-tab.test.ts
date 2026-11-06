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
import { t } from '../src/renderer/i18n/locale'
import { ribbonProps } from './helpers/ribbon-props'

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
  sheet: '',
  fields: ['FirstName', 'LastName'],
  records: [
    { FirstName: 'Ada', LastName: 'Lovelace' },
    { FirstName: 'Grace', LastName: 'Hopper' },
  ],
}

/** drives the controlled select/input the way the browser does */
function setValue(el: HTMLSelectElement | HTMLInputElement, value: string) {
  const proto =
    el instanceof HTMLSelectElement ? HTMLSelectElement.prototype : HTMLInputElement.prototype
  Object.getOwnPropertyDescriptor(proto, 'value')!.set!.call(el, value)
  el.dispatchEvent(
    new Event(el instanceof HTMLSelectElement ? 'change' : 'input', { bubbles: true }),
  )
}

describe('the Mailings tab', () => {
  let editor: Editor
  let root: Root
  let container: HTMLElement
  const onInsertField = vi.fn()

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
  })

  function renderMailings(overrides: Record<string, unknown> = {}) {
    act(() =>
      root.render(
        createElement(Ribbon, {
          ...ribbonProps(editor, computeFormatState(editor)),
          mergeSource: source,
          onInsertField,
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
    const labels = items.map((b) => b.textContent)
    expect(labels).toContain(t('ribbonMailingsInsertNext'))
    expect(labels.filter((l) => l !== t('ribbonMailingsInsertNext'))).toEqual([
      'FirstName',
      'LastName',
    ])
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
    expect(onInsertField).toHaveBeenCalledWith(mergeFieldInstruction('FirstName'))
  })

  it('inserts a NEXT rule when the user clicks the next-record button', () => {
    const view = renderMailings()
    const fieldButton = [...view.querySelectorAll<HTMLButtonElement>('button.rb-big')].find(
      (b) => b.textContent === t('ribbonMailingsInsertField'),
    )!
    act(() => fieldButton.click())
    const nextItem = [...view.querySelectorAll<HTMLButtonElement>('button.rb-field-item')].find(
      (b) => b.textContent === t('ribbonMailingsInsertNext'),
    )!
    act(() => nextItem.click())
    expect(onInsertField).toHaveBeenCalledWith('NEXT')
  })

  it('has no field list without a data source', () => {
    const view = renderMailings({ mergeSource: null })
    const fieldButton = [...view.querySelectorAll<HTMLButtonElement>('button.rb-big')].find(
      (b) => b.textContent === t('ribbonMailingsInsertField'),
    )!
    expect(fieldButton.disabled).toBe(true)
  })

  it('shows the record the preview walks', () => {
    const view = renderMailings()
    expect(view.textContent).toContain(t('ribbonMailingsRecord', { n: 1, total: 2 }))
  })

  it('opens the recipient list beside the start button', () => {
    const view = renderMailings()
    const listButton = [...view.querySelectorAll<HTMLButtonElement>('button.rb-big')].find(
      (b) => b.textContent === t('ribbonMailingsRecipientList'),
    )!
    expect(listButton).toBeTruthy()
    act(() => listButton.click())
    const table = view.querySelector('.mailmerge-recipient-table')
    expect(table).toBeTruthy()
    expect(view.querySelectorAll('tbody input[type="checkbox"]').length).toBe(2)
  })

  it('takes a record out of the merge when the user unchecks it', () => {
    const view = renderMailings()
    const listButton = [...view.querySelectorAll<HTMLButtonElement>('button.rb-big')].find(
      (b) => b.textContent === t('ribbonMailingsRecipientList'),
    )!
    act(() => listButton.click())
    const boxes = view.querySelectorAll<HTMLInputElement>('tbody input[type="checkbox"]')
    act(() => boxes[0].click())
    expect(view.textContent).toContain(t('ribbonMailingsRecord', { n: 1, total: 1 }))
  })

  it('puts every record back when the user selects all', () => {
    const view = renderMailings()
    const listButton = [...view.querySelectorAll<HTMLButtonElement>('button.rb-big')].find(
      (b) => b.textContent === t('ribbonMailingsRecipientList'),
    )!
    act(() => listButton.click())
    const boxes = view.querySelectorAll<HTMLInputElement>('tbody input[type="checkbox"]')
    act(() => boxes[0].click())
    const selectAll = view.querySelector<HTMLInputElement>('thead input[type="checkbox"]')!
    act(() => selectAll.click())
    expect(view.textContent).toContain(t('ribbonMailingsRecord', { n: 1, total: 2 }))
  })

  it('filters the records shown in the list', () => {
    const view = renderMailings()
    const listButton = [...view.querySelectorAll<HTMLButtonElement>('button.rb-big')].find(
      (b) => b.textContent === t('ribbonMailingsRecipientList'),
    )!
    act(() => listButton.click())
    const filter = [...view.querySelectorAll<HTMLButtonElement>('.mailmerge-toolbar button')].find(
      (b) => b.textContent === t('ribbonMailingsFilter'),
    )!
    act(() => filter.click())
    act(() =>
      setValue(view.querySelector<HTMLSelectElement>('.mailmerge-cond-field')!, 'FirstName'),
    )
    act(() => setValue(view.querySelector<HTMLInputElement>('.mailmerge-cond-value')!, 'Ada'))
    act(() => view.querySelector<HTMLButtonElement>('.mailmerge-cond-footer .primary')!.click())
    const firstNames = [...view.querySelectorAll('tbody tr td:nth-child(2)')].map(
      (td) => td.textContent,
    )
    expect(firstNames).toEqual(['Ada'])
    expect(view.querySelectorAll('tbody input[type="checkbox"]').length).toBe(1)
  })

  it('says so when nothing matches the filter', () => {
    const view = renderMailings()
    const listButton = [...view.querySelectorAll<HTMLButtonElement>('button.rb-big')].find(
      (b) => b.textContent === t('ribbonMailingsRecipientList'),
    )!
    act(() => listButton.click())
    const filter = [...view.querySelectorAll<HTMLButtonElement>('.mailmerge-toolbar button')].find(
      (b) => b.textContent === t('ribbonMailingsFilter'),
    )!
    act(() => filter.click())
    act(() =>
      setValue(view.querySelector<HTMLSelectElement>('.mailmerge-cond-field')!, 'FirstName'),
    )
    act(() => setValue(view.querySelector<HTMLInputElement>('.mailmerge-cond-value')!, 'zzz'))
    act(() => view.querySelector<HTMLButtonElement>('.mailmerge-cond-footer .primary')!.click())
    expect(view.textContent).toContain(t('ribbonMailingsNoFilterResults'))
  })

  it('sorts the records by the chosen field', () => {
    const view = renderMailings()
    const listButton = [...view.querySelectorAll<HTMLButtonElement>('button.rb-big')].find(
      (b) => b.textContent === t('ribbonMailingsRecipientList'),
    )!
    act(() => listButton.click())
    const sort = [...view.querySelectorAll<HTMLButtonElement>('.mailmerge-toolbar button')].find(
      (b) => b.textContent === t('ribbonMailingsSort'),
    )!
    act(() => sort.click())
    const [field] = view.querySelectorAll<HTMLSelectElement>('.mailmerge-cond-modal select')
    act(() => setValue(field, 'LastName'))
    act(() => view.querySelector<HTMLButtonElement>('.mailmerge-cond-footer .primary')!.click())
    const cells = [
      ...view.querySelectorAll('.mailmerge-recipient-table tbody tr td:nth-child(3)'),
    ].map((td) => td.textContent)
    expect(cells).toEqual(['Hopper', 'Lovelace'])
  })

  it('asks for the file and the sheet when the merge starts', async () => {
    const onPickFile = vi.fn(async () => ({
      path: '/tmp/recipients.xlsx',
      name: 'recipients.xlsx',
      sheets: ['Customers', 'Orders'],
    }))
    const onReadSource = vi.fn(async () => true)
    const view = renderMailings({
      onPickMergeDataSourceFile: onPickFile,
      onReadMergeDataSource: onReadSource,
    })
    const startButton = [...view.querySelectorAll<HTMLButtonElement>('button.rb-big')].find(
      (b) => b.textContent === t('ribbonMailingsStart'),
    )!
    await act(async () => {
      startButton.click()
    })
    // the dialog names the sheets of the picked file
    const radios = view.querySelectorAll('input[type="radio"]')
    expect(radios.length).toBe(2)
    expect(view.textContent).toContain('Customers')
    expect(view.textContent).toContain('Orders')
  })

  it('reads the sheet the user confirmed', async () => {
    const onPickFile = vi.fn(async () => ({
      path: '/tmp/recipients.xlsx',
      name: 'recipients.xlsx',
      sheets: ['Customers', 'Orders'],
    }))
    const onReadSource = vi.fn(async () => true)
    const view = renderMailings({
      onPickMergeDataSourceFile: onPickFile,
      onReadMergeDataSource: onReadSource,
    })
    const startButton = [...view.querySelectorAll<HTMLButtonElement>('button.rb-big')].find(
      (b) => b.textContent === t('ribbonMailingsStart'),
    )!
    await act(async () => {
      startButton.click()
    })
    const second = view.querySelectorAll<HTMLInputElement>('input[type="radio"]')[1]!
    act(() => second.click())
    const confirm = [...view.querySelectorAll<HTMLButtonElement>('button')].find(
      (b) => b.textContent === t('ribbonMailingsConfirmSource'),
    )!
    await act(async () => {
      confirm.click()
    })
    expect(onReadSource).toHaveBeenCalledWith('/tmp/recipients.xlsx', 'Orders')
  })
})
