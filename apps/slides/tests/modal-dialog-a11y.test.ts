/**
 * Insert / chart modals share PrintDialog's contract: role=dialog, aria-modal,
 * heading label, Escape closes, first control focused, opener refocused.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, createElement, type ReactElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'

import { ChartTypeDialog } from '../src/renderer/components/ChartTypeDialog'
import { TableInsertDialog } from '../src/renderer/components/InsertDialogs'

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

const roots: Array<{ root: Root; container: HTMLElement }> = []
afterEach(() => {
  for (const { root, container } of roots.splice(0)) {
    act(() => root.unmount())
    container.remove()
  }
  document.body.innerHTML = ''
})

async function mount(element: ReactElement) {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  await act(async () => {
    root.render(element)
  })
  roots.push({ root, container })
  return { container, root }
}

function expectDialogSemantics(container: HTMLElement) {
  const dlg = container.querySelector<HTMLElement>('.modal')!
  expect(dlg.getAttribute('role')).toBe('dialog')
  expect(dlg.getAttribute('aria-modal')).toBe('true')
  const titleId = dlg.getAttribute('aria-labelledby')
  expect(titleId).toBeTruthy()
  expect(dlg.querySelector('h2')?.id).toBe(titleId)
  expect(dlg.contains(document.activeElement)).toBe(true)
  return dlg
}

describe('slides modal semantics', () => {
  it('ChartTypeDialog is a labelled modal that closes on Escape', async () => {
    const onClose = vi.fn()
    const { container } = await mount(
      createElement(ChartTypeDialog, { current: 'bar', onConfirm: () => {}, onClose }),
    )
    expectDialogSemantics(container)
    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }))
    })
    expect(onClose).toHaveBeenCalledTimes(1)
  })

  it('TableInsertDialog focuses a control on open and hands focus back on close', async () => {
    const opener = document.createElement('button')
    document.body.appendChild(opener)
    opener.focus()
    const { container, root } = await mount(
      createElement(TableInsertDialog, { onInsert: () => {}, onClose: () => {} }),
    )
    expectDialogSemantics(container)
    await act(async () => root.unmount())
    roots.pop()
    container.remove()
    expect(document.activeElement).toBe(opener)
  })
})
