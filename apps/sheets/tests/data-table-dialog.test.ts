// @vitest-environment jsdom
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, describe, expect, it } from 'vitest'

import { DataTableDialog } from '../src/renderer/DataTableDialog'
import type { DataTableInput, DataTableOutcome } from '../src/renderer/data-table'

let root: Root | null = null
let container: HTMLDivElement | null = null

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
})

afterEach(async () => {
  if (root) await act(async () => root?.unmount())
  container?.remove()
  root = null
  container = null
  document.body.innerHTML = ''
})

async function mount(element: React.ReactElement): Promise<HTMLDivElement> {
  container = document.createElement('div')
  document.body.appendChild(container)
  root = createRoot(container)
  await act(async () => {
    root!.render(element)
    await Promise.resolve()
  })
  return container
}

function type(input: HTMLInputElement, text: string): void {
  act(() => {
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set
    setter?.call(input, text)
    input.dispatchEvent(new Event('input', { bubbles: true }))
  })
}

function click(element: Element): void {
  act(() => element.dispatchEvent(new MouseEvent('click', { bubbles: true })))
}

type Apply = (input: DataTableInput) => Promise<DataTableOutcome>

describe('Data Table dialog validation', () => {
  it('disables OK and hints until an input cell is filled', async () => {
    const applied: DataTableInput[] = []
    const host = await mount(
      createElement(DataTableDialog, {
        initialRange: 'B4:C7',
        onApply: ((input) => {
          applied.push(input)
          return Promise.resolve({ kind: 'column', cells: 1 })
        }) satisfies Apply,
        onClose: () => undefined,
      }),
    )
    const inputs = host.querySelectorAll<HTMLInputElement>('input')
    const ok = host.querySelector<HTMLButtonElement>('button.primary-action')!
    expect(inputs[0]!.value).toBe('B4:C7')
    expect(ok.disabled).toBe(true)
    expect(host.textContent).toContain('请至少填写一个输入单元格')

    type(inputs[2]!, 'D1')
    expect(ok.disabled).toBe(false)
    expect(host.textContent).not.toContain('请至少填写一个输入单元格')
    expect(applied).toHaveLength(0)
  })

  it('keeps OK disabled for a malformed range or input reference', async () => {
    const host = await mount(
      createElement(DataTableDialog, {
        initialRange: '',
        onApply: (() => Promise.resolve({ kind: 'column', cells: 1 })) satisfies Apply,
        onClose: () => undefined,
      }),
    )
    const inputs = host.querySelectorAll<HTMLInputElement>('input')
    const ok = host.querySelector<HTMLButtonElement>('button.primary-action')!
    type(inputs[0]!, 'B4-C7')
    type(inputs[2]!, 'D1')
    expect(ok.disabled).toBe(true)
    type(inputs[0]!, 'B4:C7')
    expect(ok.disabled).toBe(false)
    type(inputs[2]!, 'D1:D8')
    expect(ok.disabled).toBe(true)
  })

  it('normalizes the input, shows progress, then the outcome', async () => {
    const applied: DataTableInput[] = []
    let release: (outcome: DataTableOutcome) => void = () => undefined
    const apply: Apply = (input) => {
      applied.push(input)
      return new Promise((resolve) => (release = resolve))
    }
    const host = await mount(
      createElement(DataTableDialog, {
        initialRange: 'b4:c7',
        onApply: apply,
        onClose: () => undefined,
      }),
    )
    type(host.querySelectorAll<HTMLInputElement>('input')[2]!, 'd1')
    const ok = host.querySelector<HTMLButtonElement>('button.primary-action')!
    click(ok)

    expect(applied).toHaveLength(1)
    const argument = applied[0]!
    expect(argument.tableRange).toBe('B4:C7')
    expect(argument.columnInputCell).toBe('D1')
    expect(argument.rowInputCell).toBeNull()
    expect(typeof argument.onProgress).toBe('function')

    act(() => argument.onProgress?.(1, 3))
    expect(host.textContent).toContain('运算中… 1/3')

    await act(async () => {
      release({ kind: 'column', cells: 3 })
      await Promise.resolve()
      await Promise.resolve()
    })
    expect(host.textContent).toContain('已写入 3 个结果')
    expect(ok.disabled).toBe(false)
  })

  it('surfaces a rejected run as an alert and re-enables OK', async () => {
    const host = await mount(
      createElement(DataTableDialog, {
        initialRange: 'B4:C7',
        onApply: (() => Promise.reject(new Error('区域形状不符'))) satisfies Apply,
        onClose: () => undefined,
      }),
    )
    type(host.querySelectorAll<HTMLInputElement>('input')[2]!, 'D1')
    const ok = host.querySelector<HTMLButtonElement>('button.primary-action')!
    click(ok)
    await act(async () => {
      await Promise.resolve()
      await Promise.resolve()
    })
    const alert = host.querySelector('[role="alert"]')!
    expect(alert.textContent).toContain('区域形状不符')
    expect(ok.disabled).toBe(false)
  })
})
