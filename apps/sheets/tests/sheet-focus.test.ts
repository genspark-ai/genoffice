import { afterEach, describe, expect, it, vi } from 'vitest'
import { focusWorksheet, keepActiveSheet } from '../src/renderer/sheet-focus'

afterEach(() => vi.useRealTimers())

function fixture() {
  const summary = { getSheetId: () => 'summary' }
  const data = { getSheetId: () => 'data' }
  let active = summary
  const workbook = {
    getActiveSheet: () => active,
    setActiveSheet: (sheet: unknown) => {
      active = sheet as typeof summary
    },
  }
  const target = { ...data, getWorkbook: () => workbook }
  return { summary, data, workbook, target }
}

describe('background sheet restoration', () => {
  it('restores an unintended delayed activation by background loading', async () => {
    vi.useFakeTimers()
    const { summary, data, workbook, target } = fixture()
    keepActiveSheet(target, () => {
      queueMicrotask(() => workbook.setActiveSheet(data))
    })
    await vi.runAllTimersAsync()
    expect(workbook.getActiveSheet()).toBe(summary)
  })

  it('keeps new background loads from stealing focus after explicit navigation', async () => {
    vi.useFakeTimers()
    const { summary, data, workbook, target } = fixture()
    focusWorksheet(target, () => workbook.setActiveSheet(data))
    const background = { ...summary, getWorkbook: () => workbook }
    keepActiveSheet(background, () => {
      queueMicrotask(() => workbook.setActiveSheet(summary))
    })
    await vi.runAllTimersAsync()
    expect(workbook.getActiveSheet()).toBe(data)
  })

  it('does not undo an explicit focus on the sheet being loaded', async () => {
    vi.useFakeTimers()
    const { data, workbook, target } = fixture()
    keepActiveSheet(target, () => workbook.setActiveSheet(data))
    focusWorksheet(target, () => workbook.setActiveSheet(data))
    await vi.runAllTimersAsync()
    expect(workbook.getActiveSheet()).toBe(data)
  })
})
