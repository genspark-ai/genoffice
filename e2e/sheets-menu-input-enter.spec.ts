import { test, expect } from '@playwright/test'
import { mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, type LaunchedApp } from './helpers'

// the preload exposes window.__genofficeDebug only under this env var
process.env.GENOFFICE_DEBUG_HOOKS = '1'

interface DebugRange {
  setValues(v: unknown[][]): Promise<unknown>
  getValues(): unknown[][]
  getCellRect(): { x: number; y: number; width: number; height: number }
}

type DebugWindow = Window & {
  __genofficeDebug: {
    univerAPI: {
      getActiveWorkbook(): {
        getActiveSheet(): {
          getRange(row: number, column: number, rows: number, columns: number): DebugRange
          getFilter(): {
            getColumnFilterCriteria(column: number): unknown
            getFilteredOutRows(): number[]
          } | null
        }
      }
    }
  }
}

const HEADER = { width: 46, height: 24 }
const COL_WIDTH = 86

async function openBlankSheet(
  launched: LaunchedApp,
  scratch: string,
): Promise<{ sheets: Page; grid: { x: number; y: number } }> {
  const { app, page } = launched
  await app.evaluate(({ app: electronApp }, dir) => {
    electronApp.setPath('documents', dir)
  }, scratch)

  await expect(page.locator('.quick-card').nth(1)).toContainText('AI Sheets')
  await page.locator('.quick-card').nth(1).click()

  const sheets = await waitForPageWithUrl(app, '://sheets/')
  await sheets.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
    timeout: 30_000,
  })
  await sheets.waitForTimeout(1_500)

  const grid = await sheets.evaluate(() => {
    for (const canvas of document.querySelectorAll('canvas')) {
      const rect = canvas.getBoundingClientRect()
      if (rect.width > 500 && rect.height > 300) return { x: rect.x, y: rect.y }
    }
    return null
  })
  if (!grid) throw new Error('worksheet canvas not found')
  return { sheets, grid }
}

function setValues(sheets: Page, row: number, column: number, values: unknown[][]): Promise<void> {
  return sheets.evaluate(
    async ({ row, column, values }) => {
      await (window as unknown as DebugWindow).__genofficeDebug.univerAPI
        .getActiveWorkbook()
        .getActiveSheet()
        .getRange(row, column, values.length, values[0]?.length ?? 1)
        .setValues(values)
    },
    { row, column, values },
  )
}

async function cellCenter(
  sheets: Page,
  grid: { x: number; y: number },
  row: number,
  column: number,
): Promise<{ x: number; y: number }> {
  const rect = await sheets.evaluate(
    ({ row, column }) =>
      (window as unknown as DebugWindow).__genofficeDebug.univerAPI
        .getActiveWorkbook()
        .getActiveSheet()
        .getRange(row, column, 1, 1)
        .getCellRect(),
    { row, column },
  )
  return { x: grid.x + rect.x + rect.width / 2, y: grid.y + rect.y + rect.height / 2 }
}

function getValues(
  sheets: Page,
  row: number,
  column: number,
  rows: number,
  columns: number,
): Promise<unknown[][]> {
  return sheets.evaluate(
    ({ row, column, rows, columns }) =>
      (window as unknown as DebugWindow).__genofficeDebug.univerAPI
        .getActiveWorkbook()
        .getActiveSheet()
        .getRange(row, column, rows, columns)
        .getValues(),
    { row, column, rows, columns },
  )
}

/**
 * Regression for "typed a count into the right-click insert-N-columns box,
 * pressed Enter, nothing happened" (user report): Enter in a menu
 * count box must run the row's action like a click, not just commit the
 * number and leave the menu open.
 */
test.describe('sheets: Enter runs the context-menu insert-N action', () => {
  test('insert 3 columns left of B via the count box and Enter', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-menu-enter-'))
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-menu-input-enter',
    })
    try {
      const { sheets, grid } = await openBlankSheet(launched, scratch)
      await setValues(sheets, 0, 1, [['marker']])

      // right-click column B's header → context menu with the count boxes
      await sheets.mouse.click(
        grid.x + HEADER.width + COL_WIDTH * 1.5,
        grid.y + HEADER.height / 2,
        { button: 'right' },
      )
      const countBox = sheets.locator('input.univer-h-7').first()
      await expect(countBox).toBeVisible()

      await countBox.click()
      await sheets.keyboard.press('ControlOrMeta+a')
      await sheets.keyboard.type('3', { delay: 50 })
      await sheets.keyboard.press('Enter')

      // the insert ran: the marker moved from B1 to E1 and the menu closed
      await expect(async () => {
        const values = (await getValues(sheets, 0, 0, 1, 6))[0] ?? []
        expect(values[4]).toBe('marker')
        expect(values[1]).toBeNull()
      }).toPass({ timeout: 10_000 })
      await expect(countBox).toBeHidden()
    } finally {
      await closeAndSaveVideo(launched, 'sheets-menu-input-enter')
    }
  })
})

test.describe('sheets: cell context menu Excel entries', () => {
  test('Pick From Drop-down List writes the pick and Filter by value hides the rest', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-menu-pick-'))
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-menu-pick-from-list',
    })
    try {
      const { sheets, grid } = await openBlankSheet(launched, scratch)
      await setValues(sheets, 0, 1, [['alpha'], ['beta']])
      await setValues(sheets, 3, 1, [['gamma']])

      // right-click the empty B3 between the two blocks
      const b3 = await cellCenter(sheets, grid, 2, 1)
      await sheets.mouse.click(b3.x, b3.y, { button: 'right' })
      await expect(sheets.getByText('Number Format', { exact: true })).toBeVisible()
      await expect(sheets.getByText('Format Cells…', { exact: true })).toBeVisible()
      await sheets.getByText('Pick From Drop-down List…', { exact: true }).click()

      for (const option of ['alpha', 'beta', 'gamma']) {
        await expect(sheets.getByText(option, { exact: true })).toBeVisible()
      }
      await sheets.getByText('gamma', { exact: true }).click()

      await expect(async () => {
        const values = await getValues(sheets, 2, 1, 1, 1)
        expect(values[0]?.[0]).toBe('gamma')
      }).toPass({ timeout: 10_000 })

      // Filter ▸ Filter by Selected Cell's Value creates the auto-filter and hides 'beta'
      await sheets.mouse.click(b3.x, b3.y, { button: 'right' })
      await sheets.getByText('Filter', { exact: true }).last().hover()
      await sheets.getByText("Filter by Selected Cell's Value", { exact: true }).click()
      await expect(async () => {
        const filter = await sheets.evaluate(() => {
          const sheet = (window as unknown as DebugWindow).__genofficeDebug.univerAPI
            .getActiveWorkbook()
            .getActiveSheet()
          const active = sheet.getFilter()
          return active
            ? { criteria: active.getColumnFilterCriteria(1), hidden: active.getFilteredOutRows() }
            : null
        })
        expect(filter?.criteria).toEqual({ colId: 1, filters: { filters: ['gamma'] } })
        expect(filter?.hidden).toEqual([1])
      }).toPass({ timeout: 10_000 })
    } finally {
      await closeAndSaveVideo(launched, 'sheets-menu-pick-from-list')
    }
  })
})
