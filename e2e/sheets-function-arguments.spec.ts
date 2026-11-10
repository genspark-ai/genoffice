import { test, expect } from '@playwright/test'
import { mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

process.env.GENOFFICE_DEBUG_HOOKS = '1'

interface SheetFacade {
  getActiveWorkbook(): {
    getActiveSheet(): {
      getRange(
        row: number,
        column: number,
        rows: number,
        columns: number,
      ): {
        setValues(v: unknown[][]): Promise<unknown>
        activate(): unknown
        getFormulas(): unknown[][]
      }
    }
  }
}

test('Insert Function leads into Function Arguments; fx reopens the call pre-filled', async () => {
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-fn-args-'))
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'sheets-function-arguments',
  })
  try {
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

    const debugSheet = (fn: string) =>
      sheets.evaluate((body) => {
        const debug = (window as unknown as { __genofficeDebug: { univerAPI: SheetFacade } })
          .__genofficeDebug
        const sheet = debug.univerAPI.getActiveWorkbook().getActiveSheet()
        return new Function('sheet', body)(sheet) as unknown
      }, fn)
    await debugSheet('return sheet.getRange(0, 0, 3, 1).setValues([[1], [2], [3]])')
    // the grid needs real focus for Shift+F3; the target is then set programmatically
    await sheets.mouse.click(grid.x + 46 + 43, grid.y + 24 + 12)
    await debugSheet('sheet.getRange(4, 0, 1, 1).activate()')

    await sheets.keyboard.press('Shift+F3')
    const insert = sheets.locator('.insert-function-dialog')
    await expect(insert).toBeVisible()
    await insert.locator('input').first().fill('SUM')
    await insert
      .locator('.fn-row')
      .filter({ hasText: /^SUM\(/ })
      .first()
      .click()
    await insert.locator('.dialog-actions .primary-action').click()

    const args = sheets.locator('.function-args-dialog')
    await expect(args).toBeVisible()
    await expect(args.locator('.fn-args-name')).toHaveText('SUM')
    await expect(args.locator('.fn-args-row label').first()).toHaveClass(/required/)
    await args.locator('#fn-arg-0').fill('A1:A3')
    await expect(args.locator('.fn-args-row output').first()).toHaveText('= {1;2;3}')
    await expect(args.locator('.fn-args-result strong')).toHaveText('6')
    // a second repeating row appears as number2
    await expect(args.locator('.fn-args-row')).toHaveCount(2)
    await sheets.screenshot({ path: 'test-results/function-arguments.png' })

    // a grid click fills the active argument (number2)
    await args.locator('#fn-arg-1').focus()
    await sheets.mouse.click(grid.x + 46 + 43 + 73, grid.y + 24 + 12)
    await expect(args.locator('#fn-arg-1')).toHaveValue('B1')
    await args.locator('#fn-arg-1').fill('')
    await args.locator('.dialog-actions .primary-action').click()
    await expect(args).toBeHidden()
    await expect(async () => {
      expect(await debugSheet('return sheet.getRange(4, 0, 1, 1).getFormulas()')).toEqual([
        ['=SUM(A1:A3)'],
      ])
    }).toPass({ timeout: 10_000 })

    // fx on the formula cell skips the catalog and edits SUM in place
    await debugSheet('sheet.getRange(4, 0, 1, 1).activate()')
    await sheets
      .locator('[data-u-comp="formula-bar"] .univer-w-20 > span:last-child')
      .first()
      .click()
    await expect(args).toBeVisible()
    await expect(args.locator('.fn-args-name')).toHaveText('SUM')
    await expect(args.locator('#fn-arg-0')).toHaveValue('A1:A3')
    await args.locator('#fn-arg-0').fill('A1:A2')
    await args.locator('.dialog-actions .primary-action').click()
    await expect(args).toBeHidden()
    await expect(async () => {
      expect(await debugSheet('return sheet.getRange(4, 0, 1, 1).getFormulas()')).toEqual([
        ['=SUM(A1:A2)'],
      ])
    }).toPass({ timeout: 10_000 })
  } finally {
    await closeAndSaveVideo(launched, 'sheets-function-arguments')
  }
})
