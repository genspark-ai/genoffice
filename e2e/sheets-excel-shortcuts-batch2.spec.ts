import { test, expect } from '@playwright/test'
import { mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

process.env.GENOFFICE_DEBUG_HOOKS = '1'

/**
 * Excel selection / editing chords added on top of Univer: Ctrl+Shift+*
 * current region, Shift+Home and Ctrl+Shift+Home extension, Ctrl+2 bold,
 * Ctrl+' formula copy from the cell above.
 */

interface Range {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
}

interface DebugApi {
  univerAPI: {
    getActiveWorkbook(): {
      getActiveRange(): { getRange(): Range } | null
      setActiveRange(range: unknown): unknown
      getActiveSheet(): {
        getRange(
          row: number,
          column: number,
        ): {
          setValue(v: number | string): Promise<unknown>
          getFormula(): string
          getCellStyleData(): { bl?: number } | null
        }
      }
    }
  }
}

// evaluated inside the page: no closure over spec-file helpers
const DEBUG_API = 'debugApi'
declare function debugApi(): DebugApi
const INSTALL_DEBUG_API = `window.${DEBUG_API} = () => window.__genofficeDebug`

async function selection(sheets: Page): Promise<Range | null> {
  return sheets.evaluate(() => {
    const range = debugApi().univerAPI.getActiveWorkbook().getActiveRange()?.getRange()
    return range
      ? {
          startRow: range.startRow,
          endRow: range.endRow,
          startColumn: range.startColumn,
          endColumn: range.endColumn,
        }
      : null
  })
}

test.describe('sheets: Excel selection and editing chords', () => {
  test('current region, shift+home extension, ctrl+2 bold, ctrl+quote', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-xl-keys-'))
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'sheets-xl-keys' })
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
      await sheets.evaluate(INSTALL_DEBUG_API)

      const grid = await sheets.evaluate(() => {
        for (const canvas of document.querySelectorAll('canvas')) {
          const rect = canvas.getBoundingClientRect()
          if (rect.width > 500 && rect.height > 300) return { x: rect.x, y: rect.y }
        }
        return null
      })
      if (!grid) throw new Error('worksheet canvas not found')

      // B2:C3 block plus a lone formula in D5
      await sheets.evaluate(async () => {
        const sheet = debugApi().univerAPI.getActiveWorkbook().getActiveSheet()
        await sheet.getRange(1, 1).setValue(1)
        await sheet.getRange(1, 2).setValue(2)
        await sheet.getRange(2, 1).setValue(3)
        await sheet.getRange(2, 2).setValue(4)
        await sheet.getRange(4, 3).setValue('=SUM(B2:C3)')
      })
      // click A1 to focus the grid, then place the cursor through the Facade
      // (default row heights / column widths differ per platform)
      await sheets.mouse.click(grid.x + 46 + 43, grid.y + 24 + 12)
      await expect(sheets.locator('[data-u-comp="defined-name"] input')).toHaveValue('A1')
      const activate = async (row: number, column: number): Promise<void> => {
        await sheets.evaluate(
          ([r, c]) => {
            const workbook = debugApi().univerAPI.getActiveWorkbook()
            workbook.setActiveRange(workbook.getActiveSheet().getRange(r, c))
          },
          [row, column],
        )
      }

      await activate(2, 2)
      await expect(sheets.locator('[data-u-comp="defined-name"] input')).toHaveValue('C3')

      await sheets.keyboard.press('ControlOrMeta+Shift+Digit8')
      await expect
        .poll(() => selection(sheets))
        .toEqual({ startRow: 1, endRow: 2, startColumn: 1, endColumn: 2 })

      // anchor stays on C3; the far corner slides to column A
      await sheets.keyboard.press('Shift+Home')
      await expect
        .poll(() => selection(sheets))
        .toEqual({ startRow: 1, endRow: 2, startColumn: 0, endColumn: 2 })

      await sheets.keyboard.press('ControlOrMeta+Digit2')
      await expect
        .poll(() =>
          sheets.evaluate(
            () =>
              debugApi()
                .univerAPI.getActiveWorkbook()
                .getActiveSheet()
                .getRange(1, 0)
                .getCellStyleData()?.bl,
          ),
        )
        .toBe(1)

      await sheets.keyboard.press('ControlOrMeta+Shift+Home')
      await expect
        .poll(() => selection(sheets))
        .toEqual({ startRow: 0, endRow: 2, startColumn: 0, endColumn: 2 })

      // Ctrl+' copies the formula above verbatim, no reference shift
      await activate(5, 3)
      await expect(sheets.locator('[data-u-comp="defined-name"] input')).toHaveValue('D6')
      await sheets.keyboard.press('ControlOrMeta+Quote')
      await expect
        .poll(() =>
          sheets.evaluate(() =>
            debugApi().univerAPI.getActiveWorkbook().getActiveSheet().getRange(5, 3).getFormula(),
          ),
        )
        .toBe('=SUM(B2:C3)')
    } finally {
      await closeAndSaveVideo(launched, 'sheets-xl-keys')
    }
  })
})
