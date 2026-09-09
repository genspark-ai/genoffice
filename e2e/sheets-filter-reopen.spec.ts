import { test, expect } from '@playwright/test'
import { execSync } from 'node:child_process'
import { mkdtemp, readdir } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

// the preload exposes window.__genofficeDebug only under this env var
process.env.GENOFFICE_DEBUG_HOOKS = '1'

/**
 * Regression for "filter dropdown selections vanish after reopening the
 * file" (alpha feedback): the save wrote each column's criteria into
 * the xlsx autoFilter, but reopening only restored the filter range — the
 * criteria were lost and the filtered-out rows came back as plain manual
 * hides, so the dropdown lost its checked values and other columns' lists
 * were no longer narrowed. The reopen must restore the criteria AND hand the
 * hidden rows back to the filter model, so a later criteria change can
 * unhide them.
 */
test.describe('sheets: filter criteria survive save and reopen', () => {
  test('criteria restore, rows stay filtered, and re-filtering unhides', async () => {
    test.setTimeout(180_000)
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-filter-reopen-'))
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-filter-reopen',
    })
    let savedPath = ''
    try {
      const { app, page } = launched
      await app.evaluate(({ app: electronApp }, dir) => {
        electronApp.setPath('documents', dir)
      }, scratch)

      await expect(page.locator('.quick-card').nth(1)).toContainText('AI Sheets')
      await page.locator('.quick-card').nth(1).click()

      const sheets = await waitForPageWithUrl(app, 'sheets/out')
      await sheets.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
        timeout: 30_000,
      })
      await sheets.waitForTimeout(1_500)

      // A1:B5 — header row plus four data rows, then filter B to "keep"
      await sheets.evaluate(async () => {
        const debug = (window as unknown as Record<string, unknown>).__genofficeDebug as {
          univerAPI: { getActiveWorkbook(): { getActiveSheet(): unknown } }
        }
        const sheet = debug.univerAPI.getActiveWorkbook().getActiveSheet() as {
          getRange(
            row: number,
            column: number,
            rows: number,
            columns: number,
          ): { setValues(values: unknown[][]): Promise<unknown>; createFilter(): unknown }
          getFilter(): {
            setColumnFilterCriteria(column: number, criteria: unknown): unknown
          } | null
        }
        await sheet.getRange(0, 0, 5, 2).setValues([
          ['Name', 'Cat'],
          ['a', 'keep'],
          ['b', 'drop'],
          ['c', 'keep'],
          ['d', 'drop'],
        ])
        sheet.getRange(0, 0, 5, 2).createFilter()
        sheet.getFilter()?.setColumnFilterCriteria(1, {
          colId: 1,
          filters: { filters: ['keep'] },
        })
      })
      await sheets.waitForTimeout(800)

      await app.evaluate(({ webContents }) => {
        const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('sheets/out'))
        wc?.send('menu:action', 'save')
      })

      const saveDir = join(scratch, 'GenOffice')
      await expect(async () => {
        const files = (await readdir(saveDir)).filter((f) => f.endsWith('.xlsx'))
        expect(files).toHaveLength(1)
        savedPath = join(saveDir, files[0])
        const xml = execSync(`unzip -p "${savedPath}" xl/worksheets/sheet1.xml`).toString()
        expect(xml).toContain('<autoFilter ref="A1:B5">')
        expect(xml).toContain('<filterColumn colId="1"><filters><filter val="keep"/></filters>')
        expect(xml).toMatch(/<row r="3"[^>]* hidden="1"/)
        expect(xml).toMatch(/<row r="5"[^>]* hidden="1"/)
      }).toPass({ timeout: 15_000 })
    } finally {
      await closeAndSaveVideo(launched, 'sheets-filter-reopen')
    }

    // reopen the saved file in a fresh app instance
    const relaunched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-filter-reopen-reopened',
      openFile: savedPath,
    })
    try {
      const { app } = relaunched
      const sheets = await waitForPageWithUrl(app, 'sheets/out')
      await sheets.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
        timeout: 30_000,
      })

      // the restore runs once the sheet finishes indexing — poll for it
      await sheets.waitForFunction(
        () => {
          const debug = (window as unknown as Record<string, unknown>).__genofficeDebug as {
            univerAPI: { getActiveWorkbook(): { getActiveSheet(): unknown } }
          }
          const sheet = debug.univerAPI.getActiveWorkbook().getActiveSheet() as {
            getFilter(): {
              getColumnFilterCriteria(column: number): unknown
            } | null
          }
          return sheet.getFilter()?.getColumnFilterCriteria(1) != null
        },
        null,
        { timeout: 30_000 },
      )

      const restored = await sheets.evaluate(() => {
        const debug = (window as unknown as Record<string, unknown>).__genofficeDebug as {
          univerAPI: { getActiveWorkbook(): { getActiveSheet(): unknown } }
        }
        const sheet = debug.univerAPI.getActiveWorkbook().getActiveSheet() as {
          getFilter(): {
            getColumnFilterCriteria(column: number): { filters?: { filters?: string[] } } | null
            getFilteredOutRows(): number[]
          } | null
          getSheet(): { getRowRawVisible(row: number): boolean }
        }
        const filter = sheet.getFilter()
        return {
          criteria: filter?.getColumnFilterCriteria(1)?.filters?.filters ?? null,
          filteredOut: filter?.getFilteredOutRows() ?? null,
          // raw visibility ignores the filter: true = no manual hd flag left
          rawVisible: [2, 4].map((row) => sheet.getSheet().getRowRawVisible(row)),
        }
      })
      // the dropdown's checked values are back...
      expect(restored.criteria).toEqual(['keep'])
      // ...the filter (not manual hides) owns the hidden rows...
      expect(restored.filteredOut).toEqual([2, 4])
      expect(restored.rawVisible).toEqual([true, true])

      // ...so broadening the filter really unhides them
      await sheets.evaluate(() => {
        const debug = (window as unknown as Record<string, unknown>).__genofficeDebug as {
          univerAPI: { getActiveWorkbook(): { getActiveSheet(): unknown } }
        }
        const sheet = debug.univerAPI.getActiveWorkbook().getActiveSheet() as {
          getFilter(): {
            setColumnFilterCriteria(column: number, criteria: unknown): unknown
          } | null
        }
        sheet.getFilter()?.setColumnFilterCriteria(1, {
          colId: 1,
          filters: { filters: ['keep', 'drop'] },
        })
      })
      await sheets.waitForTimeout(500)
      const widened = await sheets.evaluate(() => {
        const debug = (window as unknown as Record<string, unknown>).__genofficeDebug as {
          univerAPI: { getActiveWorkbook(): { getActiveSheet(): unknown } }
        }
        const sheet = debug.univerAPI.getActiveWorkbook().getActiveSheet() as {
          getFilter(): { getFilteredOutRows(): number[] } | null
        }
        return sheet.getFilter()?.getFilteredOutRows() ?? null
      })
      expect(widened).toEqual([])
    } finally {
      await closeAndSaveVideo(relaunched, 'sheets-filter-reopen-reopened')
    }
  })
})
