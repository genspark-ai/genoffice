import { test, expect } from '@playwright/test'
import { mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, screenshotPath } from './helpers'

// the preload exposes window.__genofficeDebug only under this env var; the
// spec reads cell data and cell geometry through Univer's Facade
process.env.GENOFFICE_DEBUG_HOOKS = '1'

/**
 * Home › Superscript / Subscript (issue #2212). Two halves:
 *  - a closed cell gets a whole-cell vertAlign, and the grid must PAINT it:
 *    Univer's plain-text painter ignores `va`, so the glyphs used to stay at
 *    the baseline at full size while the button lit up;
 *  - inside the cell editor the button formats the selected characters only
 *    ("1st" with a raised "st"), like Excel, instead of restyling the cell.
 */

interface InkBox {
  top: number
  bottom: number
  left: number
  right: number
  height: number
}

interface CellProbe {
  styleVa: number | null | undefined
  cellVa: number | null | undefined
  runs: Array<{ st: number; ed: number; va: number | null | undefined }> | null
  text: string
}

/** Dark pixels painted inside a cell's box (canvas-local CSS px), grid lines excluded. */
async function inkBox(sheets: Page, row: number, col: number): Promise<InkBox | null> {
  return sheets.evaluate(
    ({ row, col }) => {
      const debug = (window as unknown as Record<string, unknown>).__genofficeDebug as {
        univerAPI: {
          getActiveWorkbook(): {
            getActiveSheet(): {
              getRange(
                r: number,
                c: number,
              ): {
                getCellRect(): { x: number; y: number; width: number; height: number }
              }
            }
          }
        }
      }
      const rect = debug.univerAPI
        .getActiveWorkbook()
        .getActiveSheet()
        .getRange(row, col)
        .getCellRect()
      let canvas: HTMLCanvasElement | null = null
      for (const candidate of document.querySelectorAll('canvas')) {
        const box = candidate.getBoundingClientRect()
        if (box.width > 500 && box.height > 300) canvas = candidate
      }
      if (!canvas) throw new Error('worksheet canvas not found')
      const scale = canvas.width / canvas.clientWidth
      const inset = 2
      const x0 = Math.round((rect.x + inset) * scale)
      const y0 = Math.round((rect.y + inset) * scale)
      const w = Math.round((rect.width - inset * 2) * scale)
      const h = Math.round((rect.height - inset * 2) * scale)
      const ctx = canvas.getContext('2d')
      if (!ctx) throw new Error('no 2d context')
      const { data } = ctx.getImageData(x0, y0, w, h)
      let top = Number.POSITIVE_INFINITY
      let bottom = -1
      let left = Number.POSITIVE_INFINITY
      let right = -1
      for (let y = 0; y < h; y += 1) {
        for (let x = 0; x < w; x += 1) {
          const i = (y * w + x) * 4
          const luminance = (data[i] + data[i + 1] + data[i + 2]) / 3
          // text is black (anti-aliased dark); grid lines are light gray
          if (data[i + 3] > 0 && luminance < 128) {
            if (y < top) top = y
            if (y > bottom) bottom = y
            if (x < left) left = x
            if (x > right) right = x
          }
        }
      }
      if (bottom < 0) return null
      return {
        top: top / scale,
        bottom: bottom / scale,
        left: left / scale,
        right: right / scale,
        height: (bottom - top + 1) / scale,
      }
    },
    { row, col },
  )
}

async function probeCell(sheets: Page, row: number, col: number): Promise<CellProbe> {
  return sheets.evaluate(
    ({ row, col }) => {
      const debug = (window as unknown as Record<string, unknown>).__genofficeDebug as {
        univerAPI: {
          getActiveWorkbook(): {
            getActiveSheet(): {
              getRange(
                r: number,
                c: number,
              ): {
                getCellStyleData(): { va?: number | null } | null
                getCellData(): {
                  v?: unknown
                  s?: unknown
                  p?: {
                    body?: {
                      dataStream?: string
                      textRuns?: Array<{ st: number; ed: number; ts?: { va?: number | null } }>
                    }
                  }
                } | null
              }
            }
          }
        }
      }
      const range = debug.univerAPI.getActiveWorkbook().getActiveSheet().getRange(row, col)
      const style = range.getCellStyleData()
      const cell = range.getCellData()
      const cellStyle = cell?.s
      const runs = cell?.p?.body?.textRuns
      return {
        styleVa: style?.va,
        cellVa:
          cellStyle && typeof cellStyle === 'object'
            ? (cellStyle as { va?: number | null }).va
            : undefined,
        runs: runs ? runs.map((run) => ({ st: run.st, ed: run.ed, va: run.ts?.va })) : null,
        text: cell?.p?.body?.dataStream ?? String(cell?.v ?? ''),
      }
    },
    { row, col },
  )
}

test.describe('sheets: superscript and subscript', () => {
  test('whole-cell vertAlign is painted; in the editor it formats the selection', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-superscript-'))
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-superscript',
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
      // cell centers: 46px row header, 24px column header, 20px rows
      const cellA1 = { x: grid.x + 46 + 43, y: grid.y + 24 + 10 }
      const cellA2 = { x: grid.x + 46 + 43, y: grid.y + 24 + 30 }
      const parked = { x: grid.x + 46 + 43 + 87 * 3, y: grid.y + 24 + 10 + 20 * 6 }

      // ── whole cell: A1 = "H2O", baseline measurement with the selection parked away ──
      await sheets.mouse.click(cellA1.x, cellA1.y)
      await sheets.keyboard.type('H2O', { delay: 30 })
      await sheets.keyboard.press('Enter')
      await sheets.mouse.click(parked.x, parked.y)
      await sheets.waitForTimeout(300)
      const baseline = await inkBox(sheets, 0, 0)
      expect(baseline, 'A1 paints its text').not.toBeNull()

      // Superscript on the closed cell: style set, glyphs smaller and raised
      await sheets.mouse.click(cellA1.x, cellA1.y)
      await sheets.getByRole('button', { name: 'Superscript' }).click()
      await expect
        .poll(async () => (await probeCell(sheets, 0, 0)).styleVa, { timeout: 5_000 })
        .toBe(3)
      await sheets.mouse.click(parked.x, parked.y)
      await sheets.waitForTimeout(300)
      const raised = await inkBox(sheets, 0, 0)
      expect(raised, 'superscript A1 still paints its text').not.toBeNull()
      await sheets.screenshot({ path: screenshotPath('sheets-superscript-whole-cell') })
      expect(raised!.height, 'superscript glyphs are smaller').toBeLessThan(baseline!.height - 1)
      expect(raised!.bottom, 'superscript glyphs sit above the baseline').toBeLessThan(
        baseline!.bottom - 1,
      )

      // pressing the active button again returns to the baseline
      await sheets.mouse.click(cellA1.x, cellA1.y)
      await sheets.getByRole('button', { name: 'Superscript' }).click()
      await expect
        .poll(async () => (await probeCell(sheets, 0, 0)).styleVa ?? null, { timeout: 5_000 })
        .toBeNull()
      await sheets.mouse.click(parked.x, parked.y)
      await sheets.waitForTimeout(300)
      const restored = await inkBox(sheets, 0, 0)
      expect(restored).not.toBeNull()
      expect(Math.abs(restored!.height - baseline!.height)).toBeLessThanOrEqual(1)
      expect(Math.abs(restored!.bottom - baseline!.bottom)).toBeLessThanOrEqual(1)

      // Subscript: smaller glyphs, lowered
      await sheets.mouse.click(cellA1.x, cellA1.y)
      await sheets.getByRole('button', { name: 'Subscript' }).click()
      await expect
        .poll(async () => (await probeCell(sheets, 0, 0)).styleVa, { timeout: 5_000 })
        .toBe(2)
      await sheets.mouse.click(parked.x, parked.y)
      await sheets.waitForTimeout(300)
      const lowered = await inkBox(sheets, 0, 0)
      expect(lowered, 'subscript A1 still paints its text').not.toBeNull()
      expect(lowered!.height, 'subscript glyphs are smaller').toBeLessThan(baseline!.height - 1)
      expect(lowered!.top, 'subscript glyphs start lower').toBeGreaterThan(baseline!.top + 1)

      // ── in the editor: A2 = "1st", select "st", Superscript formats just those runs ──
      await sheets.mouse.click(cellA2.x, cellA2.y)
      await sheets.keyboard.type('1st', { delay: 30 })
      await sheets.keyboard.press('Enter')
      await sheets.mouse.click(cellA2.x, cellA2.y)
      await sheets.keyboard.press('F2')
      await sheets.waitForTimeout(300)
      await sheets.keyboard.press('Shift+ArrowLeft')
      await sheets.keyboard.press('Shift+ArrowLeft')
      await sheets.getByRole('button', { name: 'Superscript' }).click()
      await sheets.waitForTimeout(200)
      await sheets.keyboard.press('Enter')
      await expect
        .poll(async () => (await probeCell(sheets, 1, 0)).runs?.some((run) => run.va === 3), {
          timeout: 5_000,
        })
        .toBe(true)
      const partial = await probeCell(sheets, 1, 0)
      await sheets.screenshot({ path: screenshotPath('sheets-superscript-partial') })
      expect(partial.text.startsWith('1st')).toBe(true)
      const rawRuns = partial.runs ?? []
      const raisedRun = rawRuns.find((run) => run.va === 3)
      expect(raisedRun).toEqual(expect.objectContaining({ st: 1, ed: 3 }))
      // the "1" stays on the baseline and the cell itself was not restyled
      expect(rawRuns.some((run) => run.st === 0 && run.ed > 0 && run.va === 3)).toBe(false)
      expect(partial.styleVa ?? null).toBeNull()
      expect(partial.cellVa ?? null).toBeNull()
    } finally {
      await closeAndSaveVideo(launched, 'sheets-superscript')
    }
  })
})
