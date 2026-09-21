import { test, expect } from '@playwright/test'
import { copyFile, mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import type { Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

const FIXTURE = resolve(__dirname, '../apps/sheets/fixtures/generated/compatibility-basic.xlsx')

async function waitForWorkbook(page: Page, sheetName = 'Sheet1'): Promise<void> {
  await page.waitForFunction((name) => document.body.textContent?.includes(name), sheetName, {
    timeout: 30_000,
  })
  await page.waitForTimeout(1_500)
}

// The dialog card is a single-column grid: the tab row cannot compress below
// its labels, so the column track — and with it every row — used to outgrow
// the fixed-width card in locales with long tab labels (fr needs ~490px vs
// the 420px base; even en overflowed by a few px). The card now sizes to its
// widest incompressible row, so nothing may paint past its padding edge.
for (const lang of ['fr', 'en']) {
  test(`sheets: Format Cells dialog content stays inside the card (${lang})`, async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-fmtcells-e2e-'))
    const workbook = join(scratch, 'format-cells.xlsx')
    await copyFile(FIXTURE, workbook)

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: `sheets-format-cells-overflow-${lang}`,
      openFile: workbook,
      lang,
    })
    try {
      const sheets = await waitForPageWithUrl(launched.app, '://sheets/')
      await waitForWorkbook(sheets)

      // focus a grid cell, then open the dialog with its shortcut
      const grid = await sheets.locator('canvas').first().boundingBox()
      if (!grid) throw new Error('no grid canvas')
      await sheets.mouse.click(grid.x + 80, grid.y + 8)
      await sheets.waitForTimeout(400)
      await sheets.keyboard.press('Control+1')
      await sheets.waitForSelector('.format-cells-dialog', { timeout: 10_000 })

      // the currency category renders the widest pane (sample, symbol
      // dropdown, negative-numbers list)
      await sheets.locator('.numfmt-cats button').nth(2).click()
      await sheets.waitForTimeout(300)

      const report = await sheets.evaluate(() => {
        const card = document.querySelector('.format-cells-dialog')
        if (!card) return null
        const rect = card.getBoundingClientRect()
        const style = getComputedStyle(card)
        const innerRight =
          rect.right - parseFloat(style.paddingRight) - parseFloat(style.borderRightWidth)
        const parts: Record<string, number> = {}
        for (const [name, selector] of [
          ['tabs', '.dialog-tabs'],
          ['sample', '.numfmt-sample output'],
          ['negativesList', '.numfmt-list'],
          ['actions', '.dialog-actions'],
        ] as const) {
          const el = card.querySelector(selector)
          if (el) parts[name] = el.getBoundingClientRect().right
        }
        return { innerRight, parts }
      })
      expect(report).not.toBeNull()
      const { innerRight, parts } = report!
      expect(Object.keys(parts)).toContain('tabs')
      for (const [name, right] of Object.entries(parts)) {
        expect(right, `${name} paints past the dialog card`).toBeLessThanOrEqual(innerRight + 1)
      }
    } finally {
      await closeAndSaveVideo(launched)
    }
  })
}
