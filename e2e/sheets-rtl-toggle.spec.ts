import { test, expect } from '@playwright/test'
import { execSync } from 'node:child_process'
import { copyFile, mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, screenshotPath } from './helpers'

/**
 * The View tab's Right-to-Left toggle (genoffice#1963): the checkbox flips the active
 * sheet's grid direction live (sheetView/@rightToLeft through Univer's own
 * command, mirrored by rtl-grid-mirror) and the echo follows the active
 * sheet's journal + live config. Persistence rides the save pipeline
 * (page-setup journal → sheetView attribute, unit-covered at both ends).
 */
test('the View tab toggles the sheet grid direction', async () => {
  test.setTimeout(240_000)
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'sheets-rtl-toggle' })
  try {
    await launched.page.locator('.quick-card').nth(1).click()
    const page = await waitForPageWithUrl(launched.app, '://sheets/')
    // the sheet tab strip is DOM; a blank workbook's sheet must exist before
    // the toggle has an active worksheet to flip
    await page.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
      timeout: 30_000,
    })
    await page.locator('.ribbon-tabs button', { hasText: /^View$/ }).click()
    const rtlItem = page.locator('.check-item', { hasText: 'Right-to-Left' })
    await expect(rtlItem).toBeVisible()
    await expect(rtlItem.locator('.check-box')).toHaveText('')
    await rtlItem.click()
    await expect(rtlItem.locator('.check-box')).toHaveText('✓')
    await page.waitForTimeout(1_500)
    await page.screenshot({ path: screenshotPath('sheets-rtl-on') })
    // flipping back exercises the LTR restore (mirrored scene teardown)
    await rtlItem.click()
    await expect(rtlItem.locator('.check-box')).toHaveText('')
  } finally {
    await closeAndSaveVideo(launched, 'sheets-rtl-toggle')
  }
})

const FIXTURE = resolve(__dirname, '../apps/sheets/fixtures/generated/compatibility-basic.xlsx')

function sheetXml(workbookPath: string): string {
  return execSync(`unzip -p "${workbookPath}" xl/worksheets/sheet1.xml`).toString()
}

/** A save after toggling must carry sheetView/@rightToLeft — the checkbox
 * echo alone reads the live worksheet config, which stays correct even if
 * the journaling from the mutation never runs (the review finding). */
test('a save after toggling carries rightToLeft in the file', async () => {
  test.setTimeout(240_000)
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-sheets-rtl-e2e-'))
  const workbook = join(scratch, 'rtl-toggle.xlsx')
  await copyFile(FIXTURE, workbook)
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'sheets-rtl-save',
    openFile: workbook,
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://sheets/')
    await page.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
      timeout: 30_000,
    })
    await page.locator('.ribbon-tabs button', { hasText: /^View$/ }).click()
    const rtlItem = page.locator('.check-item', { hasText: 'Right-to-Left' })
    await expect(rtlItem.locator('.check-box')).toHaveText('')

    await rtlItem.click()
    await expect(rtlItem.locator('.check-box')).toHaveText('✓')
    // File > Save, routed to the sheets view the same way the app menu does it
    await launched.app.evaluate(({ webContents }) => {
      const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('://sheets/'))
      wc?.send('menu:action', 'save')
    })
    await expect(() => {
      expect(sheetXml(workbook)).toContain('rightToLeft="1"')
    }).toPass({ timeout: 15_000 })

    // flipping back and saving drops the attribute again (default restored)
    await rtlItem.click()
    await expect(rtlItem.locator('.check-box')).toHaveText('')
    await launched.app.evaluate(({ webContents }) => {
      const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('://sheets/'))
      wc?.send('menu:action', 'save')
    })
    await expect(() => {
      expect(sheetXml(workbook)).not.toContain('rightToLeft')
    }).toPass({ timeout: 15_000 })
  } finally {
    await closeAndSaveVideo(launched, 'sheets-rtl-save')
  }
})
