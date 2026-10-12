import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, screenshotPath } from './helpers'

/**
 * The View tab's Right-to-Left toggle (#1963): the checkbox flips the active
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
