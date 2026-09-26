/**
 * Word's zoom range and status bar: the slider spans 10–500 and the page keeps
 * rendering at both ends, the percentage opens the Zoom dialog, the page
 * counter opens Go To, the view buttons switch layouts and Track Changes
 * toggles from the status bar.
 */
import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

const POLL = { timeout: 15_000, intervals: [250, 500, 1000] }

test('zoom 10–500, zoom dialog, page counter, view buttons and track changes', async () => {
  test.setTimeout(120_000)
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-zoom-status-bar' })
  const { app, page } = launched
  try {
    await page.locator('.quick-card').first().click()
    const editor = await waitForPageWithUrl(app, '://docs/')
    const docPage = editor.locator('.doc-page').first()
    await editor.locator('.doc-page[contenteditable="true"]').waitFor()
    await docPage.click()
    await editor.keyboard.type('Zoom range check.', { delay: 10 })
    await expect(docPage).toContainText('Zoom range check.')

    const zoomValue = editor.locator('.zoom-value')
    await expect(zoomValue).toHaveText('100%')
    const width100 = (await docPage.boundingBox())!.width
    const widthRatio = async () => (await docPage.boundingBox())!.width / width100

    // ---- slider ends: 500% and 10%, the page still lays out ----
    const slider = editor.locator('.zoom-slider')
    await expect(slider).toHaveAttribute('min', '10')
    await expect(slider).toHaveAttribute('max', '500')
    await slider.focus()
    await editor.keyboard.press('End')
    await expect(zoomValue).toHaveText('500%')
    await expect.poll(widthRatio, POLL).toBeGreaterThan(4.9)
    await expect(docPage).toContainText('Zoom range check.')
    await editor.keyboard.press('Home')
    await expect(zoomValue).toHaveText('10%')
    await expect.poll(widthRatio, POLL).toBeLessThan(0.11)
    await expect(docPage).toContainText('Zoom range check.')

    // ---- ± buttons step ten and stop at the ends ----
    const zoomBtns = editor.locator('.zoom-btn')
    await zoomBtns.nth(0).click()
    await expect(zoomValue).toHaveText('10%')
    await zoomBtns.nth(1).click()
    await expect(zoomValue).toHaveText('20%')

    // ---- percentage opens the Zoom dialog ----
    await zoomValue.click()
    const dialog = editor.locator('.zoom-dialog')
    await dialog.waitFor()
    await dialog.locator('.zoom-dialog-radio', { hasText: '200%' }).click()
    await expect(dialog.locator('input[type="number"]')).toHaveValue('200')
    await dialog.locator('button[type="submit"]').click()
    await expect(dialog).toHaveCount(0)
    await expect(zoomValue).toHaveText('200%')

    await zoomValue.click()
    const percent = dialog.locator('input[type="number"]')
    await percent.fill('900')
    await percent.press('Enter')
    await expect(dialog).toHaveCount(0)
    await expect(zoomValue).toHaveText('500%')

    await zoomValue.click()
    await percent.fill('0')
    await percent.press('Enter')
    await expect(zoomValue).toHaveText('10%')

    await zoomValue.click()
    await percent.fill('37')
    await percent.press('Enter')
    await expect(zoomValue).toHaveText('37%')

    // ---- View tab Zoom button opens the same dialog ----
    await editor.locator('.ribbon-tab', { hasText: /^View$/ }).click()
    await editor.locator('.rb-big', { hasText: /^Zoom$/ }).click()
    await dialog.waitFor()
    await dialog.locator('button[type="button"]').click()
    await expect(dialog).toHaveCount(0)

    // ---- status bar view buttons ----
    const views = editor.locator('.status-view-btn')
    await expect(views).toHaveCount(4)
    await expect(views.nth(0)).toHaveClass(/\bon\b/)
    await views.nth(1).click()
    await expect(editor.locator('.doc-zoom')).toHaveClass(/view-web/)
    await expect(views.nth(1)).toHaveClass(/\bon\b/)
    await views.nth(2).click()
    await expect(editor.locator('.doc-zoom')).toHaveClass(/view-outline/)
    await views.nth(0).click()
    await expect(editor.locator('.doc-zoom')).toHaveClass(/view-print/)
    await views.nth(3).click()
    await expect(editor.locator('.app')).toHaveClass(/read-mode/)
    await expect(zoomValue).toHaveText('100%')
    await views.nth(3).click()
    await expect(editor.locator('.app')).not.toHaveClass(/read-mode/)
    await expect(zoomValue).toHaveText('37%')

    // ---- page counter opens Go To ----
    await editor.locator('.status-page').click()
    const panel = editor.locator('.find-panel')
    await expect(panel.locator('.find-tab.on')).toHaveText('Go To')
    await panel.locator('.find-close').click()

    // ---- Track Changes readout toggles the recorder ----
    const track = editor.locator('.status-track')
    await expect(track).toHaveText('Track Changes: Off')
    await track.click()
    await expect(track).toHaveText('Track Changes: On')
    await expect(track).toHaveAttribute('aria-pressed', 'true')
    await track.click()
    await expect(track).toHaveText('Track Changes: Off')
  } finally {
    await closeAndSaveVideo(launched, 'docs-zoom-status-bar')
  }
})
