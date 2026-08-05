import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, screenshotPath } from './helpers'

test.describe('new file from home', () => {
  test('AI Docs quick card opens a docs editor tab', async () => {
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'new-doc-tab' })
    const { app, page } = launched
    try {
      await expect(page.locator('.quick-card').first()).toContainText('AI Docs')

      await page.locator('.quick-card').first().click()

      const editorTab = page.locator('.tab-bar .tab-item:not(.tab-home)')
      await expect(editorTab).toHaveCount(1)
      await expect(editorTab).toHaveClass(/active/)
      await page.screenshot({ path: screenshotPath('new-doc-tab-bar') })

      // the docs editor loads in a WebContentsView, which surfaces as a new
      // page — poll for it instead of waitForLoadState, which hangs on Linux
      // when Playwright attaches mid-navigation and misses lifecycle events
      const editorPage = await waitForPageWithUrl(app, 'docs/out')
      await expect(editorPage.locator('body')).toBeVisible()
      await editorPage.screenshot({ path: screenshotPath('new-doc-editor') })
    } finally {
      await closeAndSaveVideo(launched, 'new-doc-tab')
    }
  })
})
