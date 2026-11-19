import { test, expect, type ElectronApplication } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

// Word for Mac's menu bar order; Docs mirrors it so a Word user finds each menu where it was.
const WORD_MENU_ORDER = [
  'File',
  'Edit',
  'View',
  'Insert',
  'Format',
  'Tools',
  'Table',
  'Window',
  'Help',
]

const topLevelLabels = (app: ElectronApplication): Promise<string[]> =>
  app.evaluate(({ Menu }) => (Menu.getApplicationMenu()?.items ?? []).map((item) => item.label))

test('the Docs menu bar follows the Word for Mac order', async () => {
  test.setTimeout(120_000)
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-menu-order' })
  const { page, app } = launched
  try {
    await page.evaluate(() =>
      (window as never as { aiOffice: { newDoc(): Promise<void> } }).aiOffice.newDoc(),
    )
    await waitForPageWithUrl(app, '://docs/')
    await expect
      .poll(async () => (await topLevelLabels(app)).filter((l) => WORD_MENU_ORDER.includes(l)), {
        timeout: 30_000,
      })
      .toEqual(WORD_MENU_ORDER)
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-order')
  }
})
