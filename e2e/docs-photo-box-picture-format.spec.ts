/**
 * A behind-text cover photo anchored in the same paragraph as text boxes is
 * a picture: pressing it opens Picture Format for the photo (its own wrap
 * mode), not Shape Format for the paragraph's first text box.
 */
import { test, expect } from '@playwright/test'
import { join } from 'node:path'
import { launchShell, waitForPageWithUrl } from './helpers'

test('pressing a photo box opens Picture Format with the photo wrap', async () => {
  test.setTimeout(120_000)
  const { app } = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-photo-box',
    openFile: join(__dirname, 'assets', 'cover-photo-with-textboxes.docx'),
  })
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    const pic = page.locator('.doc-page .doc-textbox-picture').first()
    await pic.waitFor({ timeout: 30_000 })
    const box = (await pic.boundingBox())!
    await page.mouse.click(box.x + 40, box.y + box.height - 40)
    await expect(page.locator('.ribbon-tab.active')).toHaveText('Picture Format')
    await expect(page.locator('.rb-wrap-dd').first()).toContainText('Behind Text')
  } finally {
    await app.close()
  }
})
