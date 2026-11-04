/**
 * Right-clicking a picture opens Word's picture menu (clipboard, picture,
 * wrap / arrange, link, comment), not the text menu with Font / Paragraph /
 * Synonyms / Translate and the table block on top.
 */
import { test, expect } from '@playwright/test'
import { resolve } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

const DOC = resolve(__dirname, '../fixtures/generated/kitchen-sink.docx')

test('picture context menu lists only Word picture items', async () => {
  test.setTimeout(120_000)
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-picture-context-menu',
    openFile: DOC,
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    await page.waitForFunction(
      () => Boolean((window as unknown as { __aidocs?: { editor?: unknown } }).__aidocs?.editor),
      undefined,
      { timeout: 30_000 },
    )
    const img = page.locator('.doc-page[contenteditable="true"] img').first()
    await img.scrollIntoViewIfNeeded()
    await expect(img).toBeVisible()
    const box = (await img.boundingBox())!
    await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2, { button: 'right' })
    const menu = page.locator('.ctx-menu')
    await menu.waitFor({ timeout: 15_000 })
    await page.screenshot({ path: test.info().outputPath('picture-menu.png') })
    const labels = await menu
      .locator('.ctx-label')
      .evaluateAll((els) => els.map((e) => e.textContent))
    expect(labels.slice(0, 5)).toEqual(['Cut', 'Copy', 'Paste', 'View Image', 'Save Image As…'])
    expect(labels).toContain('Wrap Text')
    for (const gone of ['Paste as Plain Text', 'Font…', 'Paragraph…', 'Synonyms', 'Translate']) {
      expect(labels).not.toContain(gone)
    }
    expect(labels.length).toBeLessThanOrEqual(9)
  } finally {
    await closeAndSaveVideo(launched)
  }
})
