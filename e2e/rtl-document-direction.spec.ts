import { test, expect, type Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * An RTL UI must mirror the chrome, never the document (genoffice#1861).
 *
 * `dir="rtl"` on <html> is inherited by everything, so an English document
 * opened under the Arabic UI was laid out — and, for the canvas editors, drawn —
 * right-to-left on screen while its saved content stayed left-to-right. The six
 * document surfaces pin `direction: ltr`; these cases assert the *computed*
 * style, which is what Konva and Univer pick up through their inherited
 * `ctx.direction`.
 */

/** The Home screen's new-document cards, in the order the shell lists them. */
const APPS = [
  { name: 'docs', card: 0, url: '://docs/', host: '.editor-scroll' },
  { name: 'sheets', card: 1, url: '://sheets/', host: '#univer-container' },
  { name: 'slides', card: 2, url: '://slides/', host: '.stage-wrap' },
  { name: 'markdown', card: 3, url: '://markdown/', host: '.editor-scroll' },
  { name: 'html', card: 4, url: '://html/', host: '.preview-stage' },
  { name: 'pdf', card: 5, url: '://pdf/', host: '.pdf-scroll' },
] as const

const directionOf = (page: Page, host: string) =>
  page.evaluate((sel) => {
    const el = document.querySelector(sel)
    return el ? getComputedStyle(el).direction : null
  }, host)

test('the UI chrome itself mirrors under ar', async () => {
  test.setTimeout(180_000)
  const launched = await launchShell({ onboardingSeen: true, lang: 'ar', videoDir: 'rtl-chrome' })
  try {
    await expect(launched.page.locator('html')).toHaveAttribute('dir', 'rtl')
  } finally {
    await closeAndSaveVideo(launched, 'rtl-chrome')
  }
})

for (const { name, card, url, host } of APPS) {
  test(`${name}: a new document's surface stays ltr`, async () => {
    test.setTimeout(240_000)
    const launched = await launchShell({
      onboardingSeen: true,
      lang: 'ar',
      videoDir: `rtl-${name}`,
    })
    try {
      await launched.page.locator('.quick-card').nth(card).click()
      const page = await waitForPageWithUrl(launched.app, url)
      const surface = page.locator(host).first()
      await expect(surface).toBeVisible({ timeout: 45_000 })
      // the assertion the review asked for
      expect(await directionOf(page, host)).toBe('ltr')
      // and what depends on it: an RTL scroller reports a negative scrollLeft,
      // which is what sent zoom-to-cursor jumping in docs, pdf and slides
      const scrollLeft = await surface.evaluate((el) => el.scrollLeft)
      expect(scrollLeft).toBeGreaterThanOrEqual(0)
    } finally {
      await closeAndSaveVideo(launched, `rtl-${name}`)
    }
  })
}
