import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, screenshotPath } from './helpers'

async function findShellPage(app: ElectronApplication, timeoutMs = 15_000): Promise<Page> {
  const deadline = Date.now() + timeoutMs
  for (;;) {
    for (const candidate of app.windows()) {
      const has = await candidate
        .evaluate(() => Boolean((window as unknown as { aiOffice?: unknown }).aiOffice))
        .catch(() => false)
      if (has) return candidate
    }
    const remaining = deadline - Date.now()
    if (remaining <= 0) throw new Error('No window exposing window.aiOffice')
    await app.waitForEvent('window', { timeout: Math.min(remaining, 1_000) }).catch(() => {})
  }
}

function luminance(rgb: string): number {
  const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(rgb)
  if (!m) throw new Error(`Unparseable color: ${rgb}`)
  return 0.2126 * Number(m[1]) + 0.7152 * Number(m[2]) + 0.0722 * Number(m[3])
}

test('slides chrome darkens while the slide canvas stays paper-white', async () => {
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'theme-visual-slides' })
  try {
    const shellPage = await findShellPage(launched.app)
    await shellPage.locator('.quick-card', { hasText: 'AI Slides' }).click()
    const editorPage = await waitForPageWithUrl(launched.app, 'slides/out')
    await editorPage.waitForSelector('.stage-wrap canvas', { timeout: 20_000 })

    await shellPage.evaluate(() =>
      (window as unknown as { aiOffice: { setTheme(v: string): Promise<void> } }).aiOffice.setTheme(
        'dark',
      ),
    )
    // ribbon and thumbnail rail darken
    await expect
      .poll(async () =>
        luminance(
          await editorPage
            .locator('.ribbon')
            .first()
            .evaluate((el) => getComputedStyle(el).backgroundColor),
        ),
      )
      .toBeLessThan(80)
    // the Konva stage still paints the slide paper white: the slide-bg Rect is
    // canvas content, so assert via the stage screenshot's center pixel
    const shot = await editorPage.locator('.stage-wrap').first().screenshot()
    // PNG: naive check via Playwright — decode through a data URL in the page
    const center = await editorPage.evaluate(async (b64) => {
      const img = new Image()
      img.src = `data:image/png;base64,${b64}`
      await img.decode()
      const c = document.createElement('canvas')
      c.width = img.width
      c.height = img.height
      const ctx = c.getContext('2d')!
      ctx.drawImage(img, 0, 0)
      const d = ctx.getImageData(Math.floor(img.width / 2), Math.floor(img.height / 2), 1, 1).data
      return `rgb(${d[0]}, ${d[1]}, ${d[2]})`
    }, shot.toString('base64'))
    expect(luminance(center)).toBeGreaterThan(180)
    await editorPage.screenshot({ path: screenshotPath('theme-slides-dark') })
  } finally {
    await closeAndSaveVideo(launched, 'theme-visual-slides')
  }
})
