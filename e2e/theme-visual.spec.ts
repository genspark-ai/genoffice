import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
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

function setTheme(page: Page, theme: 'light' | 'dark' | 'system'): Promise<void> {
  return page.evaluate((t) => {
    const api = (window as unknown as { aiOffice: { setTheme(v: string): Promise<void> } }).aiOffice
    return api.setTheme(t)
  }, theme)
}

function bodyBg(page: Page): Promise<string> {
  return page.evaluate(() => getComputedStyle(document.body).backgroundColor)
}

/** relative luminance of a computed rgb()/rgba() string, 0 (black) – 255 (white) */
function luminance(rgb: string): number {
  const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(rgb)
  if (!m) throw new Error(`Unparseable color: ${rgb}`)
  return 0.2126 * Number(m[1]) + 0.7152 * Number(m[2]) + 0.0722 * Number(m[3])
}

/** minimal single-page PDF, xref offsets computed so pdf.js parses it cleanly */
function minimalPdf(): Buffer {
  const objects = [
    '<</Type/Catalog/Pages 2 0 R>>',
    '<</Type/Pages/Kids[3 0 R]/Count 1>>',
    '<</Type/Page/Parent 2 0 R/MediaBox[0 0 612 792]>>',
  ]
  let body = '%PDF-1.4\n'
  const offsets: number[] = []
  objects.forEach((obj, i) => {
    offsets.push(body.length)
    body += `${i + 1} 0 obj\n${obj}\nendobj\n`
  })
  const xrefStart = body.length
  body += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`
  for (const off of offsets) body += `${String(off).padStart(10, '0')} 00000 n \n`
  body += `trailer\n<</Size ${objects.length + 1}/Root 1 0 R>>\nstartxref\n${xrefStart}\n%%EOF\n`
  return Buffer.from(body, 'latin1')
}

test.describe('theme visual adoption', () => {
  test('markdown editor surface follows dark theme', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-theme-md-'))
    const mdPath = join(dir, 'doc.md')
    await writeFile(mdPath, '# Doc\n\nBody.\n')

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'theme-visual-md',
      openFile: mdPath,
    })
    try {
      const shellPage = await findShellPage(launched.app)
      const editorPage = await waitForPageWithUrl(launched.app, 'markdown/out')
      await expect(editorPage.locator('.doc-editor')).toBeVisible()

      const lightBg = await bodyBg(editorPage)
      expect(luminance(lightBg)).toBeGreaterThan(180)

      await setTheme(shellPage, 'dark')
      await expect.poll(async () => luminance(await bodyBg(editorPage))).toBeLessThan(80)
      // editor text is readable in dark
      const textColor = await editorPage
        .locator('.doc-editor')
        .evaluate((el) => getComputedStyle(el).color)
      expect(luminance(textColor)).toBeGreaterThan(180)
      await editorPage.screenshot({ path: screenshotPath('theme-markdown-dark') })
    } finally {
      await closeAndSaveVideo(launched, 'theme-visual-md')
    }
  })

  test('docs chrome darkens while the page stays paper-white', async () => {
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'theme-visual-docs' })
    try {
      const shellPage = await findShellPage(launched.app)
      await shellPage.locator('.quick-card', { hasText: 'AI Docs' }).click()
      const editorPage = await waitForPageWithUrl(launched.app, 'docs/out')
      await expect(editorPage.locator('.doc-page').first()).toBeVisible()

      await setTheme(shellPage, 'dark')
      // ribbon and the canvas gutter around pages darken
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
      expect(
        luminance(
          await editorPage
            .locator('.workspace')
            .first()
            .evaluate((el) => getComputedStyle(el).backgroundColor),
        ),
      ).toBeLessThan(90)
      // the paper and its ink stay untouched
      const page = editorPage.locator('.doc-page').first()
      expect(
        luminance(await page.evaluate((el) => getComputedStyle(el).backgroundColor)),
      ).toBeGreaterThan(180)
      expect(luminance(await page.evaluate((el) => getComputedStyle(el).color))).toBeLessThan(60)
      await editorPage.screenshot({ path: screenshotPath('theme-docs-dark') })
    } finally {
      await closeAndSaveVideo(launched, 'theme-visual-docs')
    }
  })

  test('sheets chrome and Univer grid follow dark theme', async () => {
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'theme-visual-sheets' })
    try {
      const shellPage = await findShellPage(launched.app)
      await shellPage.locator('.quick-card', { hasText: 'AI Sheets' }).click()
      const editorPage = await waitForPageWithUrl(launched.app, 'sheets/out')
      await editorPage.waitForSelector('canvas', { timeout: 20_000 })

      // Univer flags its dark repaint with a class on <html> (ThemeService.darkMode$)
      const univerDark = () =>
        editorPage.evaluate(() => document.documentElement.classList.contains('univer-dark'))
      expect(await univerDark()).toBe(false)

      await setTheme(shellPage, 'dark')
      // our chrome (the ribbon) darkens
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
      // Univer switched its grid palette to dark (official darkMode flag)
      await expect.poll(univerDark).toBe(true)
      await editorPage.screenshot({ path: screenshotPath('theme-sheets-dark') })
    } finally {
      await closeAndSaveVideo(launched, 'theme-visual-sheets')
    }
  })

  test('pdf chrome follows dark theme while the page stays paper-white', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-theme-pdf-'))
    const pdfPath = join(dir, 'doc.pdf')
    await writeFile(pdfPath, minimalPdf())

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'theme-visual-pdf',
      openFile: pdfPath,
    })
    try {
      const shellPage = await findShellPage(launched.app)
      const editorPage = await waitForPageWithUrl(launched.app, 'pdf/out')
      await expect(editorPage.locator('.pdf-page').first()).toBeVisible()

      await setTheme(shellPage, 'dark')
      // chrome (the scroll gutter around pages) darkens
      await expect
        .poll(async () =>
          luminance(
            await editorPage
              .locator('.pdf-scroll')
              .evaluate((el) => getComputedStyle(el).backgroundColor),
          ),
        )
        .toBeLessThan(80)
      // the page itself stays paper-white and un-inverted (night mode is separate)
      const page = editorPage.locator('.pdf-page').first()
      expect(
        luminance(await page.evaluate((el) => getComputedStyle(el).backgroundColor)),
      ).toBeGreaterThan(180)
      await editorPage.screenshot({ path: screenshotPath('theme-pdf-dark') })
    } finally {
      await closeAndSaveVideo(launched, 'theme-visual-pdf')
    }
  })
})
