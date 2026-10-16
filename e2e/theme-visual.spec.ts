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
  test('html editor surface follows dark theme', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-theme-html-'))
    const htmlPath = join(dir, 'doc.html')
    await writeFile(htmlPath, '<html><body><p>Body.</p></body></html>\n')

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'theme-visual-html',
      openFile: htmlPath,
    })
    try {
      const shellPage = await findShellPage(launched.app)
      const editorPage = await waitForPageWithUrl(launched.app, '://html/')
      // preview is the default view; the source pane is what this test measures
      await editorPage.locator('.rb-view', { hasText: /Source/ }).click()
      await expect(editorPage.locator('.source-editor .cm-content')).toBeVisible()

      const lightBg = await bodyBg(editorPage)
      expect(luminance(lightBg)).toBeGreaterThan(180)

      await setTheme(shellPage, 'dark')
      await expect.poll(async () => luminance(await bodyBg(editorPage))).toBeLessThan(80)
      const textColor = await editorPage
        .locator('.source-editor .cm-content')
        .evaluate((el) => getComputedStyle(el).color)
      expect(luminance(textColor)).toBeGreaterThan(180)
      await editorPage.screenshot({ path: screenshotPath('theme-html-dark') })
    } finally {
      await closeAndSaveVideo(launched, 'theme-visual-html')
    }
  })

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
      const editorPage = await waitForPageWithUrl(launched.app, '://markdown/')
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

  test('docs renders a dark page in dark theme; View ▸ Dark Mode switches back to white paper', async () => {
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'theme-visual-docs' })
    try {
      const shellPage = await findShellPage(launched.app)
      await shellPage.locator('.quick-card', { hasText: 'AI Docs' }).click()
      const editorPage = await waitForPageWithUrl(launched.app, '://docs/')
      const page = editorPage.locator('.doc-page').first()
      await expect(page).toBeVisible()
      const pageBg = () => page.evaluate((el) => getComputedStyle(el).backgroundColor)
      const pageInk = () => page.evaluate((el) => getComputedStyle(el).color)

      // light theme: white paper, dark ink
      expect(luminance(await pageBg())).toBeGreaterThan(180)
      expect(luminance(await pageInk())).toBeLessThan(60)

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
      // Word-style dark page: the paper goes dark and the ink is remapped light
      await expect.poll(async () => luminance(await pageBg())).toBeLessThan(60)
      expect(luminance(await pageInk())).toBeGreaterThan(180)
      await editorPage.screenshot({ path: screenshotPath('theme-docs-dark') })

      // View ▸ Dark Mode is Word's Switch Modes: back to white paper under the dark chrome
      await editorPage.getByRole('button', { name: 'View', exact: true }).click()
      await editorPage.getByRole('button', { name: 'Dark Mode', exact: true }).click()
      await expect.poll(async () => luminance(await pageBg())).toBeGreaterThan(180)
      expect(luminance(await pageInk())).toBeLessThan(60)
      await editorPage.screenshot({ path: screenshotPath('theme-docs-dark-white-page') })
    } finally {
      await closeAndSaveVideo(launched, 'theme-visual-docs')
    }
  })

  test('sheets chrome and Univer grid follow dark theme', async () => {
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'theme-visual-sheets' })
    try {
      const shellPage = await findShellPage(launched.app)
      await shellPage.locator('.quick-card', { hasText: 'AI Sheets' }).click()
      const editorPage = await waitForPageWithUrl(launched.app, '://sheets/')
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

  // genoffice#2041: Univer renders its menus from the DOCUMENT theme while
  // `color-scheme` on :root follows the APP theme. With a dark app and the
  // document theme pinned light the UA button text went white, so every
  // context-menu label vanished on the light menu surface.
  test('sheets context menu stays readable with a dark app and a light document theme', async () => {
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'theme-visual-sheets-menu',
      settings: { documentTheme: 'light' },
    })
    try {
      const shellPage = await findShellPage(launched.app)
      await shellPage.locator('.quick-card', { hasText: 'AI Sheets' }).click()
      const editorPage = await waitForPageWithUrl(launched.app, '://sheets/')
      await editorPage.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
        timeout: 30_000,
      })
      await editorPage.waitForTimeout(1_500)

      await setTheme(shellPage, 'dark')
      await expect
        .poll(() => editorPage.evaluate(() => document.documentElement.getAttribute('data-theme')))
        .toBe('dark')
      // the document theme stays pinned light, so Univer keeps its light palette
      await expect
        .poll(() =>
          editorPage.evaluate(() => document.documentElement.getAttribute('data-doc-theme')),
        )
        .toBe('light')
      expect(
        await editorPage.evaluate(() => document.documentElement.classList.contains('univer-dark')),
      ).toBe(false)

      // right-click cell A2 → the Univer cell context menu
      const grid = await editorPage.evaluate(() => {
        for (const canvas of document.querySelectorAll('canvas')) {
          const rect = canvas.getBoundingClientRect()
          if (rect.width > 500 && rect.height > 300) return { x: rect.x, y: rect.y }
        }
        return null
      })
      if (!grid) throw new Error('worksheet canvas not found')
      await editorPage.mouse.click(grid.x + 60, grid.y + 60, { button: 'right' })
      await expect(editorPage.getByText('Number Format', { exact: true })).toBeVisible({
        timeout: 10_000,
      })

      const probe = await editorPage.evaluate(() => {
        const panels = Array.from(
          document.querySelectorAll<HTMLElement>('div.univer-min-w-52'),
        ).filter((el) => {
          const r = el.getBoundingClientRect()
          return r.x >= 0 && r.y >= 0 && r.width > 0 && r.height > 0
        })
        const panel = panels[0]
        if (!panel) return null
        // read the same item the text assertion above targeted
        const item = Array.from(panel.querySelectorAll('button')).find((b) =>
          (b.textContent ?? '').includes('Number Format'),
        )
        if (!item) return null
        const label = item.querySelector('span')
        return {
          panelBg: getComputedStyle(panel).backgroundColor,
          labelColor: getComputedStyle(label ?? item).color,
          labelText: (item.textContent ?? '').trim(),
        }
      })
      expect(probe).not.toBeNull()
      expect(probe?.labelText.length ?? 0).toBeGreaterThan(0)
      // the menu surface is light (document theme) …
      expect(luminance(probe!.panelBg)).toBeGreaterThan(180)
      // … and the label reads dark ink on it, not the UA dark-scheme white
      expect(luminance(probe!.labelColor)).toBeLessThan(120)
      await editorPage.screenshot({ path: screenshotPath('theme-sheets-menu-mixed') })
    } finally {
      await closeAndSaveVideo(launched, 'theme-visual-sheets-menu')
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
      const editorPage = await waitForPageWithUrl(launched.app, '://pdf/')
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
