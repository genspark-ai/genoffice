/**
 * Pixel-regression gate for the docs renderer (self-baseline track).
 *
 * Every corpus document renders in the real built shell and each page is
 * compared 1:1 (zero diff pixels) against a committed PNG under
 * e2e/visual-baselines/ — any rendering drift fails CI. The corpus starts
 * with the cap-table shapes from the floating-table overflow fix (#1111);
 * grow it by adding a fixture to packages/docx-engine/scripts/
 * generate-fixtures.ts and a row to CORPUS below.
 *
 * Baselines are rendered on Linux under xvfb with the fonts CI pins
 * (fonts-crosextra-carlito/caladea for Calibri/Cambria metrics,
 * fonts-noto-cjk); the suite skips on other platforms. After an intentional
 * rendering change, regenerate on Linux:
 *   npm run build:all && npm run fixtures
 *   xvfb-run --auto-servernum -- npm run test:e2e -- docs-visual --update-snapshots
 * If a local render disagrees with CI (font-stack drift between distros),
 * commit the `-actual.png` files from the failed run's e2e-artifacts instead.
 */
import { resolve } from 'node:path'
import { test, expect, type Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

const CORPUS: Array<{ name: string; file: string; pages: number }> = [
  { name: 'kitchen-sink', file: 'kitchen-sink.docx', pages: 1 },
  { name: 'captable-inline', file: 'visual-captable-inline.docx', pages: 1 },
  { name: 'captable-float', file: 'visual-captable-float.docx', pages: 1 },
  { name: 'captable-pct', file: 'visual-captable-pct.docx', pages: 1 },
  { name: 'captable-page-anchor', file: 'visual-captable-page-anchor.docx', pages: 1 },
]

const FIXTURES_DIR = resolve(__dirname, '../fixtures/generated')

/** wait until pagination stops mutating: page count + last-page bottom stable */
async function settledPageCount(page: Page): Promise<number> {
  await page.waitForSelector('.doc-page', { timeout: 30_000 })
  await page.evaluate(() => document.fonts.ready.then(() => undefined))
  const deadline = Date.now() + 30_000
  let prev = ''
  let stableSince = Date.now()
  while (Date.now() < deadline) {
    const cur = await page.evaluate(() => {
      const pages = document.querySelectorAll('.doc-page')
      const last = pages[pages.length - 1]
      const bottom = last ? Math.round(last.getBoundingClientRect().bottom) : 0
      return `${pages.length}:${bottom}`
    })
    if (cur !== prev) {
      prev = cur
      stableSince = Date.now()
    } else if (Date.now() - stableSince > 1_500) {
      return Number(cur.split(':')[0])
    }
    await page.waitForTimeout(200)
  }
  throw new Error(`doc layout never settled (last: ${prev})`)
}

test.describe('docs visual regression', () => {
  test.skip(process.platform !== 'linux', 'pixel baselines are rendered on Linux CI')

  for (const doc of CORPUS) {
    test(`${doc.name} renders pixel-identical pages`, async () => {
      const launched = await launchShell({
        onboardingSeen: true,
        videoDir: `docs-visual-${doc.name}`,
        openFile: resolve(FIXTURES_DIR, doc.file),
      })
      try {
        const editorPage = await waitForPageWithUrl(launched.app, 'docs/out')
        // Keep the default width (narrower clamps zoom below 100% via
        // width-fit), but grow the height until a whole A4 page fits in the
        // viewport: capturing beyond the viewport drops the page border and
        // bakes the horizontal scrollbar into the image. Xvfb allows windows
        // larger than its screen.
        await launched.app.evaluate(({ BrowserWindow }) => {
          const win = BrowserWindow.getAllWindows()[0]
          win?.setBounds({ x: 0, y: 0, width: 1360, height: 1500 })
        })
        const pageCount = await settledPageCount(editorPage)
        expect(pageCount).toBe(doc.pages)
        // fixed chrome (status bar with its async word count) must never
        // overlap a capture — keep baselines document-only
        await editorPage.addStyleTag({
          content: '.status-bar { visibility: hidden !important; }',
        })
        for (let i = 0; i < pageCount; i++) {
          const pageEl = editorPage.locator('.doc-page').nth(i)
          // zoom must be exactly 100% (A4 = 1123 css px) or baselines are moot
          await pageEl.scrollIntoViewIfNeeded()
          expect(Math.round((await pageEl.boundingBox())?.height ?? 0)).toBe(1123)
          await expect(pageEl).toHaveScreenshot(`${doc.name}-p${i + 1}.png`, {
            maxDiffPixels: 0,
            animations: 'disabled',
          })
        }
      } finally {
        await closeAndSaveVideo(launched, `docs-visual-${doc.name}`)
      }
    })
  }
})
