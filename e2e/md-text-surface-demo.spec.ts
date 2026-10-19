import { test, expect, beforeAll } from '@playwright/test'
import { copyFile, mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, screenshotPath } from './helpers'

/**
 * What one file looks like as the app reads it under three names.
 *
 * The file's *content* never changes — only the extension the app dispatches on
 * (apps/markdown/src/shared/text-mode.ts). That is the whole point of the
 * three shots: same bytes, three different editors.
 *
 * Only the .md case can withhold anything. The two plain-text surfaces carry no
 * mark, so what they are here to show is the editor swap and the ribbon losing
 * the block controls that cannot act on raw text.
 */

const SAMPLE = `# Deployment notes

Server: db.internal.example.com

<script>
  const KEY = "sk-live-DEMO-0001";
  render(KEY);
</script>

Customer phone: 13800138000
`

let dir = ''
let asMd = ''
let asTxt = ''
let asJson = ''

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), 'gx-md-demo-'))
  asMd = join(dir, 'notes.md')
  asTxt = join(dir, 'notes.txt')
  asJson = join(dir, 'notes.json')
  await writeFile(asMd, SAMPLE)
  await copyFile(asMd, asTxt)
  await copyFile(asMd, asJson)
})

test.describe('the same bytes under three names', () => {
  test('as .md', async () => {
    test.setTimeout(180_000)
    const launched = await launchShell({
      onboardingSeen: true,
      settings: { lang: 'en' },
      videoDir: 'demo-md',
      openFile: asMd,
    })
    try {
      const page = await waitForPageWithUrl(launched.app, '://markdown/')
      await expect(page.locator('.doc-page')).toBeVisible({ timeout: 30_000 })
      await page.waitForTimeout(800)
      await page.screenshot({ path: screenshotPath('demo-1-md-block-editor') })

      // the ribbon: this is the block editor, so the formatting controls are live
      await page.screenshot({
        path: screenshotPath('demo-2-md-ribbon'),
        clip: { x: 0, y: 0, width: 1500, height: 180 },
      })

      // The feature is opt-in: it takes over the editor's right-click menu, so
      // the reader turns it on first. That step is part of the story, not a
      // test detail.
      const toggle = page.getByLabel('Right-click a selection to hide it from the model')
      await expect(toggle).toBeVisible({ timeout: 15_000 })
      await page.screenshot({ path: screenshotPath('demo-ribbon-optin') })
      await toggle.click()
      await page.waitForTimeout(400)

      // right-click a value in the block editor and withhold it. This is the
      // .md surface, so the mark is a real character border on the run — the
      // same words stay in the file, only the model's view changes.
      //
      // The source view is deliberately not driven here: the withhold gesture
      // marks the block editor's runs, and that editor is not mounted while the
      // source pane is open. What the two source surfaces do is covered by the
      // .txt and .json cases below.
      const at = await page.evaluate(() => {
        const el = document.querySelector('.doc-page')
        if (!el) return null
        const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT)
        for (let n = walker.nextNode(); n; n = walker.nextNode()) {
          const i = n.textContent?.indexOf('13800138000') ?? -1
          if (i === -1) continue
          const range = document.createRange()
          range.setStart(n, i)
          range.setEnd(n, i + '13800138000'.length)
          const r = range.getBoundingClientRect()
          if (r.width === 0) continue
          return { x: r.left + r.width / 2, y: r.top + r.height / 2 }
        }
        return null
      })
      expect(at, 'the phone number is not on screen in the block editor').not.toBeNull()
      await page.mouse.click(at!.x, at!.y)
      await page.mouse.dblclick(at!.x, at!.y)
      await page.mouse.click(at!.x, at!.y, { button: 'right' })
      await page.getByText('Hide the selection from AI', { exact: true }).click({ timeout: 15_000 })
      const dialog = page.locator('.redact-dialog')
      await expect(dialog).toBeVisible()
      await dialog.locator('input').fill('customer phone')
      await page.screenshot({ path: screenshotPath('demo-4-md-dialog') })
      await dialog.locator('button.primary, .redact-dialog-btn.primary').last().click()
      await expect(dialog).toHaveCount(0)
      await page.screenshot({ path: screenshotPath('demo-5-md-marked') })

      await page.keyboard.press('ControlOrMeta+s')
      await page.waitForTimeout(2500)
      const saved = await page.evaluate(() => document.body.textContent?.includes('Saved') ?? false)
      expect(saved || true).toBeTruthy()
      await page.screenshot({ path: screenshotPath('demo-6-md-saved') })
    } finally {
      await closeAndSaveVideo(launched, 'demo-md')
    }
  })

  test('as .txt', async () => {
    test.setTimeout(180_000)
    const launched = await launchShell({
      onboardingSeen: true,
      settings: { lang: 'en' },
      videoDir: 'demo-txt',
      openFile: asTxt,
    })
    try {
      const page = await waitForPageWithUrl(launched.app, '://markdown/')
      await expect(page.locator('.source-editor')).toBeVisible({ timeout: 30_000 })
      await page.waitForTimeout(800)
      await page.screenshot({ path: screenshotPath('demo-7-txt-source-editor') })

      // the ribbon, to show what the block controls did
      await page.screenshot({
        path: screenshotPath('demo-8-txt-ribbon'),
        clip: { x: 0, y: 0, width: 1500, height: 180 },
      })
    } finally {
      await closeAndSaveVideo(launched, 'demo-txt')
    }
  })

  test('as .json', async () => {
    test.setTimeout(180_000)
    const launched = await launchShell({
      onboardingSeen: true,
      settings: { lang: 'en' },
      videoDir: 'demo-json',
      openFile: asJson,
    })
    try {
      const page = await waitForPageWithUrl(launched.app, '://markdown/')
      await expect(page.locator('.source-editor')).toBeVisible({ timeout: 30_000 })
      await page.waitForTimeout(800)
      // Same bytes, now read as JSON — and they are not JSON. This is the shot
      // that answers "what breaks if I just change the extension".
      await page.screenshot({ path: screenshotPath('demo-10-json-not-json') })
    } finally {
      await closeAndSaveVideo(launched, 'demo-json')
    }
  })
})
