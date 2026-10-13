import { test, expect } from '@playwright/test'
import { writeFileSync, mkdirSync } from 'node:fs'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/** The theme picker persists its choice, re-scopes the paper via
 *  data-markdown-theme, and survives a restart. */

async function awaitDocsEditor(page: import('@playwright/test').Page): Promise<void> {
  const editor = page.locator('.ProseMirror').first()
  try {
    await editor.waitFor({ timeout: 45_000 })
    return
  } catch {
    await page.reload({ waitUntil: 'domcontentloaded' })
    await editor.waitFor({ timeout: 90_000 })
  }
}

test('markdown theme picker applies, persists and restarts', async () => {
  test.setTimeout(300_000)
  const dir = join(__dirname, 'artifacts')
  mkdirSync(dir, { recursive: true })
  const docPath = join(dir, 'theme-probe.md')
  writeFileSync(
    docPath,
    '# Theme probe\n\n\u4e00\u6bb5\u4e2d\u6587\u6b63\u6587\uff0c\u7528\u4e8e\u89c2\u5bdf\u7eb8\u9762\u4e0e\u58a8\u8272\u3002\n\n```js\nconst x = 1\n```\n',
  )
  const launched = await launchShell({
    onboardingSeen: true,
    openFile: docPath,
    videoDir: 'markdown-theme-picker',
  })
  const { app } = launched
  let closed = false
  try {
    const page = await waitForPageWithUrl(app, '://markdown/')
    await awaitDocsEditor(page)

    const themeButton = page.locator('.md-theme-wrap > .rb-btn').first()
    await themeButton.click()
    const menu = page.locator('.md-theme-menu')
    await expect(menu).toBeVisible()
    // every theme is listed
    expect(await menu.locator('.md-theme-item').count()).toBe(7)

    await menu.locator('.md-theme-item', { hasText: 'Nord' }).click()
    await expect(menu).toBeHidden()
    await page.waitForTimeout(600)

    const applied = await page.evaluate(() => ({
      attr: document.documentElement.getAttribute('data-markdown-theme'),
      paper: getComputedStyle(document.querySelector('.doc-page')!).backgroundColor,
      canvas: getComputedStyle(document.querySelector('.editor-scroll')!).backgroundColor,
    }))
    expect(applied.attr).toBe('nord')
    expect(applied.paper).toBe('rgb(46, 52, 64)') // nord0 #2e3440
    await page.screenshot({ path: 'artifacts/md-theme-nord.png' })

    // a real restart: close this instance, then boot again on the same
    // userDataDir (a second live instance would just be forwarded to this one)
    closed = true
    await closeAndSaveVideo(launched, 'markdown-theme-picker')
    const relaunch = await launchShell({
      onboardingSeen: true,
      openFile: docPath,
      videoDir: 'markdown-theme-picker',
      userDataDir: launched.userDataDir,
    })
    try {
      const page2 = await waitForPageWithUrl(relaunch.app, '://markdown/')
      await awaitDocsEditor(page2)
      await page2.waitForTimeout(1000)
      const attr2 = await page2.evaluate(() =>
        document.documentElement.getAttribute('data-markdown-theme'),
      )
      expect(attr2).toBe('nord')
      await page2.screenshot({ path: 'artifacts/md-theme-nord-restart.png' })
    } finally {
      await closeAndSaveVideo(relaunch, 'markdown-theme-picker')
    }
  } finally {
    if (!closed) await closeAndSaveVideo(launched, 'markdown-theme-picker')
  }
})
