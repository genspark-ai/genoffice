import { test, expect } from '@playwright/test'
import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * Settings → General → "Open the AI panel in new documents" (genoffice#1589), the way the
 * reporter hit it: flip the toggle in the UI, restart, then start a *new*
 * document from the Home screen. Pre-seeding app-settings.json is not the same
 * thing — it skips the write, and the write is half of what is being reported.
 */

/** collapsed = the rail only; expanded = the composer is on screen */
async function aiPanelState(
  editor: import('@playwright/test').Page,
): Promise<'collapsed' | 'expanded'> {
  return (await editor.locator('.ai-dock.collapsed').count()) > 0 ? 'collapsed' : 'expanded'
}

async function turnOffOpenInNewDocs(page: import('@playwright/test').Page): Promise<void> {
  await page.getByRole('button', { name: 'Settings', exact: true }).click()
  await page.locator('.set-nav-item').filter({ hasText: 'General' }).click()
  // role=switch, not an implicit button
  const toggle = page.getByRole('switch', { name: 'Open the AI panel in new documents' })
  if ((await toggle.getAttribute('aria-checked')) === 'true') await toggle.click()
  await expect(toggle).toHaveAttribute('aria-checked', 'false')
  await page.locator('.set-close').click()
}

test.describe('#1589 the "open the AI panel in new documents" setting', () => {
  test('a new document from the Home screen opens collapsed after the toggle is turned off', async () => {
    test.setTimeout(300_000)
    const first = await launchShell({ onboardingSeen: true, videoDir: 'ai-panel-1589-on' })
    let userDataDir: string
    try {
      await turnOffOpenInNewDocs(first.page)
      // the write half: the flag has to be on disk, not only in the shell's cache
      const saved = JSON.parse(await readFile(join(first.userDataDir, 'app-settings.json'), 'utf8'))
      expect(saved.aiPanelOpenInNewDocs).toBe(false)
      userDataDir = first.userDataDir
    } finally {
      await closeAndSaveVideo(first, 'ai-panel-1589-on')
    }

    const restarted = await launchShell({ userDataDir, videoDir: 'ai-panel-1589-new-doc' })
    try {
      await expect(restarted.page.locator('.quick-card').first()).toContainText('AI Docs')
      await restarted.page.locator('.quick-card').first().click()
      const editor = await waitForPageWithUrl(restarted.app, '://docs/')
      await expect(editor.locator('.ProseMirror').first()).toBeVisible({ timeout: 30_000 })
      expect(await aiPanelState(editor)).toBe('collapsed')
      await expect(editor.locator('.ai-composer textarea')).toBeHidden()
    } finally {
      await closeAndSaveVideo(restarted, 'ai-panel-1589-new-doc')
    }
  })

  test('the same new document still opens with the panel when the setting is left on', async () => {
    // The other half, and the reason the fix is not `setShowAi(false)`: the
    // opening document and a *new* document have to agree. This one goes
    // through newFile too — a test that opened an existing file instead would
    // never run the line under test and would pass either way.
    test.setTimeout(300_000)
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'ai-panel-1589-on-new' })
    try {
      await expect(launched.page.locator('.quick-card').first()).toContainText('AI Docs')
      await launched.page.locator('.quick-card').first().click()
      const editor = await waitForPageWithUrl(launched.app, '://docs/')
      await expect(editor.locator('.ProseMirror').first()).toBeVisible({ timeout: 30_000 })
      expect(await aiPanelState(editor)).toBe('expanded')
      await expect(editor.locator('.ai-composer textarea')).toBeVisible()
    } finally {
      await closeAndSaveVideo(launched, 'ai-panel-1589-on-new')
    }
  })
})
