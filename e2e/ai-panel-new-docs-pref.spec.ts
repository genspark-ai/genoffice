import { test, expect } from '@playwright/test'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * Settings → General → "Open the AI panel in new documents" (#1589).
 *
 * The setting is read once per renderer, before the first render, and the panel
 * state is a `useState` initialiser — so the only moment the setting can be
 * honoured is that first read. Both directions are asserted: a test that only
 * checked the collapsed case would also pass if the panel never opened at all.
 */

const FIXTURE = join(__dirname, 'assets/justify-pagegap-fr.docx')

/** collapsed = the rail only; expanded = the composer is on screen */
async function aiPanelState(
  editor: import('@playwright/test').Page,
): Promise<'collapsed' | 'expanded'> {
  return (await editor.locator('.ai-dock.collapsed').count()) > 0 ? 'collapsed' : 'expanded'
}

test.describe('the "open the AI panel in new documents" setting', () => {
  test('a new document opens with the panel collapsed when the setting is off', async () => {
    test.setTimeout(180_000)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'ai-panel-new-docs',
      openFile: FIXTURE,
      // the exact state the reporter's app-settings.json held
      settings: { aiPanelOpenInNewDocs: false },
    })
    try {
      const editor = await waitForPageWithUrl(launched.app, '://docs/')
      await expect(editor.locator('.ProseMirror').first()).toBeVisible({ timeout: 30_000 })
      await expect(editor.locator('.ai-dock')).toHaveClass(/collapsed/)
      await expect(editor.locator('.ai-composer textarea')).toBeHidden()
    } finally {
      await closeAndSaveVideo(launched, 'ai-panel-new-docs')
    }
  })

  test('a new document opens with the panel expanded when the setting is on', async () => {
    test.setTimeout(180_000)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'ai-panel-new-docs-on',
      openFile: FIXTURE,
      settings: { aiPanelOpenInNewDocs: true },
    })
    try {
      const editor = await waitForPageWithUrl(launched.app, '://docs/')
      await expect(editor.locator('.ProseMirror').first()).toBeVisible({ timeout: 30_000 })
      await expect(editor.locator('.ai-dock')).not.toHaveClass(/collapsed/)
      await expect(editor.locator('.ai-composer textarea')).toBeVisible()
    } finally {
      await closeAndSaveVideo(launched, 'ai-panel-new-docs-on')
    }
  })

  test('the setting is read from app-settings.json, not left at its default', async () => {
    // the failure mode this guards is silent: a rejected getAiPanelPrefs leaves
    // the defaults in place, and the default for this setting is "on", so the
    // panel opens and nothing says why
    test.setTimeout(180_000)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'ai-panel-new-docs-side',
      openFile: FIXTURE,
      settings: { aiPanelOpenInNewDocs: false, aiPanelSide: 'right' },
    })
    try {
      const editor = await waitForPageWithUrl(launched.app, '://docs/')
      await expect(editor.locator('.ProseMirror').first()).toBeVisible({ timeout: 30_000 })
      // proves the prefs round-tripped: the side came from the same payload
      await expect(editor.locator('html')).toHaveAttribute('data-ai-panel-side', 'right')
      expect(await aiPanelState(editor)).toBe('collapsed')
    } finally {
      await closeAndSaveVideo(launched, 'ai-panel-new-docs-side')
    }
  })
})
