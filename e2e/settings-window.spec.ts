/**
 * Drives the standalone settings window: every section renders, the network
 * and rules sections persist to disk, and a skill written in the editor comes
 * back with the scope it was given.
 */
import { test, expect } from '@playwright/test'
import { readFile, readdir } from 'node:fs/promises'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, screenshotPath, waitForPageWithUrl } from './helpers'

test.describe('settings window', () => {
  test('configures network, rules and skills, and persists them', async () => {
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'settings-window' })
    const { app, page, userDataDir } = launched
    try {
      await page.locator('.account-btn').click()
      await page.locator('.account-menu .lang-row', { hasText: 'Settings' }).click()

      const settings = await waitForPageWithUrl(app, 'settings.html')
      await expect(settings.locator('.settings-nav-item')).toHaveCount(5)

      // ── Model: the Kimi recommendation is the section's last line ──
      await expect(settings.locator('.settings-note')).toHaveText(
        'GenOffice works best with Kimi K3.',
      )
      await settings.screenshot({ path: screenshotPath('settings-model') })

      // ── Network: Tavily key + proxy ──
      await settings.locator('.settings-nav-item', { hasText: 'Network & Search' }).click()
      const netInputs = settings.locator('.ai-input')
      await netInputs.nth(0).fill('tvly-test-key')
      await netInputs.nth(1).fill('127.0.0.1:7890')
      // the guidance for BYOK / mainland-China users has to be visible here
      await expect(settings.locator('.ai-field-hint').first()).toContainText('BYOK')
      await settings.locator('.btn-primary').click()
      await expect(settings.locator('.settings-saved')).toBeVisible()
      await settings.screenshot({ path: screenshotPath('settings-network') })

      const aiFile = JSON.parse(await readFile(join(userDataDir, 'ai-settings.json'), 'utf-8')) as {
        tavilyApiKey: string
        proxyUrl: string
      }
      expect(aiFile.tavilyApiKey).toBe('tvly-test-key')
      // a bare host:port is normalized to a URL undici can actually use
      expect(aiFile.proxyUrl).toBe('http://127.0.0.1:7890')

      // ── Rules: global + per-app ──
      await settings.locator('.settings-nav-item', { hasText: 'Rules' }).click()
      const areas = settings.locator('.settings-textarea')
      await expect(areas).toHaveCount(5)
      await areas.nth(0).fill('Never invent figures.')
      await areas.nth(1).fill('Use Heading 1 for titles.')
      await settings.locator('.btn-primary').click()
      await expect(settings.locator('.settings-saved')).toBeVisible()
      await settings.screenshot({ path: screenshotPath('settings-rules') })

      const rules = JSON.parse(
        await readFile(join(userDataDir, 'agent-rules.json'), 'utf-8'),
      ) as Record<string, string>
      expect(rules.global).toBe('Never invent figures.')
      expect(rules.docx).toBe('Use Heading 1 for titles.')

      // ── Skills: write one, scope it to slides only ──
      await settings.locator('.settings-nav-item', { hasText: 'Skills' }).click()
      await expect(settings.locator('.settings-empty')).toBeVisible()
      await settings.locator('.btn', { hasText: 'New skill' }).click()

      const skillInputs = settings.locator('.ai-input')
      await skillInputs.nth(0).fill('Deck review')
      await skillInputs.nth(1).fill('Check a deck before export')
      // "All apps" is the default; picking a specific app replaces it, because
      // global already covers everything
      await settings
        .locator('.settings-scope-option', { hasText: 'Slides' })
        .locator('input')
        .check()
      await expect(
        settings.locator('.settings-scope-option', { hasText: 'All apps' }).locator('input'),
      ).not.toBeChecked()
      await settings.locator('.settings-textarea').fill('## Steps\n1. Check contrast.')
      await settings.locator('.btn-primary', { hasText: 'Save' }).click()

      const card = settings.locator('.settings-skill')
      await expect(card).toHaveCount(1)
      await expect(card).toContainText('Deck review')
      await expect(card.locator('.settings-chip')).toHaveText('Slides (.pptx)')
      await settings.screenshot({ path: screenshotPath('settings-skills') })

      // stored as a real markdown file with front matter, editable outside the app
      const skillDir = join(userDataDir, 'agent-skills')
      const files = await readdir(skillDir)
      expect(files).toHaveLength(1)
      const md = await readFile(join(skillDir, files[0]!), 'utf-8')
      expect(md).toContain('name: Deck review')
      expect(md).toContain('scopes: pptx')
      expect(md).toContain('1. Check contrast.')

      // ── the agent prompt advertises the title but not the body ──
      const prompt = await app.evaluate(async ({ ipcMain }) => {
        type Handler = (event: unknown, ...args: unknown[]) => unknown
        const invoke = (channel: string, ...args: unknown[]) =>
          (
            (ipcMain as unknown as { _invokeHandlers: Map<string, Handler> })._invokeHandlers.get(
              channel,
            ) as Handler
          )({}, ...args) as Promise<string>
        return {
          pptx: await invoke('ai:instructions-prompt', 'pptx'),
          sheets: await invoke('ai:instructions-prompt', 'sheets'),
          // the load_skill tool honours scope too
          bodyInScope: await invoke('ai:skill-body', 'pptx', 'deck-review'),
          bodyOutOfScope: await invoke('ai:skill-body', 'sheets', 'deck-review'),
        }
      })
      expect(prompt.bodyInScope).toContain('Check contrast')
      expect(prompt.bodyOutOfScope).toContain('No skill')
      // slides sees the skill; sheets does not, because of the scope
      expect(prompt.pptx).toContain('Deck review')
      expect(prompt.pptx).not.toContain('Check contrast')
      expect(prompt.sheets).not.toContain('Deck review')
      // global rules reach both, the docx rule reaches neither
      expect(prompt.pptx).toContain('Never invent figures.')
      expect(prompt.sheets).toContain('Never invent figures.')
      expect(prompt.sheets).not.toContain('Heading 1')
    } finally {
      await closeAndSaveVideo(launched, 'settings-window')
    }
  })
})
