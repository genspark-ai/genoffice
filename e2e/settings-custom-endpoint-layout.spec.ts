import { test, expect } from '@playwright/test'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { defaultAiSettings, resolveCustomEndpoints } from '@genoffice/ai-provider'
import { closeAndSaveVideo, launchShell, screenshotPath } from './helpers'

const locales = [
  { lang: 'en', settings: 'Settings', pane: 'AI Model' },
  { lang: 'zh', settings: '\u8bbe\u7f6e', pane: 'AI \u6a21\u578b' },
  { lang: 'de', settings: 'Einstellungen', pane: 'KI-Modell' },
] as const

for (const locale of locales) {
  test(`custom endpoint fields fit both regular and minimum windows in ${locale.lang}`, async () => {
    const userDataDir = await mkdtemp(join(tmpdir(), 'genoffice-endpoint-layout-'))
    const settings = defaultAiSettings()
    settings.provider = 'custom'
    settings.providers.custom = {
      apiKey: '',
      model: 'local-model',
      baseUrl: 'http://127.0.0.1:1/v1',
      profile: {
        api: 'anthropic',
        headers: { 'X-Tenant': 'office' },
        modelProfiles: [{ id: 'local-model', maxTokens: 8192 }],
      },
    }
    await writeFile(
      join(userDataDir, 'ai-settings.json'),
      JSON.stringify(resolveCustomEndpoints(settings)),
    )
    const launched = await launchShell({
      userDataDir,
      lang: locale.lang,
      onboardingSeen: true,
      settings: { starPrompt: { resolved: true } },
      videoDir: 'settings-endpoint-layout',
    })
    try {
      const { page } = launched
      await page.getByRole('button', { name: locale.settings, exact: true }).click()
      await page.locator('.set-nav-item').filter({ hasText: locale.pane }).click()
      await expect(page.locator('.set-ai-profile-switches')).toBeVisible()
      for (const size of [
        { width: 1280, height: 800 },
        { width: 720, height: 550 },
      ]) {
        await launched.app.evaluate(({ BrowserWindow }, { width, height }) => {
          BrowserWindow.getAllWindows()[0]!.setSize(width, height)
        }, size)
        await expect(page.locator('.set-pane-ai-model')).toBeVisible()
        // Poll through window resize/layout rather than taking a measurement
        // while Electron is still applying the requested size.
        await expect.poll(() => page.evaluate(() => innerWidth)).toBe(size.width)
        const overflow = await page.locator('.set-pane-ai-model').evaluate((pane) => {
          const selectors =
            '.set-field, .set-field-text, .set-ai-profile-row, .set-ai-profile-label'
          const elements = [pane, ...pane.querySelectorAll(selectors)]
          return elements
            .filter((element) => element.scrollWidth > element.clientWidth + 1)
            .map((element) => ({
              className: element.className,
              text: element.textContent?.trim().slice(0, 80),
              width: element.clientWidth,
              scrollWidth: element.scrollWidth,
            }))
        })
        expect(overflow).toEqual([])
        const overlap = await page.locator('.set-ai-profile-row').evaluateAll(
          (rows) =>
            rows.filter((row) => {
              const label = row.firstElementChild!.getBoundingClientRect()
              const control = row.lastElementChild!.getBoundingClientRect()
              return label.right > control.left + 1
            }).length,
        )
        expect(overlap).toBe(0)
        for (const section of ['switches', 'models'] as const) {
          const target =
            section === 'switches'
              ? page.locator('.set-ai-profile-switches')
              : page.locator('#set-ai-endpoint-models')
          await target.scrollIntoViewIfNeeded()
          await page.screenshot({
            path: screenshotPath(`endpoint-layout-${locale.lang}-${size.width}-${section}`),
          })
        }
      }
    } finally {
      await closeAndSaveVideo(launched, `settings-endpoint-layout-${locale.lang}`)
    }
  })
}
