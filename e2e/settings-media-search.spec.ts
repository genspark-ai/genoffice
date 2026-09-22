import { test, expect } from '@playwright/test'
import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, screenshotPath } from './helpers'

test('Jev reranking lives in the AI Media & Search pane, saves with it, and reports its own test verdict', async () => {
  // the star prompt card sits over the pane footer on a small window
  const launched = await launchShell({
    onboardingSeen: true,
    settings: { starPrompt: { resolved: true } },
    videoDir: 'settings-media-search',
  })
  const { page, userDataDir } = launched
  try {
    await page.getByRole('button', { name: 'Settings', exact: true }).click()
    await page.locator('.set-nav-item').filter({ hasText: 'General' }).click()
    await expect(page.getByRole('switch', { name: 'Jev search reranking' })).toHaveCount(0)

    await page.locator('.set-nav-item').filter({ hasText: 'AI Media & Search' }).click()
    const block = page.locator('section', {
      has: page.locator('.set-pane-subtitle', { hasText: 'Local file search' }),
    })
    const toggle = block.getByRole('switch', { name: 'Jev search reranking' })
    await expect(toggle).toHaveAttribute('aria-checked', 'false')
    await toggle.click()
    await block.getByRole('button', { name: 'Jev endpoint', exact: true }).click()
    await page.getByRole('option', { name: 'TypeSafe', exact: true }).click()

    // the block reports its own verdict: no key entered, nothing leaves the machine
    await page.getByRole('button', { name: 'Test connection', exact: true }).click()
    await expect(block.locator('.set-ai-status.err')).toHaveText('Enter an API key')
    // the footer names the first failing block and its provider; which one comes
    // first depends on whether this machine is signed in to Genspark
    await expect(page.locator('.set-pane-actions .set-ai-status.err')).toHaveText(
      /^(Web search · Genspark|Local file search · TypeSafe): .+/,
    )
    await page.screenshot({ path: screenshotPath('settings-media-search-test') })

    await block.locator('#set-search-jev-key').fill('ts-key')
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    await expect(page.locator('.set-pane-actions .set-ai-status.ok')).toHaveText('Saved')
    await expect
      .poll(async () => {
        const saved = JSON.parse(await readFile(join(userDataDir, 'app-settings.json'), 'utf8'))
        return saved.fileSearch
      })
      .toEqual({
        rerank: true,
        jevEndpoint: 'direct',
        jevKeys: { openrouter: '', direct: 'ts-key' },
      })
  } finally {
    await closeAndSaveVideo(launched, 'settings-media-search')
  }
})
