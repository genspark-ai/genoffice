import { test, expect } from '@playwright/test'
import { mkdirSync, writeFileSync } from 'node:fs'
import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo } from './helpers'

function seedFakeGgufAndRuntime(userDataDir: string): void {
  const modelsDir = join(userDataDir, 'models', 'lfm3-vl-2b')
  const runtimeDir = join(userDataDir, 'runtimes', 'llama')
  mkdirSync(modelsDir, { recursive: true })
  mkdirSync(runtimeDir, { recursive: true })
  const gguf = Buffer.alloc(64)
  Buffer.from('GGUF', 'ascii').copy(gguf, 0)
  writeFileSync(join(modelsDir, 'LFM3-VL-2B-Q4_K_M.gguf'), gguf)
  writeFileSync(join(modelsDir, 'mmproj-LFM3-VL-2B-Q8_0.gguf'), gguf)
  const exe = process.platform === 'win32' ? 'llama-server.exe' : 'llama-server'
  writeFileSync(join(runtimeDir, exe), 'stub')
}

test.describe('Local AI (LFM3-VL-2B Q4_K_M) settings & OCR fallback', () => {
  test('AI Model, AI Media, and OCR escalation settings expose LFM3-VL-2B controls and persist choices', async () => {
    const launched = await launchShell({
      onboardingSeen: true,
      settings: { starPrompt: { resolved: true } },
      videoDir: 'local-vlm-settings',
    })
    const { page, userDataDir } = launched
    try {
      await page.getByRole('button', { name: 'Settings', exact: true }).click()

      // 1. AI Model pane: select Local AI (LFM3-VL-2B)
      await page.locator('.set-nav-item').filter({ hasText: 'AI Model' }).click()
      await page.getByRole('button', { name: 'Provider', exact: true }).click()
      await page.getByRole('option', { name: 'Local AI (LFM3-VL-2B)', exact: true }).click()

      await expect(page.getByRole('button', { name: 'Download LFM3-VL-2B (Q4_K_M)' })).toBeVisible()
      const fallbackToggle = page.getByRole('switch', {
        name: 'Auto-fallback when Antigravity CLI is unavailable',
      })
      await expect(fallbackToggle).toHaveAttribute('aria-checked', 'true')
      await expect(page.locator('#set-local-vlm-model-path')).toBeVisible()
      await expect(page.locator('#set-local-vlm-mmproj-path')).toBeVisible()
      await expect(page.locator('#set-local-vlm-runtime-path')).toBeVisible()

      // Save Local AI as the chat provider
      await page.getByRole('button', { name: 'Save', exact: true }).click()
      await expect
        .poll(async () => {
          const saved = JSON.parse(await readFile(join(userDataDir, 'ai-settings.json'), 'utf8'))
          return saved.provider
        })
        .toBe('local-vlm')

      // 2. AI Media & Search pane: select Local AI (LFM3-VL-2B) for Image analysis
      await page.locator('.set-nav-item').filter({ hasText: 'AI Media & Search' }).click()
      await page.getByRole('button', { name: 'Image analysis', exact: true }).click()
      await page.getByRole('option', { name: 'Local AI (LFM3-VL-2B)', exact: true }).click()
      await expect(page.locator('.set-local-vlm')).toBeVisible()

      // 3. General pane -> OCR escalation backend selector
      await page.locator('.set-nav-item').filter({ hasText: 'General' }).click()
      const ocrSelect = page.getByLabel('OCR escalation backend')
      await expect(ocrSelect).toBeVisible()
      await ocrSelect.selectOption('local-vlm')
      await expect
        .poll(async () => {
          const saved = JSON.parse(await readFile(join(userDataDir, 'app-settings.json'), 'utf8'))
          return saved.agyOcr?.escalationBackend
        })
        .toBe('local-vlm')

      // 4. Seed valid GGUF + runtime files and verify IPC reports ready & Start server appears
      seedFakeGgufAndRuntime(userDataDir)
      await page.locator('.set-nav-item').filter({ hasText: 'AI Model' }).click()
      const status = await page.evaluate(() => window.aiOffice.getLocalVlmStatus?.())
      expect(status?.installed).toBe(true)
      expect(status?.runtimeInstalled).toBe(true)
    } finally {
      await closeAndSaveVideo(launched, 'local-vlm-settings')
    }
  })
})
