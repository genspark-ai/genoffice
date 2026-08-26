import { test, expect } from '@playwright/test'
import { execFileSync } from 'node:child_process'
import { mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * The fixture deck (one slide whose only run uses the catalog font Rubik) is
 * kept as plain-text OOXML parts under assets/font-manager-rubik/ so no binary
 * lives in the repo; zip them into a real .pptx at test time.
 */
async function buildRubikFixture(): Promise<string> {
  const out = join(
    await mkdtemp(join(tmpdir(), 'genoffice-font-manager-')),
    'font-manager-rubik.pptx',
  )
  execFileSync('zip', ['-X', '-q', '-r', out, '.'], {
    cwd: resolve(__dirname, 'assets/font-manager-rubik'),
  })
  return out
}

/**
 * Font manager smoke: source builds without an injected font CDN hide download
 * prompts. Providing GENOFFICE_FONT_CDN_URL enables the catalog/banner; the
 * actual network download runs only with E2E_FONT_CDN=1.
 */
test('font download UI follows CDN configuration', async () => {
  test.setTimeout(180_000)
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'slides-font-manager',
    openFile: await buildRubikFixture(),
  })
  try {
    const editorPage = await waitForPageWithUrl(launched.app, 'slides/out')
    await editorPage.waitForSelector('.stage-wrap canvas', { timeout: 30_000 })

    const banner = editorPage.locator('.font-missing-banner')
    const fontState = await editorPage.evaluate(async () => {
      const api = (
        window as unknown as {
          slidesApi: {
            fontCatalog(): Promise<Array<{ family: string; installed: boolean }>>
            fontMissing(): Promise<string[]>
          }
        }
      ).slidesApi
      return {
        catalog: await api.fontCatalog(),
        missing: await api.fontMissing(),
      }
    })
    const configured = Boolean(process.env.GENOFFICE_FONT_CDN_URL?.trim())
    if (!configured) {
      expect(fontState.catalog).toEqual([])
      expect(fontState.missing).toEqual([])
      await expect(banner).toHaveCount(0)
      return
    }

    expect(fontState.catalog.some((c) => c.family === 'Rubik' && !c.installed)).toBe(true)
    expect(fontState.missing).toContain('Rubik')
    await expect(banner).toBeVisible({ timeout: 20_000 })
    await expect(banner).toContainText('Rubik')

    if (process.env.E2E_FONT_CDN === '1') {
      await banner.locator('.fmb-download').click()
      await expect(banner).toBeHidden({ timeout: 60_000 })
      const installed = await editorPage.evaluate(async () => {
        const api = (
          window as unknown as {
            slidesApi: {
              fontCatalog(): Promise<Array<{ family: string; installed: boolean }>>
            }
          }
        ).slidesApi
        const cat = await api.fontCatalog()
        return cat.find((c) => c.family === 'Rubik')?.installed
      })
      expect(installed).toBe(true)
    }
  } finally {
    await closeAndSaveVideo(launched, 'slides-font-manager')
  }
})
