import { test, expect } from '@playwright/test'
import { execSync } from 'node:child_process'
import { copyFile, mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import type { Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, screenshotPath } from './helpers'

const FIXTURE = resolve(__dirname, '../apps/sheets/fixtures/generated/compatibility-basic.xlsx')

async function waitForWorkbook(page: Page): Promise<void> {
  await page.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
    timeout: 30_000,
  })
  await page.waitForTimeout(1_500)
}

test.describe('sheets: Insert → Screenshot', () => {
  test('captures a screen source and saves it into the workbook', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-screenshot-e2e-'))
    const workbook = join(scratch, 'screenshot.xlsx')
    await copyFile(FIXTURE, workbook)

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-insert-screenshot',
      openFile: workbook,
    })
    try {
      const sheets = await waitForPageWithUrl(launched.app, 'sheets/out')
      await waitForWorkbook(sheets)

      // Always stub the OS capturer in the main process: real capture is
      // host-dependent (macOS gates it behind the Screen Recording
      // permission, and on headless CI xvfb intermittently advertises a
      // screen source whose full-resolution grab comes back as an empty
      // frame, failing the capture after the picker showed it). The test's
      // job is the picker → capture → insert → save pipeline, which runs for
      // real either way, so make the source deterministic everywhere.
      await launched.app.evaluate(({ desktopCapturer, nativeImage, systemPreferences }) => {
        const pixel =
          'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
        const frame = nativeImage.createFromDataURL(pixel).resize({ width: 800, height: 600 })
        desktopCapturer.getSources = async () => [
          { id: 'screen:0:0', name: 'Mock Screen', thumbnail: frame } as never,
        ]
        systemPreferences.getMediaAccessStatus = () => 'granted'
      })

      await sheets.getByRole('button', { name: 'Insert', exact: true }).click()
      await sheets.getByRole('button', { name: 'Screenshot' }).click()
      const dialog = sheets.getByRole('dialog', { name: 'Screenshot' })
      await expect(dialog).toBeVisible()

      const grid = dialog.locator('.screenshot-grid button')
      await expect(grid.first()).toBeVisible()
      await sheets.screenshot({ path: screenshotPath('screenshot-picker') })
      await grid.first().click()

      await expect(sheets.getByRole('status')).toContainText('Picture inserted', {
        timeout: 20_000,
      })
      await expect(dialog).not.toBeVisible()
      await expect(sheets.locator('.xlsx-image')).toBeVisible({ timeout: 15_000 })
      await sheets.screenshot({ path: screenshotPath('screenshot-inserted') })

      await launched.app.evaluate(({ webContents }) => {
        const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('sheets/out'))
        wc?.send('menu:action', 'save')
      })
      await expect(() => {
        const listing = execSync(`unzip -l "${workbook}"`).toString()
        expect(listing).toContain('xl/media/')
      }).toPass({ timeout: 20_000 })
    } finally {
      await closeAndSaveVideo(launched, 'sheets-insert-screenshot')
    }
  })
})
