import { test, expect, type ElectronApplication } from '@playwright/test'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * One Settings entry, one accelerator (⌘, on macOS, Ctrl+, elsewhere), in
 * every app's menu — and it lands on the Settings view (genoffice#2037).
 *
 * The macOS application menu is spelled out rather than left as Electron's
 * fixed role:'appMenu' template, so the entry exists on macOS too; the other
 * platforms put it at the bottom of the File menu. Either way the assertion
 * is the same: find the entry, check its accelerator, click it, and see the
 * Settings dialog.
 */

type SettingsReport = {
  found: number
  labels: string[]
  accelerators: string[]
}

const settingsReport = (app: ElectronApplication): Promise<SettingsReport> =>
  app.evaluate(({ Menu }) => {
    const menu = Menu.getApplicationMenu()
    if (!menu) return { found: 0, labels: [], accelerators: [] }
    const items = menu.items.flatMap((i) => i.submenu?.items ?? [])
    // matched by the accelerator: the label is localized, the shortcut is not
    const settings = items.filter((i) => i.accelerator === 'CmdOrCtrl+,')
    return {
      found: settings.length,
      labels: settings.map((i) => i.label ?? ''),
      accelerators: settings.map((i) => i.accelerator ?? ''),
    }
  })

const clickSettings = (app: ElectronApplication): Promise<boolean> =>
  app.evaluate(({ Menu }) => {
    const menu = Menu.getApplicationMenu()
    if (!menu) return false
    for (const top of menu.items) {
      for (const item of top.submenu?.items ?? []) {
        if (item.accelerator === 'CmdOrCtrl+,' && item.click) {
          item.click()
          return true
        }
      }
    }
    return false
  })

test('every app menu carries a Settings entry that opens the Settings view', async () => {
  test.setTimeout(300_000)
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-settings-'))
  const mdPath = join(dir, 'sample.md')
  await writeFile(mdPath, '# Sample\n\nA markdown note.\n')

  const launched = await launchShell({ onboardingSeen: true, videoDir: 'menu-settings' })
  const { page, app } = launched
  try {
    // the home menu (the menu is installed a moment after the window opens)
    await expect
      .poll(async () => (await settingsReport(app)).found, { timeout: 30_000 })
      .toBeGreaterThan(0)
    const home = await settingsReport(app)
    expect(
      home.accelerators.some((a) => a === 'CmdOrCtrl+,'),
      JSON.stringify(home),
    ).toBe(true)

    // clicking it opens the Settings view
    expect(await clickSettings(app)).toBe(true)
    await expect(page.locator('.set-dialog')).toBeVisible({ timeout: 30_000 })

    // and the same holds in another editor, whose menu is built elsewhere
    await page.evaluate((p) => window.aiOffice.openPath(p), mdPath)
    await waitForPageWithUrl(app, '://markdown/')
    await page.waitForTimeout(1500)

    await expect
      .poll(async () => (await settingsReport(app)).found, { timeout: 30_000 })
      .toBeGreaterThan(0)
    const markdown = await settingsReport(app)
    expect(markdown.accelerators.some((a) => a === 'CmdOrCtrl+,')).toBe(true)
  } finally {
    await closeAndSaveVideo(launched, 'menu-settings')
  }
})
