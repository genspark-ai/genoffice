import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

function themeAttr(page: Page): Promise<string | null> {
  return page.evaluate(() => document.documentElement.getAttribute('data-theme'))
}

function hasHomeApi(page: Page): Promise<boolean> {
  return page
    .evaluate(() => Boolean((window as unknown as { aiOffice?: unknown }).aiOffice))
    .catch(() => false)
}

/** firstWindow() order differs between platforms — find the shell page by its API */
async function findShellPage(app: ElectronApplication, timeoutMs = 15_000): Promise<Page> {
  const deadline = Date.now() + timeoutMs
  for (;;) {
    for (const candidate of app.windows()) {
      if (await hasHomeApi(candidate)) return candidate
    }
    const remaining = deadline - Date.now()
    if (remaining <= 0) throw new Error('No window exposing window.aiOffice')
    await app.waitForEvent('window', { timeout: Math.min(remaining, 1_000) }).catch(() => {})
  }
}

function setTheme(page: Page, theme: 'light' | 'dark' | 'system'): Promise<void> {
  return page.evaluate((t) => {
    const api = (window as unknown as { aiOffice: { setTheme(v: string): Promise<void> } }).aiOffice
    return api.setTheme(t)
  }, theme)
}

test.describe('theme pipeline', () => {
  test('setTheme reaches home and editor tabs, persists across relaunch', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-theme-'))
    const mdPath = join(dir, 'doc.md')
    await writeFile(mdPath, '# Doc\n\nBody.\n')

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'theme-pipeline',
      openFile: mdPath,
    })
    const { app } = launched
    try {
      const shellPage = await findShellPage(app)
      const editorPage = await waitForPageWithUrl(app, 'markdown/out')
      await expect(editorPage.locator('.doc-editor')).toBeVisible()
      expect(await themeAttr(shellPage)).toBeNull()
      expect(await themeAttr(editorPage)).toBeNull()

      await setTheme(shellPage, 'dark')
      await expect.poll(() => themeAttr(shellPage)).toBe('dark')
      await expect.poll(() => themeAttr(editorPage)).toBe('dark')

      // native chrome follows the explicit choice
      expect(await app.evaluate(({ nativeTheme }) => nativeTheme.themeSource)).toBe('dark')

      await setTheme(shellPage, 'system')
      await expect.poll(() => themeAttr(shellPage)).toBeNull()
      await expect.poll(() => themeAttr(editorPage)).toBeNull()

      await setTheme(shellPage, 'dark')
      await expect.poll(() => themeAttr(shellPage)).toBe('dark')
    } finally {
      await closeAndSaveVideo(launched, 'theme-pipeline')
    }

    // relaunch with the same userData: the persisted theme applies before first
    // paint (no onboardingSeen — that option rewrites app-settings.json wholesale)
    const relaunched = await launchShell({
      userDataDir: launched.userDataDir,
      videoDir: 'theme-pipeline-relaunch',
    })
    try {
      const shellPage = await findShellPage(relaunched.app)
      await expect.poll(() => themeAttr(shellPage)).toBe('dark')
      expect(await relaunched.app.evaluate(({ nativeTheme }) => nativeTheme.themeSource)).toBe(
        'dark',
      )
    } finally {
      await closeAndSaveVideo(relaunched, 'theme-pipeline-relaunch')
    }
  })
})
