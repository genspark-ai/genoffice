import { test, expect } from '@playwright/test'
import { existsSync, mkdirSync, mkdtempSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, screenshotPath, waitForPageWithUrl } from './helpers'

/**
 * The editors' Files pane (shared @genoffice/ui component over the shell's
 * folder IPC), exercised in the Markdown app: toggle from the ribbon, current
 * file highlighted, click another file → the shell opens it in a new tab,
 * create a folder from the pane header.
 */
test.describe('editor files pane', () => {
  let root: string

  test.beforeEach(() => {
    root = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-files-')))
    mkdirSync(join(root, 'Notes'))
    writeFileSync(join(root, 'Notes', 'first.md'), '# first')
    writeFileSync(join(root, 'Notes', 'second.md'), '# second')
    writeFileSync(join(root, 'plan.docx'), 'x')
  })

  test.afterEach(() => {
    rmSync(root, { recursive: true, force: true })
  })

  test('markdown: ribbon toggle, current highlight, open sibling in a new tab, new folder', async () => {
    const launched = await launchShell({
      onboardingSeen: true,
      settings: { defaultSaveDir: root },
      videoDir: 'files-pane',
      openFile: join(root, 'Notes', 'first.md'),
    })
    const { app } = launched
    try {
      // with a file on argv the editor view may be the first window Playwright sees
      const shellPage = await waitForPageWithUrl(app, 'shell/out/renderer/')
      const editor = await waitForPageWithUrl(app, '://markdown/')
      await expect(editor.locator('.doc-editor')).toBeVisible()
      // closed by default: the edge tab is the way in
      await expect(editor.locator('.files-pane')).toHaveCount(0)
      await editor.locator('.files-edge-tab').click()
      const pane = editor.locator('.files-pane')
      await expect(pane).toBeVisible()
      await expect(pane.locator('.fp-title')).toHaveText('Files')
      // the current file's folder is revealed and the file highlighted
      await expect(pane.locator('.fp-row.current .fp-name')).toHaveText('first.md')
      await expect(pane.locator('.fp-name', { hasText: 'plan.docx' })).toBeVisible()
      await editor.screenshot({ path: screenshotPath('files-pane-markdown') })

      // clicking a sibling opens it in a second shell tab
      await pane
        .locator('.fp-row', { has: editor.locator('.fp-name', { hasText: 'second.md' }) })
        .click()
      await expect(shellPage.locator('.tab-bar .tab-item', { hasText: 'second.md' })).toBeVisible()

      // the ribbon toggle closes it again; the edge tab comes back
      await editor.locator('.rb-btn[aria-label="Files"]').click()
      await expect(editor.locator('.files-pane')).toHaveCount(0)
      await expect(editor.locator('.files-edge-tab')).toBeVisible()

      // new folder from the pane header lands next to the current file
      await editor.locator('.rb-btn[aria-label="Files"]').click()
      await pane.locator('.fp-head-btn[aria-label="New folder"]').click()
      const input = pane.locator('.fp-input')
      await input.fill('Drafts')
      await input.press('Enter')
      await expect(pane.locator('.fp-name', { hasText: 'Drafts' })).toBeVisible()
      expect(existsSync(join(root, 'Notes', 'Drafts'))).toBe(true)
    } finally {
      await closeAndSaveVideo(launched, 'files-pane')
    }
  })
})
