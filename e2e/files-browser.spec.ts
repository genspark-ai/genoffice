import { test, expect } from '@playwright/test'
import { mkdirSync, mkdtempSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * The files browser tab (issue #542): the Home sidebar card opens a
 * single-instance ?mode=files tab listing the folder roots; clicking a file
 * opens it in an editor tab without going through the Home screen, and
 * opening the browser again focuses the existing tab.
 */
test.describe('files browser tab', () => {
  let root: string

  test.beforeEach(() => {
    root = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-files-')))
    mkdirSync(join(root, 'Clients'), { recursive: true })
    writeFileSync(join(root, 'report.docx'), 'x')
    writeFileSync(join(root, 'notes.md'), '# notes')
    writeFileSync(join(root, 'Clients', 'deal.md'), '# deal')
  })

  test.afterEach(() => {
    rmSync(root, { recursive: true, force: true })
  })

  test('opens from Home, lists the default folder, and opens a file into a docs tab', async () => {
    const launched = await launchShell({
      onboardingSeen: true,
      settings: { defaultSaveDir: root },
      videoDir: 'files-browser',
    })
    const { app, page } = launched
    try {
      // the sidebar card opens the browser tab
      await page.locator('.nav-item', { hasText: 'Files' }).click()
      const files = await waitForPageWithUrl(app, '://files/')
      await expect(files.locator('.files-title')).toHaveText('Files')

      // the tree roots at the default save folder, expanded one level
      const rootName = root.split('/').pop()!
      const tree = files.locator('.files-tree')
      await expect(tree.locator('.files-row-name', { hasText: rootName })).toBeVisible()
      await expect(tree.locator('.files-row-name', { hasText: 'Clients' })).toBeVisible()

      // the detail pane lists the selected root's files; clicking one opens it
      await files.locator('.files-row.is-file', { hasText: 'report.docx' }).click()
      await waitForPageWithUrl(app, '://docs/')
      await expect(page.locator('.tab-item', { hasText: 'report.docx' })).toBeVisible()

      // single instance: the sidebar entry focuses the existing tab
      await page.locator('.tab-item', { hasText: rootName }).first().click()
      await page.locator('.nav-item', { hasText: 'Files' }).click()
      await expect(page.locator('.tab-item', { hasText: 'Files' })).toHaveCount(1)
      await expect(page.locator('.tab-item.active', { hasText: 'Files' })).toBeVisible()
    } finally {
      await closeAndSaveVideo(launched, 'files-browser')
    }
  })

  test('the tab strip titles the browser from the Files menu key and closes like a document tab', async () => {
    const launched = await launchShell({
      onboardingSeen: true,
      settings: { defaultSaveDir: root },
      videoDir: 'files-browser',
    })
    const { app, page } = launched
    try {
      await page.locator('.nav-item', { hasText: 'Files' }).click()
      await waitForPageWithUrl(app, '://files/')

      const filesTab = page.locator('.tab-item', { hasText: 'Files' })
      await expect(filesTab).toBeVisible()
      // closable like any document tab (only Home is pinned)
      await expect(filesTab.locator('.tab-close')).toBeVisible()
      await filesTab.hover()
      await filesTab.locator('.tab-close').click()
      await expect(page.locator('.tab-item', { hasText: 'Files' })).toHaveCount(0)
    } finally {
      await closeAndSaveVideo(launched, 'files-browser')
    }
  })
})
