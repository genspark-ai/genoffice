import { test, expect } from '@playwright/test'
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readdirSync,
  realpathSync,
  rmSync,
  writeFileSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, screenshotPath } from './helpers'

/**
 * Home "Folders" panel: the tree over the default save folder, the folder
 * view, and the create / rename / move / delete actions against the real disk.
 */
test.describe('home folders panel', () => {
  let root: string

  test.beforeEach(() => {
    root = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-root-')))
    mkdirSync(join(root, 'Clients', 'A Corp'), { recursive: true })
    mkdirSync(join(root, 'Clients', 'Contracts'), { recursive: true })
    mkdirSync(join(root, 'Personal'))
    writeFileSync(join(root, 'report.docx'), 'x')
    writeFileSync(join(root, 'notes.md'), '# notes')
    writeFileSync(join(root, 'Clients', 'Contracts', 'deal.md'), '# deal')
  })

  test.afterEach(() => {
    rmSync(root, { recursive: true, force: true })
  })

  test('shows the tree, opens a folder, then creates / moves / renames / deletes on disk', async () => {
    const launched = await launchShell({
      onboardingSeen: true,
      settings: { defaultSaveDir: root },
      videoDir: 'home-folders',
    })
    const { page } = launched
    try {
      const tree = page.locator('.folder-panel .tree')
      await expect(page.locator('.folder-panel-title')).toHaveText('Folders')
      const rootRow = tree.locator('.tree-row').first()
      await expect(rootRow).toContainText(root.split('/').pop()!)
      // root is expanded by default: its first-level folders are listed
      await expect(tree.locator('.tree-name', { hasText: 'Clients' })).toBeVisible()
      await expect(tree.locator('.tree-name', { hasText: 'Personal' })).toBeVisible()
      await expect(tree.locator('.tree-name', { hasText: 'Contracts' })).toHaveCount(0)

      // expand Clients → its children appear
      const clientsRow = tree.locator('.tree-row', {
        has: page.locator('.tree-name', { hasText: 'Clients' }),
      })
      await clientsRow.locator('.tree-chevron').click()
      await expect(tree.locator('.tree-name', { hasText: 'Contracts' })).toBeVisible()

      // select Contracts → active tree row + its file
      await tree.locator('.tree-name', { hasText: 'Contracts' }).click()
      await expect(tree.locator('.tree-row.active .tree-name')).toHaveText('Contracts')
      await expect(page.locator('.recent-list .recent-name', { hasText: 'deal.md' })).toBeVisible()
      await page.screenshot({ path: screenshotPath('home-folders-contracts') })

      // new folder from the sidebar header: lands under the selected folder
      await page.locator('.folder-add-btn').click()
      const input = page.locator('.folder-panel .tree .folder-rename-input')
      await input.fill('Drafts')
      await input.press('Enter')
      await expect(
        page.locator('.recent-list .folder-item .recent-name', { hasText: 'Drafts' }),
      ).toBeVisible()
      expect(existsSync(join(root, 'Clients', 'Contracts', 'Drafts'))).toBe(true)

      // move deal.md → Personal through the picker
      const dealRow = page.locator('.recent-row', {
        has: page.locator('.recent-name', { hasText: 'deal.md' }),
      })
      await dealRow.locator('.more-btn').click()
      await page.locator('.row-menu button', { hasText: 'Move to folder' }).click()
      const picker = page.locator('.picker-modal')
      await expect(picker).toBeVisible()
      await picker
        .locator('.picker-row', { has: page.locator('.picker-name', { hasText: 'Personal' }) })
        .click()
      await picker.locator('.btn-primary').click()
      await expect(picker).toHaveCount(0)
      await expect(page.locator('.recent-list .recent-name', { hasText: 'deal.md' })).toHaveCount(0)
      expect(existsSync(join(root, 'Personal', 'deal.md'))).toBe(true)
      expect(existsSync(join(root, 'Clients', 'Contracts', 'deal.md'))).toBe(false)

      // rename Drafts → Final from the sub-folder row menu
      const draftsRow = page.locator('.recent-row', {
        has: page.locator('.recent-name', { hasText: 'Drafts' }),
      })
      await draftsRow.locator('.more-btn').click()
      await page.locator('.folder-menu button', { hasText: 'Rename' }).click()
      const renameInput = page.locator('.recent-list .rename-input')
      await renameInput.fill('Final')
      await renameInput.press('Enter')
      await expect(
        page.locator('.recent-list .folder-item .recent-name', { hasText: 'Final' }),
      ).toBeVisible()
      expect(existsSync(join(root, 'Clients', 'Contracts', 'Final'))).toBe(true)

      // delete Final (confirm dialog → trash)
      const finalRow = page.locator('.recent-row', {
        has: page.locator('.recent-name', { hasText: 'Final' }),
      })
      await finalRow.locator('.more-btn').click()
      await page.locator('.folder-menu button.danger').click()
      await page.locator('.modal .btn-danger').click()
      await expect(page.locator('.recent-list .recent-name', { hasText: 'Final' })).toHaveCount(0)
      expect(existsSync(join(root, 'Clients', 'Contracts', 'Final'))).toBe(false)

      // Personal now lists the moved file; the Recent view shows its location
      await tree.locator('.tree-name', { hasText: 'Personal' }).click()
      await expect(page.locator('.recent-list .recent-name', { hasText: 'deal.md' })).toBeVisible()
      await page.screenshot({ path: screenshotPath('home-folders-personal') })
    } finally {
      await closeAndSaveVideo(launched, 'home-folders')
    }
  })

  test('a file created from a folder view lands in that folder', async () => {
    const launched = await launchShell({
      onboardingSeen: true,
      settings: { defaultSaveDir: root },
      videoDir: 'home-folders-new-file',
    })
    const { page, app } = launched
    try {
      await page.locator('.folder-panel .tree .tree-name', { hasText: 'Personal' }).click()
      await expect(page.locator('.folder-panel .tree .tree-row.active .tree-name')).toHaveText(
        'Personal',
      )
      // the blank PDF is written synchronously by the shell, so it exercises the pending-folder path
      await page.locator('.quick-card', { hasText: 'AI PDF' }).click()
      const hasPdf = (dir: string) =>
        existsSync(dir) && readdirSync(dir).some((f) => f.endsWith('.pdf'))
      await expect.poll(() => hasPdf(join(root, 'Personal')), { timeout: 15_000 }).toBe(true)
      expect(hasPdf(root)).toBe(false)
      // and the tab opened on the moved path, not the vanished root one
      await expect(page.locator('.tab-bar .tab-item', { hasText: '.pdf' })).toBeVisible()
      // a second New from the same folder view lands there too (the folder is not a one-shot slot)
      await page.locator('.tab-bar .tab-item.tab-home').click()
      await page.locator('.quick-card', { hasText: 'AI Sheets' }).click()
      const hasXlsx = (dir: string) =>
        existsSync(dir) && readdirSync(dir).some((f) => f.endsWith('.xlsx'))
      await expect.poll(() => hasXlsx(join(root, 'Personal')), { timeout: 15_000 }).toBe(true)
      expect(hasXlsx(root)).toBe(false)
      // back home: the new tab must not block shutdown
      await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0]?.focus())
    } finally {
      await closeAndSaveVideo(launched, 'home-folders-new-file')
    }
  })
})
