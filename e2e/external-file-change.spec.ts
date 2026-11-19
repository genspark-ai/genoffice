import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

// watchFile polls once a second; a 2 s gap keeps the mtime distinct on coarse filesystems
const settle = () => new Promise((r) => setTimeout(r, 2000))

test.describe('reload when another program changes the open file', () => {
  test('markdown: clean document reloads, dirty document asks first', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-md-disk-'))
    const mdPath = join(dir, 'shared.md')
    await writeFile(mdPath, '# First\n\nhello\n')

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'markdown-disk-reload',
      openFile: mdPath,
    })
    try {
      const page = await waitForPageWithUrl(launched.app, '://markdown/')
      const editor = page.locator('.doc-editor')
      await expect(editor.locator('h1')).toHaveText('First')

      await settle()
      await writeFile(mdPath, '# Second\n\nhello again\n')
      await expect(editor.locator('h1')).toHaveText('Second')
      await expect(editor).toContainText('hello again')
      await expect(page.locator('.disk-notice')).toHaveCount(0)

      await editor.locator('p').first().click()
      await page.keyboard.press('End')
      await page.keyboard.type(' mine')
      await settle()
      await writeFile(mdPath, '# Third\n\ntheirs\n')
      const notice = page.locator('.disk-notice')
      await expect(notice).toBeVisible()
      await expect(editor.locator('h1')).toHaveText('Second')

      await notice.getByRole('button', { name: 'Reload' }).click()
      await expect(editor.locator('h1')).toHaveText('Third')
      await expect(notice).toHaveCount(0)

      // reloading leaves the document clean: ⌘S has nothing to write back
      await page.keyboard.press(process.platform === 'darwin' ? 'Meta+s' : 'Control+s')
      await settle()
      expect(await readFile(mdPath, 'utf8')).toBe('# Third\n\ntheirs\n')
    } finally {
      await closeAndSaveVideo(launched, 'markdown-disk-reload')
    }
  })

  test('html: clean document reloads into the preview', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-html-disk-'))
    const htmlPath = join(dir, 'shared.html')
    await writeFile(htmlPath, '<!doctype html><html><body><h1>First</h1></body></html>\n')

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'html-disk-reload',
      openFile: htmlPath,
    })
    try {
      const page = await waitForPageWithUrl(launched.app, '://html/')
      const frame = page.frameLocator('.preview-frame')
      await expect(frame.locator('h1')).toHaveText('First')

      await settle()
      await writeFile(htmlPath, '<!doctype html><html><body><h1>Second</h1></body></html>\n')
      await expect(frame.locator('h1')).toHaveText('Second')
      await expect(page.locator('.disk-notice')).toHaveCount(0)
    } finally {
      await closeAndSaveVideo(launched, 'html-disk-reload')
    }
  })
})
