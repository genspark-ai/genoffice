import { test, expect } from '@playwright/test'
import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * Settings → General → "User name" (genoffice#2029): comments were signed with
 * the OS account name and nothing in the app could change it. Type a name in
 * Settings, open a new deck and leave a comment; the card must carry that name.
 */
test.describe('#2029 the comment author follows Settings → General → User name', () => {
  test('a new Slides comment is signed with the configured name', async () => {
    test.setTimeout(300_000)
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'author-name-2029' })
    try {
      const shellPage = launched.page
      await shellPage.getByRole('button', { name: 'Settings', exact: true }).click()
      await shellPage.locator('.set-nav-item').filter({ hasText: 'General' }).click()
      const field = shellPage.getByRole('textbox', { name: 'User name' })
      await expect(field).toHaveAttribute('placeholder', /.+/)
      await field.fill('  Ada   Lovelace ')
      await field.press('Enter')
      await expect(field).toHaveValue('Ada Lovelace')
      const saved = JSON.parse(
        await readFile(join(launched.userDataDir, 'app-settings.json'), 'utf8'),
      )
      expect(saved.authorName).toBe('Ada Lovelace')
      await shellPage.locator('.set-close').click()

      await shellPage.locator('.quick-card', { hasText: 'AI Slides' }).click()
      const editor = await waitForPageWithUrl(launched.app, '://slides/')
      await editor.waitForSelector('.stage-wrap canvas', { timeout: 30_000 })

      await editor.locator('.ribbon-tab', { hasText: 'Review' }).click()
      await editor.getByRole('button', { name: 'New Comment' }).click()
      await editor.locator('.comment-new textarea').fill('Signed by the settings name')
      await editor.locator('.comment-new .btn-primary').click()
      await expect(editor.locator('.comment-card .comment-author')).toHaveText('Ada Lovelace')
    } finally {
      await closeAndSaveVideo(launched, 'author-name-2029')
    }
  })
})
