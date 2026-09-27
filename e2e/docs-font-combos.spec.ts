import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * Home font and size boxes follow Word: typing in the font box opens the list,
 * completes inline to the first matching name and Enter commits it; Esc or a
 * blur without Enter restores the shown value in both boxes; the size box takes
 * a typed value on Enter only.
 */

interface AidocsWindow {
  __aidocs?: { editor?: { getAttributes: (name: string) => Record<string, unknown> } }
}

test.describe('docs font and size combos', () => {
  test('type-ahead commits on Enter while Esc and blur restore', async () => {
    test.setTimeout(120_000)
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-font-combos' })
    const { app, page } = launched
    try {
      await expect(page.locator('.quick-card').first()).toContainText('AI Docs')
      await page.locator('.quick-card').first().click()
      const editorPage = await waitForPageWithUrl(app, '://docs/')
      await editorPage.waitForFunction(
        () => Boolean((window as unknown as AidocsWindow).__aidocs?.editor),
        undefined,
        { timeout: 30_000 },
      )
      const attrs = () =>
        editorPage.evaluate(() =>
          (window as unknown as AidocsWindow).__aidocs!.editor!.getAttributes('docTextStyle'),
        )
      await editorPage.locator('.doc-page').click()
      await editorPage.keyboard.type('combo text', { delay: 10 })
      await editorPage.keyboard.press('ControlOrMeta+a')

      const fontBox = editorPage.locator('input.rb-font-family')
      const sizeBox = editorPage.locator('input.rb-font-size')
      const menu = editorPage.locator('.rb-font-family-menu')
      const shownFont = await fontBox.inputValue()
      const shownSize = await sizeBox.inputValue()
      expect(shownFont).not.toBe('')

      await fontBox.click()
      await editorPage.keyboard.type('ge')
      await expect(menu).toBeVisible()
      await expect(fontBox).toHaveValue('Georgia')
      await expect(menu.locator('button.kbd-focus')).toHaveText('Georgia')
      expect(
        await fontBox.evaluate((el: HTMLInputElement) => [el.selectionStart, el.selectionEnd]),
      ).toEqual([2, 'Georgia'.length])
      await editorPage.keyboard.press('ArrowDown')
      await expect(fontBox).toHaveValue('Verdana')
      await editorPage.keyboard.press('Escape')
      await expect(menu).toBeHidden()
      await expect(fontBox).toHaveValue(shownFont)
      expect((await attrs()).fontAscii ?? null).not.toBe('Verdana')

      await fontBox.click()
      await editorPage.keyboard.type('ta')
      await expect(fontBox).toHaveValue('Tahoma')
      await editorPage.keyboard.press('Enter')
      await expect(menu).toBeHidden()
      await expect(fontBox).toHaveValue('Tahoma')
      expect((await attrs()).fontAscii).toBe('Tahoma')

      // blur without Enter keeps the selection's font
      await fontBox.click()
      await editorPage.keyboard.type('Impact')
      await editorPage.locator('.doc-page').click()
      await expect(fontBox).toHaveValue('Tahoma')
      expect((await attrs()).fontAscii).toBe('Tahoma')
      await editorPage.locator('.rb-font-group .rb-combo-caret').first().click()
      await expect(menu).toBeVisible()
      // a list click still commits although the press blurs the box first
      await menu.locator('button', { hasText: 'Georgia' }).first().click()
      await expect(menu).toBeHidden()
      await expect(fontBox).toHaveValue('Georgia')
      expect((await attrs()).fontAscii).toBe('Georgia')

      await editorPage.keyboard.press('ControlOrMeta+a')
      await sizeBox.click()
      await editorPage.keyboard.type('30')
      await editorPage.keyboard.press('Escape')
      await expect(sizeBox).toHaveValue(shownSize)
      await sizeBox.click()
      await editorPage.keyboard.type('30')
      await editorPage.locator('.doc-page').click()
      await expect(sizeBox).toHaveValue(shownSize)
      expect((await attrs()).sizeHalfPoints ?? null).not.toBe(60)
      await editorPage.keyboard.press('ControlOrMeta+a')
      await sizeBox.click()
      await editorPage.keyboard.type('14')
      await editorPage.keyboard.press('Enter')
      await expect(sizeBox).toHaveValue('14')
      expect((await attrs()).sizeHalfPoints).toBe(28)
    } finally {
      await closeAndSaveVideo(launched, 'docs-font-combos')
    }
  })
})
