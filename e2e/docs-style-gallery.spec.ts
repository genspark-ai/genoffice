import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * Home > Styles follows Word's Quick Style gallery: a new document lists Word's
 * 16 built-in quick styles in uiPriority order, Heading 3 appears once used,
 * cards apply paragraph / character styles by id, and the Styles Pane button
 * opens a searchable pane that applies too.
 */

const WORD_GALLERY = [
  'Normal',
  'No Spacing',
  'Heading 1',
  'Heading 2',
  'Title',
  'Subtitle',
  'Subtle Emphasis',
  'Emphasis',
  'Intense Emphasis',
  'Strong',
  'Quote',
  'Intense Quote',
  'Subtle Reference',
  'Intense Reference',
  'Book Title',
  'List Paragraph',
]

interface AidocsWindow {
  __aidocs?: {
    editor?: {
      getAttributes: (name: string) => Record<string, unknown>
      isActive: (name: string, attrs?: Record<string, unknown>) => boolean
    }
  }
}

test.describe('docs quick style gallery', () => {
  test('lists Word’s set in order, applies by id and opens the Styles pane', async () => {
    test.setTimeout(120_000)
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-style-gallery' })
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
      const paraAttrs = () =>
        editorPage.evaluate(() =>
          (window as unknown as AidocsWindow).__aidocs!.editor!.getAttributes('docParagraph'),
        )
      const gallery = editorPage.locator('.style-gallery')
      await expect(gallery.locator('.style-card-label')).toHaveText(WORD_GALLERY)

      await editorPage.locator('.doc-page').click()
      await editorPage.keyboard.type('Gallery text', { delay: 10 })
      await gallery.locator('.style-card[data-style-id="Title"]').click()
      expect((await paraAttrs()).styleId).toBe('Title')
      await expect(gallery.locator('.style-card[data-style-id="Title"]')).toHaveClass(/active/)

      // Heading 3 is semiHidden until used: the Opt+Cmd+3 shortcut brings it in after Heading 2
      await editorPage.keyboard.press('Alt+ControlOrMeta+Digit3')
      await expect(gallery.locator('.style-card-label').nth(4)).toHaveText('Heading 3')
      expect(
        await editorPage.evaluate(() =>
          (window as unknown as AidocsWindow).__aidocs!.editor!.isActive('docHeading', {
            level: 3,
          }),
        ),
      ).toBe(true)

      await editorPage.keyboard.press('ControlOrMeta+a')
      await gallery.locator('.style-card[data-style-id="Strong"]').click()
      expect(
        (
          await editorPage.evaluate(() =>
            (window as unknown as AidocsWindow).__aidocs!.editor!.getAttributes('docTextStyle'),
          )
        ).styleId,
      ).toBe('Strong')

      await editorPage.locator('.rb-styles-pane').click()
      const pane = editorPage.locator('.styles-pane')
      await expect(pane).toBeVisible()
      await expect(pane.locator('.style-pane-card')).toHaveCount(17)
      await pane.locator('.styles-pane-search').fill('quo')
      await expect(pane.locator('.style-pane-card')).toHaveCount(2)
      await pane.locator('.style-pane-card[data-style-id="Quote"]').click()
      expect((await paraAttrs()).styleId).toBe('Quote')
      await pane.locator('.styles-pane-search').fill('')
      await pane.locator('.styles-pane-foot select').selectOption('all')
      expect(await pane.locator('.style-pane-card').count()).toBeGreaterThan(20)
      await pane.locator('.nav-pane-close').click()
      await expect(pane).toHaveCount(0)
    } finally {
      await closeAndSaveVideo(launched, 'docs-style-gallery')
    }
  })
})
