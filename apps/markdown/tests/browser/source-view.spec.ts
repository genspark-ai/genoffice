import { expect, test } from '@playwright/test'
import type { Editor } from '@tiptap/core'
import { openSource, source } from './helpers'

/**
 * Source view (App.tsx + SourcePane): the ribbon toggle swaps the document
 * canvas for the exact file text, and an edit made there reaches the editor and
 * therefore the save path — no separate write route.
 */

const SOURCE_BUTTON = '.rb-btn[aria-label="Source"]'

test('the source view shows the exact file text and routes edits through the editor', async ({
  page,
}) => {
  await openSource(page, false)
  const pane = page.locator('.source-textarea')
  await expect(pane).toHaveCount(0)

  await page.locator(SOURCE_BUTTON).click()
  await expect(pane).toBeVisible()
  // what a save would write, character for character
  await expect(pane).toHaveValue(source)
  // the canvas is hidden, not torn down: it comes back intact
  await expect(page.locator('.doc-editor')).toBeHidden()

  await pane.fill('Title\n=====\n\n* item edited\n')
  await expect(page.locator('.doc-editor')).toContainText('item edited')

  await page.evaluate(() => window.dispatchEvent(new Event('test:save')))
  await expect(page.locator('body')).toHaveAttribute('data-saved', /item edited/)

  await page.locator(SOURCE_BUTTON).click()
  await expect(page.locator('.doc-editor')).toBeVisible()
  await expect(page.locator('.doc-editor')).toContainText('item edited')
})

test('Cmd/Ctrl+E toggles the source view', async ({ page }) => {
  await openSource(page, false)
  const pane = page.locator('.source-textarea')
  await expect(pane).toHaveCount(0)
  await page.keyboard.press('ControlOrMeta+e')
  await expect(pane).toBeVisible()
  await expect(pane).toBeFocused()
  await page.keyboard.press('ControlOrMeta+e')
  await expect(pane).toHaveCount(0)
  await expect(page.locator('.doc-editor')).toBeVisible()
})

test('editor-shaped ribbon commands stand down while the source view is open', async ({ page }) => {
  await openSource(page, false)
  await page.locator(SOURCE_BUTTON).click()
  await expect(page.locator('.source-textarea')).toBeVisible()
  // save stays live; formatting, which acts on an invisible selection, does not
  await expect(page.locator(SOURCE_BUTTON)).toBeEnabled()
  await expect(page.locator('.rb-btn[aria-label="Bold"]')).toBeDisabled()
  await expect(page.locator('.rb-btn[aria-label="Bullet list"]')).toBeDisabled()
})

test('a write that landed while the pane sat unfocused is picked up on refocus', async ({
  page,
}) => {
  await openSource(page, false)
  const pane = page.locator('.source-textarea')
  await page.locator(SOURCE_BUTTON).click()
  await expect(pane).toBeFocused()
  // blur first, so the re-sync that matters is the one on the way back in
  await page.locator('.status-bar').click()
  // an AI run writes straight into the document, bypassing the pane
  await page.locator('.doc-editor').evaluate((node) => {
    const editor = (node as HTMLElement & { editor: Editor }).editor
    editor.commands.insertContentAt(1, 'appended by an agent\n\n')
  })
  await pane.click()
  await expect(pane).toHaveValue(/appended by an agent/)
})
