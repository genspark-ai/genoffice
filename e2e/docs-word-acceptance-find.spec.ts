/** Word acceptance tasks, find and replace: walked along Word's path and read back from the saved file. */
import { test, expect } from '@playwright/test'
import { para, saveFromMenu } from './docs-objects-fixture'
import { open, savedText } from './docs-word-acceptance-kit'

const VIDEO_DIR = 'docs-word-acceptance-find'

test('Control+H opens Replace, and Replace All changes every match in the document', async () => {
  test.setTimeout(180_000)
  const s = await open(
    VIDEO_DIR,
    ['The cat sat.', 'A cat and another cat.', 'No match here.'].map(para).join(''),
    'replace.docx',
  )
  try {
    const { app, page, docPath } = s
    await expect(page.locator('.doc-page p', { hasText: 'No match' })).toBeVisible({
      timeout: 30_000,
    })
    await page.locator('.doc-page p', { hasText: 'No match' }).click()
    await page.keyboard.press('Control+h')
    await expect(page.locator('.find-tab.on')).toHaveText('Replace')
    await page.locator('.find-input[aria-label="Find"]').fill('cat')
    await page.locator('.find-input[aria-label="Replace with"]').fill('dog')
    await expect(page.locator('.find-select')).toHaveValue('main')
    await page.screenshot({ path: test.info().outputPath('replace.png') })
    await page.getByRole('button', { name: 'Replace All', exact: true }).click()
    await expect(page.locator('.doc-page p', { hasText: 'another' })).toHaveText(
      'A dog and another dog.',
    )

    await saveFromMenu(app)
    await expect
      .poll(() => savedText(docPath), { timeout: 30_000 })
      .toEqual(['The dog sat.', 'A dog and another dog.', 'No match here.'])
  } finally {
    await s.close()
  }
})
