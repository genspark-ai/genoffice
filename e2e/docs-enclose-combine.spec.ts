/**
 * Format > Asian Layout > Enclose Characters and Combine Characters through
 * the menu, the ribbon and their dialogs: the characters draw overlaid, save
 * as the EQ field Word writes, and a field already in the file reads back.
 */
import { test, expect } from '@playwright/test'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

/** Enclose Characters as Word writes it, the enclosure enlarged */
const WORD_ENCLOSED =
  '<w:p><w:r><w:t xml:space="preserve">From Word: </w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> eq \\o\\ac(</w:instrText></w:r>' +
  '<w:r><w:rPr><w:position w:val="-4"/><w:sz w:val="31"/></w:rPr><w:instrText>\u25cb</w:instrText></w:r>' +
  '<w:r><w:instrText>,W)</w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'

const documentXml = async (path: string): Promise<string> => {
  const zip = await JSZip.loadAsync(readFileSync(path))
  return (await zip.file('word/document.xml')?.async('string')) ?? ''
}

test('characters are enclosed and combined from their dialogs and saved as the fields Word writes', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-eq-layout-')))
  const docPath = join(dir, 'layout.docx')
  writeFileSync(docPath, await docxWith(para('A grand day') + WORD_ENCLOSED + para('Last')))
  let launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-enclose-combine',
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const first = page.locator('.doc-page p', { hasText: 'A grand day' })
    await expect(first).toBeVisible({ timeout: 30_000 })
    const click = (...path: string[]) =>
      app.evaluate(({ Menu }, path) => {
        let items = Menu.getApplicationMenu()?.items ?? []
        for (const label of path.slice(0, -1))
          items = items.find((item) => item.label === label)?.submenu?.items ?? []
        items.find((item) => item.label === path[path.length - 1])?.click()
      }, path)
    /** selects document positions: the first paragraph's text starts at 1 */
    const select = (from: number, to: number) =>
      page.evaluate(
        ([from, to]) => {
          const { editor } = (
            window as never as {
              __aidocs: {
                editor: {
                  chain(): {
                    focus(): { setTextSelection(r: object): { run(): boolean } }
                  }
                }
              }
            }
          ).__aidocs
          editor.chain().focus().setTextSelection({ from, to }).run()
        },
        [from, to],
      )
    const enclosed = page.locator('.doc-page .doc-enclose-chars')
    const combined = page.locator('.doc-page .doc-combine-chars')
    // the field already in the file draws as an enclosed character
    await expect(enclosed).toHaveCount(1)
    await expect(enclosed.first()).toHaveText('\u25cbW')

    // enclose the first letter: a square, the symbol enlarged
    await select(1, 2)
    await click('Format', 'Asian Layout', 'Enclose Characters…')
    const dialog = page.locator('.eq-layout-dialog')
    await expect(dialog.locator('input[type=text]')).toHaveValue('A')
    await dialog.locator('.eq-layout-choice', { hasText: 'Enlarge symbol' }).click()
    await dialog.locator('.eq-layout-mark', { hasText: '\u25a1' }).click()
    await page.screenshot({ path: test.info().outputPath('enclose-dialog.png') })
    await dialog.locator('button.primary').click()
    await expect(enclosed).toHaveCount(2)
    await expect(enclosed.first()).toHaveText('\u25a1A')

    // combine "grand" from the Home tab's Asian Layout button
    await select(3, 8)
    await page.locator('button[aria-label="Asian Layout"]').click()
    await page.locator('.layout-menu button', { hasText: 'Combine Characters' }).click()
    await expect(dialog.locator('input[type=text]')).toHaveValue('grand')
    await page.screenshot({ path: test.info().outputPath('combine-dialog.png') })
    await dialog.locator('button.primary').click()
    await expect(combined).toHaveCount(1)
    await expect(combined.locator('span').first()).toHaveText('gra')
    await expect(combined.locator('span').last()).toHaveText('nd')
    await page.screenshot({ path: test.info().outputPath('result.png') })

    // the combination is selected: the dialog opens on it and Remove gives the text back
    await page.locator('button[aria-label="Asian Layout"]').click()
    await page.locator('.layout-menu button', { hasText: 'Combine Characters' }).click()
    await dialog.locator('button', { hasText: /^Remove$/ }).click()
    await expect(combined).toHaveCount(0)
    await expect(first).toContainText('grand day')
    // combined again for the save
    await select(3, 8)
    await page.locator('button[aria-label="Asian Layout"]').click()
    await page.locator('.layout-menu button', { hasText: 'Combine Characters' }).click()
    await dialog.locator('button.primary').click()
    await expect(combined).toHaveCount(1)

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const xml = await documentXml(docPath)
          return [
            xml.match(/ eq \\o\\ac\(/g)?.length ?? 0,
            xml.match(/ eq \\o\(\\s\\up \d+\(/g)?.length ?? 0,
            // the field Word wrote keeps the sizes Word gave it
            xml.includes('<w:position w:val="-4"/><w:sz w:val="31"/>'),
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([2, 1, true])
  } finally {
    await closeAndSaveVideo(launched, 'docs-enclose-combine')
  }

  launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-enclose-combine',
    openFile: docPath,
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(page.locator('.doc-page .doc-enclose-chars')).toHaveCount(2, { timeout: 30_000 })
    await expect(page.locator('.doc-page .doc-combine-chars')).toHaveCount(1)
    await expect(page.locator('.doc-page .doc-enclose-chars').first()).toHaveText('\u25a1A')
    await page.screenshot({ path: test.info().outputPath('reopened.png') })
  } finally {
    await closeAndSaveVideo(launched, 'docs-enclose-combine')
    rmSync(dir, { recursive: true, force: true })
  }
})
