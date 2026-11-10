import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * A `.txt` / `.json` file has no document model behind it, so the shell route
 * used to refuse it: the open dialog and the Home list both offered the file,
 * and picking it fell through to the unsupported-file notice. These two tests
 * cover the two ways in, and pin the byte-level round-trip the source editor
 * promises on the way back out.
 */

/** BOM + CRLF + no trailing newline: the three things a naive writer normalises. */
const BOM = '\uFEFF'
const ORIGINAL_TEXT = `${BOM}first line\r\nsecond line`
test.describe('source text files', () => {
  test('a .txt picked on Home opens as source and saves back byte-identical', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-txt-home-'))
    const txtPath = join(dir, 'notes.txt')
    await writeFile(txtPath, Buffer.from(ORIGINAL_TEXT, 'utf8'))

    // a watch folder is what puts the file on Home in the first place
    const launched = await launchShell({
      onboardingSeen: true,
      settings: { defaultSaveDir: dir },
      videoDir: 'txt-open-from-home',
    })
    const { app, page } = launched
    try {
      const tree = page.locator('.folder-panel .tree')
      const rootRow = tree.locator('.tree-row').first()
      await expect(rootRow).toContainText(dir.split('/').pop()!)
      await rootRow.locator('.tree-name').click()

      const row = page.locator('.recent-list .recent-name', { hasText: 'notes.txt' })
      await expect(row).toBeVisible()
      await row.click()

      // the source surface, not the block editor: a .txt has no blocks to edit
      const editorPage = await waitForPageWithUrl(app, '://markdown/')
      await expect(editorPage.locator('.source-editor')).toBeVisible()
      await expect(editorPage.locator('.source-editor .cm-content')).toContainText('first line')
      await expect(editorPage.locator('.source-editor .cm-content')).toContainText('second line')
      await expect(editorPage.locator('.doc-editor')).toHaveCount(0)

      // save without editing: the file must come back exactly as it went in
      await editorPage.locator('.source-editor .cm-content').click()
      await editorPage.keyboard.press('ControlOrMeta+s')
      await expect(editorPage.locator('.status-save')).toHaveText(/Saved/)

      expect(await readFile(txtPath)).toEqual(Buffer.from(ORIGINAL_TEXT, 'utf8'))
    } finally {
      await closeAndSaveVideo(launched, 'txt-open-from-home')
    }
  })

  test('editing a BOM + CRLF .txt keeps the BOM, the CRLF and the missing trailing newline', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-txt-roundtrip-'))
    const txtPath = join(dir, 'notes.txt')
    await writeFile(txtPath, Buffer.from(ORIGINAL_TEXT, 'utf8'))

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'txt-roundtrip',
      openFile: txtPath,
    })
    const { app } = launched
    try {
      const editorPage = await waitForPageWithUrl(app, '://markdown/')
      const content = editorPage.locator('.source-editor .cm-content')
      await expect(content).toBeVisible()
      await expect(content).toContainText('first line')
      await expect(content).toContainText('second line')

      await content.click()
      await editorPage.keyboard.press('ControlOrMeta+End')
      await editorPage.keyboard.type('!')
      await editorPage.keyboard.press('ControlOrMeta+s')
      await expect(editorPage.locator('.status-save')).toHaveText(/Saved/)

      // one character appended, and nothing else touched: same BOM, same CRLF,
      // and still no trailing newline (a writer that "helpfully" added one would
      // show up here as an extra byte)
      const expected = `${BOM}first line\r\nsecond line!`
      expect(await readFile(txtPath)).toEqual(Buffer.from(expected, 'utf8'))
    } finally {
      await closeAndSaveVideo(launched, 'txt-roundtrip')
    }
  })
})
