/**
 * Review → Markup → Track Changes Options… (Word's Advanced Track Changes
 * Options): the dialog opens from the ribbon, a fixed deletion color repaints
 * the struck run on screen, the choice persists, and the saved docx carries
 * exactly the same revision markup whatever the options say.
 */
import { test, expect } from '@playwright/test'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

type AidocsWindow = Window & { __aidocs?: { save?: () => Promise<unknown> } }

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const REV = 'w:author="Ann Lee" w:date="2026-09-01T10:00:00Z"'

async function trackedDocx(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body>` +
      '<w:p><w:r><w:t>Alpha opens the document.</w:t></w:r></w:p>' +
      `<w:p><w:r><w:t xml:space="preserve">Gamma keeps </w:t></w:r><w:del w:id="2" ${REV}><w:r><w:delText>removed words</w:delText></w:r></w:del><w:ins w:id="3" ${REV}><w:r><w:t>added words</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> around.</w:t></w:r></w:p>` +
      '</w:body></w:document>',
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function documentXml(path: string): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(path))
  return zip.file('word/document.xml')!.async('string')
}

test.describe('docs Track Changes Options', () => {
  let dir: string
  let docPath: string

  test.beforeEach(async () => {
    dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-tcopts-')))
    docPath = join(dir, 'tracked.docx')
    writeFileSync(docPath, await trackedDocx())
  })

  test.afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  test('a fixed deletion color repaints the screen but not the saved markup', async () => {
    test.setTimeout(180_000)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-track-changes-options',
      openFile: docPath,
    })
    const { app } = launched
    try {
      const page = await waitForPageWithUrl(app, '://docs/')
      await page.locator('.doc-page p', { hasText: 'Alpha opens' }).first().waitFor()
      const del = page.locator('.doc-del', { hasText: 'removed words' })
      await expect(del).toBeVisible()
      const before = await del.evaluate((el) => getComputedStyle(el).color)

      // a real write of the untouched document: the baseline markup
      const save = () => page.evaluate(() => (window as unknown as AidocsWindow).__aidocs!.save!())
      await page.locator('.doc-page p', { hasText: 'Alpha opens' }).first().click()
      await page.keyboard.press('End')
      await page.keyboard.type('Z')
      await page.keyboard.press('Backspace')
      expect(await save()).toBe(true)
      const baseline = await documentXml(docPath)
      expect(baseline).toContain('<w:delText>removed words</w:delText>')

      await page.locator('.ribbon-tab', { hasText: /^Review$/ }).click()
      await page.getByRole('button', { name: 'Markup', exact: true }).click()
      await page.getByTestId('tc-options').click()
      const dialog = page.locator('.modal.tco-dialog')
      await expect(dialog).toBeVisible()
      await expect(dialog.locator('select[name="del-mark"]')).toHaveValue('strike')
      await expect(dialog.locator('select[name="del-color"]')).toHaveValue('author')
      await dialog.locator('select[name="del-color"]').selectOption('red')
      await dialog.locator('select[name="ins-mark"]').selectOption('bold')
      await dialog.getByRole('button', { name: 'OK' }).click()
      await expect(dialog).toBeHidden()

      await expect(del).toHaveCSS('color', 'rgb(255, 0, 0)')
      expect(await del.evaluate((el) => getComputedStyle(el).color)).not.toBe(before)
      await expect(del).toHaveCSS('text-decoration-line', 'line-through')
      const ins = page.locator('.doc-ins', { hasText: 'added words' })
      await expect(ins).toHaveCSS('font-weight', '700')
      await expect(ins).toHaveCSS('text-decoration-line', 'none')
      expect(
        JSON.parse(
          await page.evaluate(() => localStorage.getItem('aidocs.trackChangesOptions') ?? '{}'),
        ),
      ).toMatchObject({ delColor: 'red', insMark: 'bold' })

      // the options are display-only: the second save writes identical markup
      await page.locator('.doc-page p', { hasText: 'Alpha opens' }).first().click()
      await page.keyboard.press('End')
      await page.keyboard.type('Z')
      await page.keyboard.press('Backspace')
      expect(await save()).toBe(true)
      expect(await documentXml(docPath)).toBe(baseline)

      // Reset to defaults puts the author ink back
      await page.getByRole('button', { name: 'Markup', exact: true }).click()
      await page.getByTestId('tc-options').click()
      await expect(dialog).toBeVisible()
      await dialog.getByRole('button', { name: 'Reset to defaults' }).click()
      await expect(dialog.locator('select[name="del-color"]')).toHaveValue('author')
      await dialog.getByRole('button', { name: 'OK' }).click()
      await expect(del).toHaveCSS('color', before)
      await expect(ins).toHaveCSS('text-decoration-line', 'underline')
    } finally {
      await closeAndSaveVideo(launched, 'docs-track-changes-options')
    }
  })
})
