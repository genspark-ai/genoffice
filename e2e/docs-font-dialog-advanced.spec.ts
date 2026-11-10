import { test, expect } from '@playwright/test'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, screenshotPath, waitForPageWithUrl } from './helpers'

/**
 * Font dialog (⌘D) Advanced tab: Expanded 2 pt + Scale 150% and the Small caps
 * effect apply to the selection, render on the span, save as w:spacing / w:w /
 * w:smallCaps on the run, and read back into the dialog after a reopen.
 */

interface AidocsWindow {
  __aidocs?: { editor?: unknown; save?: () => Promise<unknown>; openPath?: (p: string) => void }
}

async function minimalDocx(text: string): Promise<Buffer> {
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
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>${text}</w:t></w:r></w:p></w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function documentXml(path: string): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(path))
  return zip.file('word/document.xml')!.async('string')
}

test.describe('docs font dialog advanced tab', () => {
  let dir: string
  let docPath: string

  test.beforeEach(async () => {
    dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-font-')))
    docPath = join(dir, 'spacing.docx')
    writeFileSync(docPath, await minimalDocx('Spaced'))
  })

  test.afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  test('character spacing and small caps apply, save and read back', async () => {
    test.setTimeout(150_000)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-font-dialog-advanced',
      openFile: docPath,
    })
    const { app } = launched
    try {
      const editorPage = await waitForPageWithUrl(app, '://docs/')
      await editorPage.waitForFunction(
        () => Boolean((window as unknown as AidocsWindow).__aidocs?.editor),
        undefined,
        { timeout: 30_000 },
      )
      await editorPage.locator('.doc-page').click()
      await editorPage.keyboard.press('ControlOrMeta+a')
      await editorPage.keyboard.press('ControlOrMeta+d')
      const dialog = editorPage.locator('.modal.font-dialog')
      await expect(dialog).toBeVisible()

      await dialog.getByRole('tab', { name: 'Advanced' }).click()
      await dialog.getByLabel('Scale (%)').fill('150')
      await dialog.getByRole('button', { name: 'Spacing', exact: true }).click()
      await editorPage.getByRole('option', { name: 'Expanded' }).click()
      await dialog.getByLabel('By (pt)').first().fill('2')
      await dialog.screenshot({ path: screenshotPath('docs-font-dialog-advanced') })
      await dialog.getByRole('tab', { name: 'Font' }).click()
      await dialog.getByLabel('Small caps').check()
      await dialog.getByRole('button', { name: 'OK' }).click()
      await expect(dialog).toBeHidden()

      const span = editorPage.locator('.doc-page span[data-doc-style]').first()
      await expect(span).toHaveCSS('font-variant-caps', 'small-caps')
      const spacing = await span.evaluate((el) => getComputedStyle(el).letterSpacing)
      // 2pt (2.67px) plus the 150% scale approximation (+0.26em of 12pt ≈ 4.2px)
      expect(parseFloat(spacing)).toBeGreaterThan(5)

      expect(
        await editorPage.evaluate(() => (window as unknown as AidocsWindow).__aidocs!.save!()),
      ).toBe(true)
      await expect
        .poll(() => documentXml(docPath), { timeout: 15_000 })
        .toContain('<w:w w:val="150"/>')
      const xml = await documentXml(docPath)
      expect(xml).toContain('<w:smallCaps/>')
      expect(xml).toContain('<w:spacing w:val="40"/>')
      expect(xml).toContain('<w:sz w:val="24"/>')

      // reopen the dialog: the fields show the saved values
      await editorPage.locator('.doc-page').click()
      await editorPage.keyboard.press('ControlOrMeta+a')
      await editorPage.keyboard.press('ControlOrMeta+d')
      await expect(dialog).toBeVisible()
      await expect(dialog.getByLabel('Small caps')).toBeChecked()
      await dialog.getByRole('tab', { name: 'Advanced' }).click()
      await expect(dialog.getByLabel('Scale (%)')).toHaveValue('150')
      await expect(dialog.getByRole('button', { name: 'Spacing', exact: true })).toContainText(
        'Expanded',
      )
      await expect(dialog.getByLabel('By (pt)').first()).toHaveValue('2')
      await dialog.getByRole('button', { name: 'Cancel' }).click()
    } finally {
      await closeAndSaveVideo(launched, 'docs-font-dialog-advanced')
    }
  })
})
