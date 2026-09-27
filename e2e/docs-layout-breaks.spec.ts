import { test, expect } from '@playwright/test'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * Layout > Breaks follows Word: Page / Column / Text Wrapping under "Page
 * Breaks" and the four section breaks below. Each item lands as the matching
 * OOXML on save; Insert > Cover Page offers "Remove Current Cover Page" only
 * while a gallery cover exists and removing it leaves the body untouched.
 */

interface AidocsWindow {
  __aidocs?: { editor?: unknown; save?: () => Promise<unknown> }
}

async function minimalDocx(): Promise<Buffer> {
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
    '<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Alpha paragraph</w:t></w:r></w:p><w:p><w:r><w:t>Beta paragraph</w:t></w:r></w:p><w:p><w:r><w:t>Gamma paragraph</w:t></w:r></w:p><w:p><w:r><w:t>Delta paragraph</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>',
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function documentXml(path: string): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(path))
  return zip.file('word/document.xml')!.async('string')
}

test.describe('docs layout breaks menu', () => {
  let dir: string
  let docPath: string

  test.beforeEach(async () => {
    dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-breaks-')))
    docPath = join(dir, 'breaks.docx')
    writeFileSync(docPath, await minimalDocx())
  })

  test.afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  test('page, column, text wrapping and section breaks save as OOXML; cover page removal', async () => {
    test.setTimeout(180_000)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-layout-breaks',
      openFile: docPath,
    })
    const { app } = launched
    try {
      const page = await waitForPageWithUrl(app, '://docs/')
      await page.waitForFunction(
        () => Boolean((window as unknown as AidocsWindow).__aidocs?.editor),
        undefined,
        { timeout: 30_000 },
      )
      const para = (text: string) => page.locator('.doc-page p', { hasText: text }).first()
      const menuItem = (label: string) =>
        page
          .locator('[data-rb-panel] button')
          .filter({ has: page.locator('b', { hasText: new RegExp(`^${label}$`) }) })
      const openBreaks = async () => {
        await page.getByRole('button', { name: 'Breaks', exact: true }).click()
        await expect(page.locator('.layout-menu-head')).toHaveText([
          'Page Breaks',
          'Section Breaks',
        ])
      }

      await page.locator('.ribbon-tab', { hasText: /^Layout$/ }).click()

      await para('Alpha').click()
      await page.keyboard.press('End')
      await openBreaks()
      await menuItem('Continuous').click()

      await para('Beta').click()
      await page.keyboard.press('End')
      await openBreaks()
      await menuItem('Column').click()
      await expect(page.locator('.doc-page br.doc-col-br')).toHaveCount(1)

      await para('Gamma').click()
      await page.keyboard.press('End')
      await openBreaks()
      await menuItem('Text Wrapping').click()
      await expect(page.locator('.doc-page br.doc-wrap-br')).toHaveCount(1)

      await para('Delta').click()
      await page.keyboard.press('End')
      await openBreaks()
      await menuItem('Page').click()

      await page.getByRole('button', { name: 'Margins', exact: true }).click()
      await menuItem('Mirrored').click()

      await page.getByRole('button', { name: 'Size', exact: true }).click()
      await expect(page.locator('[data-rb-panel] button b')).toHaveText([
        'Letter',
        'Legal',
        'Executive',
        'A3',
        'A4',
        'A5',
        'B4 (JIS)',
        'B5 (JIS)',
        'Tabloid',
        'Statement',
        'Envelope #10',
        'Envelope DL',
        'Envelope C5',
        'More Paper Sizes…',
      ])
      await menuItem('More Paper Sizes…').click()
      const dialog = page.locator('.modal', { hasText: 'Paper Size' })
      await dialog.getByLabel(/^Width/).fill('28 cm')
      await dialog.getByLabel(/^Height/).fill('20 cm')
      await dialog.getByLabel('Apply to').selectOption('document')
      await dialog.getByRole('button', { name: 'OK' }).click()

      expect(await page.evaluate(() => (window as unknown as AidocsWindow).__aidocs!.save!())).toBe(
        true,
      )
      await expect
        .poll(() => documentXml(docPath), { timeout: 15_000 })
        .toContain('<w:br w:type="textWrapping" w:clear="all"/>')
      const xml = await documentXml(docPath)
      expect(xml).toContain('<w:br w:type="column"/>')
      expect(xml).toContain('<w:pageBreakBefore/>')
      expect(xml.match(/<w:sectPr>/g)).toHaveLength(2)
      expect(xml).toContain('<w:type w:val="continuous"/>')
      // width > height flips every section to landscape (Word's Paper tab behaviour)
      expect(xml.match(/<w:pgSz w:w="15874" w:h="11339" w:orient="landscape"\/>/g)).toHaveLength(2)
      expect(xml).toMatch(/<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1800"/)
      const settingsZip = await JSZip.loadAsync(readFileSync(docPath))
      expect(await settingsZip.file('word/settings.xml')!.async('string')).toContain(
        '<w:mirrorMargins/>',
      )

      await page.locator('.ribbon-tab', { hasText: /^Insert$/ }).click()
      const coverButton = page.getByRole('button', { name: 'Cover Page', exact: true })
      await coverButton.click()
      const remove = page.locator('.cover-gallery-remove')
      await expect(remove).toBeDisabled()
      await page.locator('.cover-card').first().click()
      await expect(page.locator('.doc-page p').first()).not.toHaveText(/Alpha/)

      await coverButton.click()
      await expect(remove).toBeEnabled()
      await remove.click()
      await expect(page.locator('.doc-page p').first()).toHaveText(/Alpha/)
      await expect(page.locator('.doc-page br.doc-col-br')).toHaveCount(1)
    } finally {
      await closeAndSaveVideo(launched, 'docs-layout-breaks')
    }
  })
})
