/**
 * genoffice#2488: with formatting marks on, Word draws a dashed "Page Break"
 * chip at a w:br page break, an arrow at a text-wrapping break and a
 * "Section Break (type)" chip on the last line of every section-ending
 * paragraph; the editor showed only the pilcrows.
 */
import { test, expect } from '@playwright/test'
import { mkdtempSync, realpathSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

const sect = (type?: string) =>
  `<w:sectPr>${type ? `<w:type w:val="${type}"/>` : ''}<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="964" w:bottom="1418" w:left="1474" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>`

const para = (text: string, inner = '', pPr = '') =>
  `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r><w:t>${text}</w:t></w:r>${inner}</w:p>`

async function breaksDocx(): Promise<Buffer> {
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
      para('Page break:', '<w:r><w:br w:type="page"/></w:r>') +
      para('Wrap break:', '<w:r><w:br w:type="textWrapping" w:clear="all"/></w:r>') +
      para('Continuous:', '', sect()) +
      para('Even page:', '', sect('continuous')) +
      para('Tail', '', sect('evenPage')) +
      sect('oddPage') +
      `</w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

test.describe('formatting marks show breaks', () => {
  test('page, wrap and section breaks get Word-style marks only while marks are on', async () => {
    test.setTimeout(240_000)
    const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-fmt-marks-')))
    const docPath = join(dir, 'breaks.docx')
    writeFileSync(docPath, await breaksDocx())

    const app = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-formatting-marks-breaks',
      openFile: docPath,
    })
    const page = await waitForPageWithUrl(app.app, '://docs/')
    try {
      await page.locator('.doc-page p', { hasText: 'Page break:' }).first().waitFor()
      await expect(page.locator('.doc-fmt-mark')).toHaveCount(0)

      await page.getByRole('button', { name: 'Show/hide formatting marks' }).click()
      await expect(page.locator('.doc-fmt-pagebr')).toHaveText('Page Break')
      await expect(page.locator('.doc-fmt-wrapbr')).toHaveCount(1)
      const sectChips = page.locator('.doc-fmt-sectbr')
      await expect(sectChips).toHaveCount(3)
      await expect(sectChips.nth(0)).toHaveText('Section Break (Continuous)')
      await expect(sectChips.nth(1)).toHaveText('Section Break (Even Page)')
      await expect(sectChips.nth(2)).toHaveText('Section Break (Odd Page)')
      // the chip replaces the paragraph's pilcrow, as in Word
      const pilcrow = await page
        .locator('.doc-page p.doc-sect-break-mark')
        .first()
        .evaluate((el) => getComputedStyle(el, '::after').content)
      expect(pilcrow).toBe('none')

      await page.getByRole('button', { name: 'Show/hide formatting marks' }).click()
      await expect(page.locator('.doc-fmt-mark')).toHaveCount(0)
    } finally {
      await closeAndSaveVideo(app, 'docs-formatting-marks-breaks')
    }
  })
})
