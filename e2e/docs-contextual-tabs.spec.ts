import { test, expect } from '@playwright/test'
import { mkdtempSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * Contextual ribbon tabs follow Word for Mac: the caret entering an existing
 * table only shows Table Design / Table Layout, a freshly inserted table
 * activates Table Design, and a picture activates Picture Format on click
 * while a double-click keeps that selection instead of opening the viewer.
 */

interface AidocsWindow {
  __aidocs?: { editor?: { chain: () => unknown } }
}

const PNG_1X1 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='

function tableXml(rows: number, cols: number): string {
  const grid = Array.from({ length: cols }, () => '<w:gridCol w:w="2400"/>').join('')
  const body = Array.from(
    { length: rows },
    (_, r) =>
      `<w:tr>${Array.from(
        { length: cols },
        (_, c) =>
          `<w:tc><w:p><w:r><w:t>${String.fromCharCode(65 + c)}${r + 1}</w:t></w:r></w:p></w:tc>`,
      ).join('')}</w:tr>`,
  ).join('')
  return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${body}</w:tbl>`
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
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Before</w:t></w:r></w:p>${tableXml(2, 2)}<w:p><w:r><w:t>After</w:t></w:r></w:p></w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

test.describe('docs contextual ribbon tabs', () => {
  let dir: string
  let docPath: string

  test.beforeEach(async () => {
    dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-ctxtabs-')))
    docPath = join(dir, 'ctx.docx')
    writeFileSync(docPath, await minimalDocx())
  })

  test.afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  test('existing table keeps the tab, new table and pictures activate theirs', async () => {
    test.setTimeout(180_000)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-contextual-tabs',
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
      const activeTab = page.locator('.ribbon-tab.active')
      const tab = (name: string) => page.locator('.ribbon-tab', { hasText: name })
      const before = page.locator('.doc-page p', { hasText: 'Before' }).first()
      const after = page.locator('.doc-page p', { hasText: 'After' }).first()
      const cell = page.locator('.doc-page .doc-table').first().locator('td, th').first()

      await before.click()
      await page.keyboard.type('x')
      await expect(activeTab).toHaveText('Home')

      await cell.click()
      await expect(tab('Table Layout')).toBeVisible()
      await expect(tab('Table Design')).toBeVisible()
      await expect(activeTab).toHaveText('Home')

      await tab('Table Layout').click()
      await expect(activeTab).toHaveText('Table Layout')
      await after.click()
      await expect(tab('Table Layout')).toBeHidden()
      await expect(activeTab).toHaveText('Home')

      await tab('Insert').click()
      await page.getByRole('button', { name: 'Table', exact: true }).click()
      await page.locator('.table-picker-grid .table-cell').nth(11).click()
      await expect(activeTab).toHaveText('Table Design')
      await expect(page.locator('.doc-page .doc-table')).toHaveCount(2)

      await after.click()
      await expect(activeTab).toHaveText('Insert')

      await page.evaluate((png) => {
        const editor = (window as unknown as AidocsWindow).__aidocs!.editor as {
          chain: () => {
            focus: () => { insertContent: (c: unknown) => { run: () => void } }
          }
        }
        editor
          .chain()
          .focus()
          .insertContent({
            type: 'docProtected',
            attrs: {
              docxIndex: null,
              blockType: 'image',
              label: 'Picture',
              imageDataUrl: `data:image/png;base64,${png}`,
              imageWidthPx: 120,
              imageHeightPx: 120,
              genImage: { base64: png, mime: 'image/png', widthPx: 120, heightPx: 120 },
            },
          })
          .run()
      }, PNG_1X1)
      const img = page.locator('.doc-page .doc-img-wrap img').first()
      await expect(img).toBeVisible()

      await after.click()
      await img.click()
      await expect(activeTab).toHaveText('Picture Format')
      await img.dblclick()
      await expect(activeTab).toHaveText('Picture Format')
      await expect(page.locator('.gs-imgview-mask')).toHaveCount(0)
      await page.waitForTimeout(300)
      await expect(activeTab).toHaveText('Picture Format')
      await expect(page.locator('.gs-imgview-mask')).toHaveCount(0)

      await img.click({ button: 'right' })
      await page.getByText('View Image', { exact: true }).click()
      await expect(page.locator('.gs-imgview-mask')).toBeVisible()
      await page.keyboard.press('Escape')
      await expect(page.locator('.gs-imgview-mask')).toHaveCount(0)

      await after.click()
      await expect(tab('Picture Format')).toBeHidden()
      await expect(activeTab).toHaveText('Insert')
    } finally {
      await closeAndSaveVideo(launched, 'docs-contextual-tabs')
    }
  })
})
