import { test, expect } from '@playwright/test'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, screenshotPath, waitForPageWithUrl } from './helpers'

/**
 * Table operations wave: Split Cells… (2×3 on a merged cell), Split Table, the
 * 9-way cell alignment, a double border from the Table Design style picker and
 * alt text from Table Properties all render and save as the matching OOXML.
 */

interface AidocsWindow {
  __aidocs?: { editor?: unknown; save?: () => Promise<unknown> }
}

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
  return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${body}</w:tbl>`
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
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Before</w:t></w:r></w:p>${tableXml(3, 3)}<w:p><w:r><w:t>After</w:t></w:r></w:p></w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function documentXml(path: string): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(path))
  return zip.file('word/document.xml')!.async('string')
}

test.describe('docs table operations', () => {
  let dir: string
  let docPath: string

  test.beforeEach(async () => {
    dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-table-')))
    docPath = join(dir, 'grid.docx')
    writeFileSync(docPath, await minimalDocx())
  })

  test.afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  test('split cells, split table, cell alignment, border style and alt text save as OOXML', async () => {
    test.setTimeout(180_000)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-table-ops',
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
      const cell = (row: number, col: number) =>
        editorPage
          .locator('.doc-page .doc-table')
          .first()
          .locator('tr')
          .nth(row)
          .locator('td, th')
          .nth(col)

      // ---- Split Cells… on a merged A1:B1 → 2 rows × 3 columns
      await cell(0, 0).hover()
      await editorPage.mouse.down()
      await cell(0, 1).hover()
      await editorPage.mouse.up()
      // an existing table only shows its contextual tabs (Word); pick Layout explicitly
      await editorPage.locator('.ribbon-tab', { hasText: 'Table Layout' }).click()
      await editorPage.getByRole('button', { name: 'Merge Cells' }).click()
      await cell(0, 0).click()
      await editorPage.getByRole('button', { name: 'Split Cells', exact: true }).click()
      const split = editorPage.locator('.modal.split-cells-dialog')
      await expect(split).toBeVisible()
      await split.getByLabel('Number of columns').fill('3')
      await split.getByLabel('Number of rows').fill('2')
      await split.getByRole('button', { name: 'OK' }).click()
      await expect(split).toBeHidden()
      // the merged pair's two grid columns become three (grid 3 → 4); rows 3 → 4
      await expect(editorPage.locator('.doc-page .doc-table').first().locator('tr')).toHaveCount(4)
      await expect(cell(0, 0)).toHaveText('A1B1')

      // ---- 9-way alignment on C1: bottom right
      await cell(0, 3).click()
      await editorPage.getByRole('button', { name: 'Align Bottom Right' }).click()
      await expect(cell(0, 3)).toHaveCSS('vertical-align', 'bottom')
      await expect(cell(0, 3).locator('p').first()).toHaveCSS('text-align', 'right')

      // ---- double border via Table Design: Border Styles, then Borders ▾ › All Borders
      await editorPage.locator('.ribbon-tab', { hasText: 'Table Design' }).click()
      await editorPage.getByRole('button', { name: 'Border Styles' }).click()
      await editorPage.getByRole('option', { name: '═════' }).click()
      await editorPage.locator('.table-split-caret').click()
      await editorPage.locator('.layout-menu button', { hasText: 'All Borders' }).click()
      await expect(cell(0, 3)).toHaveCSS('border-top-style', 'double')

      // ---- alt text through Table Properties
      await editorPage.locator('.ribbon-tab', { hasText: 'Table Layout' }).click()
      await editorPage.getByRole('button', { name: 'Table Properties' }).click()
      const props = editorPage.locator('.modal.table-properties-dialog')
      await expect(props).toBeVisible()
      await props.getByRole('tab', { name: 'Row' }).click()
      await props.getByLabel('Allow row to break across pages').uncheck()
      await props.getByRole('tab', { name: 'Alt Text' }).click()
      await props.getByLabel('Title').fill('Quarterly grid')
      await props.getByLabel('Description').fill('Three quarters of sample cells')
      await props.screenshot({ path: screenshotPath('docs-table-ops-properties') })
      await props.getByRole('button', { name: 'OK' }).click()
      await expect(props).toBeHidden()

      // ---- Split Table at the last row
      await cell(3, 0).click()
      await editorPage.getByRole('button', { name: 'Split Table' }).click()
      await expect(editorPage.locator('.doc-page .doc-table')).toHaveCount(2)

      expect(
        await editorPage.evaluate(() => (window as unknown as AidocsWindow).__aidocs!.save!()),
      ).toBe(true)
      await expect
        .poll(() => documentXml(docPath), { timeout: 15_000 })
        .toContain('<w:tblDescription w:val="Three quarters of sample cells"/>')
      const xml = await documentXml(docPath)
      expect(xml.match(/<w:tbl>/g)).toHaveLength(2)
      expect(xml).toContain('<w:tblCaption w:val="Quarterly grid"/>')
      expect(xml).toContain('<w:vAlign w:val="bottom"/>')
      expect(xml).toContain('<w:jc w:val="right"/>')
      expect(xml).toMatch(/<w:tcBorders><w:top w:val="double"/)
      expect(xml).toContain('<w:cantSplit/>')
      // the merged pair's width is shared by the three split cells, so the row below spans two grid columns
      expect(xml).toContain('<w:gridSpan w:val="2"/>')
    } finally {
      await closeAndSaveVideo(launched, 'docs-table-ops')
    }
  })
})
