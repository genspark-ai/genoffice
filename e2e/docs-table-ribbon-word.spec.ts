import { test, expect } from '@playwright/test'
import { mkdtempSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * Table contextual tabs follow Word for Mac: Table Layout groups run
 * Table · Rows & Columns · Merge · Cell Size · Alignment · Data, Select ▾
 * picks a row, the Borders ▾ split button applies its default Bottom Border to
 * the whole table when nothing is selected, Shading uses the shared theme
 * palette and Cell Margins writes the table's default cell margins.
 */

interface AidocsWindow {
  __aidocs?: { editor?: { getAttributes: (name: string) => Record<string, unknown> } }
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
  return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${body}</w:tbl>`
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

test.describe('docs table ribbon (Word layout)', () => {
  let dir: string
  let docPath: string

  test.beforeEach(async () => {
    dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-table-ribbon-')))
    docPath = join(dir, 'grid.docx')
    writeFileSync(docPath, await minimalDocx())
  })

  test.afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  test('group order, Select ▾ row, Borders ▾ default, Shading palette, Cell Margins', async () => {
    test.setTimeout(180_000)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-table-ribbon-word',
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
      const table = page.locator('.doc-page .doc-table').first()
      const cell = (row: number, col: number) =>
        table.locator('tr').nth(row).locator('td, th').nth(col)

      const tab = (name: string) => page.locator('.ribbon-tab', { hasText: name })
      await page.locator('.doc-page p', { hasText: 'Before' }).first().click()
      await expect(page.locator('.ribbon-tab.active')).toHaveText('Home')
      await cell(1, 1).click()
      await expect(tab('Table Layout')).toBeVisible()
      await tab('Table Layout').click()
      await expect(page.locator('.table-ribbon-body .ribbon-group-label')).toHaveText([
        'Table',
        'Rows & Columns',
        'Merge',
        'Cell Size',
        'Alignment',
        'Data',
      ])

      // ---- Select ▾ › Select Row highlights the caret's row
      await page.getByRole('button', { name: 'Select', exact: true }).click()
      await page.locator('.layout-menu button', { hasText: 'Select Row' }).click()
      await expect(table.locator('.selectedCell')).toHaveCount(3)

      // ---- Select Table keeps the contextual tabs (Word) and later commands still run
      await page.getByRole('button', { name: 'Select', exact: true }).click()
      await page.locator('.layout-menu button', { hasText: 'Select Table' }).click()
      await expect(page.locator('.ribbon-tab.active')).toHaveText('Table Layout')
      await page.getByRole('button', { name: 'Insert Below' }).click()
      await expect(table.locator('tr')).toHaveCount(4)

      // ---- Borders ▾ main half = Bottom Border; a bare caret formats the whole table
      await cell(1, 1).click()
      await tab('Table Design').click()
      const bordersMain = page.locator('.table-split-main')
      await expect(bordersMain).toHaveText('Borders')
      await bordersMain.click()
      await expect(cell(3, 0)).toHaveCSS('border-bottom-style', 'solid')
      await expect(cell(3, 2)).toHaveCSS('border-bottom-style', 'solid')
      await expect(cell(0, 0)).not.toHaveCSS('border-bottom-style', 'solid')

      // ---- the gallery lists Word's entries in Word's order
      await page.locator('.table-split-caret').click()
      await expect(page.locator('.layout-menu button')).toHaveText([
        'Bottom Border',
        'Top Border',
        'Left Border',
        'Right Border',
        'No Border',
        'All Borders',
        'Outside Borders',
        'Inside Borders',
        'Inside Horizontal Border',
        'Inside Vertical Border',
      ])
      await page.keyboard.press('Escape')

      // ---- Shading ▾ is the shared theme palette
      await page.getByRole('button', { name: 'Shading' }).click()
      const palette = page.locator('.gcp-palette')
      await expect(palette.locator('.gcp-auto')).toHaveText('No Color')
      await palette.locator('.gcp-swatch[title="Yellow"]').click()
      await expect(cell(1, 1)).toHaveCSS('background-color', 'rgb(255, 255, 0)')

      // ---- Cell Margins opens Table Options and writes the table's default margins
      await tab('Table Layout').click()
      await page.getByRole('button', { name: 'Cell Margins' }).click()
      const options = page.locator('.modal.cell-margins-dialog')
      await expect(options).toBeVisible()
      await options.getByLabel('Top').fill('0.5 cm')
      await options.getByRole('button', { name: 'OK' }).click()
      await expect(options).toBeHidden()
      const cellMar = await page.evaluate(
        () =>
          (window as unknown as AidocsWindow).__aidocs!.editor!.getAttributes('docTable')
            .cellMar as Record<string, number> | null,
      )
      // 0.5 cm is 283.46 twips: the unit fields convert exactly
      expect(cellMar?.top).toBe(283)
    } finally {
      await closeAndSaveVideo(launched, 'docs-table-ribbon-word')
    }
  })
})
