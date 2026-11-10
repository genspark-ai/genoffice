/**
 * Dragging a column border previews the new grid while the mouse is still
 * down: the dragged cell widens and the text reflows before mouseup, for
 * tables with and without a saved tblGrid (genoffice#1156).
 */
import { test, expect } from '@playwright/test'
import { mkdtempSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

const LONG = 'The quick brown fox jumps over the lazy dog and keeps running across the wide field. '

function tableXml(withGrid: boolean): string {
  const grid = withGrid ? `<w:tblGrid>${'<w:gridCol w:w="3000"/>'.repeat(3)}</w:tblGrid>` : ''
  const cell = (t: string) =>
    `<w:tc><w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p></w:tc>`
  const rows = [
    ['A', 'B', 'C'],
    [LONG, LONG, LONG],
  ]
    .map((r) => `<w:tr>${r.map(cell).join('')}</w:tr>`)
    .join('')
  const border = (n: string) => `<w:${n} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`
  const borders = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('')
  return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblBorders>${borders}</w:tblBorders></w:tblPr>${grid}${rows}</w:tbl>`
}

async function docx(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  const p = (t: string) => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${p('grid')}${tableXml(true)}${p('no grid')}${tableXml(false)}${p('end')}</w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

test('column border drag previews the width before mouseup', async () => {
  test.setTimeout(150_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-col-resize-')))
  const docPath = join(dir, 'tables.docx')
  writeFileSync(docPath, await docx())
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-table-resize-live',
    openFile: docPath,
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    const tables = page.locator('.doc-page table.doc-table')
    await tables.nth(1).waitFor()
    await page.waitForTimeout(1000)

    for (const idx of [0, 1]) {
      const firstCellWidth = () =>
        page.evaluate((i) => {
          const t = document.querySelectorAll('.doc-page table.doc-table')[i] as HTMLTableElement
          return Math.round(t.rows[0].cells[0].getBoundingClientRect().width)
        }, idx)
      const before = await firstCellWidth()
      const cell = await tables.nth(idx).locator('tr').first().locator('td').first().boundingBox()
      if (!cell) throw new Error('no first cell')
      const x = cell.x + cell.width - 1
      const y = cell.y + cell.height / 2
      await page.mouse.move(x, y)
      await page.waitForTimeout(200)
      await page.mouse.down()
      for (let step = 1; step <= 8; step++) {
        await page.mouse.move(x + step * 10, y)
        await page.waitForTimeout(30)
      }
      await expect.poll(firstCellWidth, { timeout: 3_000 }).toBeGreaterThan(before + 40)
      await page.mouse.up()
      await page.waitForTimeout(300)
      expect(await firstCellWidth()).toBeGreaterThan(before + 20)
    }
  } finally {
    await closeAndSaveVideo(launched)
    rmSync(dir, { recursive: true, force: true })
  }
})
