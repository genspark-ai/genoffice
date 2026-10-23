/**
 * Word's bottom-right table handle: dragging it scales the whole table, and
 * the new width survives the commit even for an AutoFit-to-window table that
 * used to spring back to the full content width.
 */
import { test, expect } from '@playwright/test'
import { mkdtempSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

function tableXml(): string {
  const cell = (t: string) => `<w:tc><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`
  const rows = [
    ['A', 'B', 'C'],
    ['D', 'E', 'F'],
  ]
    .map((r) => `<w:tr>${r.map(cell).join('')}</w:tr>`)
    .join('')
  const border = (n: string) => `<w:${n} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`
  const borders = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('')
  // 100% preferred width + autofit layout = Word's "AutoFit to Window"
  return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="5000" w:type="pct"/><w:tblLayout w:type="autofit"/><w:tblBorders>${borders}</w:tblBorders></w:tblPr><w:tblGrid>${'<w:gridCol w:w="3120"/>'.repeat(3)}</w:tblGrid>${rows}</w:tbl>`
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
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${p('before')}${tableXml()}${p('after')}</w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

test('corner handle drag scales width and height and both stick', async () => {
  test.setTimeout(150_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-corner-resize-')))
  const docPath = join(dir, 'table.docx')
  writeFileSync(docPath, await docx())
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-table-corner-resize',
    openFile: docPath,
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    const table = page.locator('.doc-page table.doc-table').first()
    await table.waitFor()
    await page.waitForTimeout(1000)

    const tableWidth = () =>
      page.evaluate(() => {
        const t = document.querySelector('.doc-page table.doc-table') as HTMLTableElement
        return Math.round(t.getBoundingClientRect().width)
      })
    const columnWidths = () =>
      page.evaluate(() => {
        const t = document.querySelector('.doc-page table.doc-table') as HTMLTableElement
        return Array.from(t.rows[0].cells).map((c) => c.getBoundingClientRect().width)
      })
    const tableHeight = () =>
      page.evaluate(() => {
        const t = document.querySelector('.doc-page table.doc-table') as HTMLTableElement
        return Math.round(t.getBoundingClientRect().height)
      })
    const before = await tableWidth()
    const heightBefore = await tableHeight()
    const columnsBefore = await columnWidths()
    const cell = await table.locator('td').first().boundingBox()
    if (!cell) throw new Error('no first cell')
    await page.mouse.move(cell.x + cell.width / 2, cell.y + cell.height / 2)
    const handle = page.locator('.doc-table-resize')
    await expect(handle).toBeVisible()
    const box = await handle.boundingBox()
    if (!box) throw new Error('no resize handle')
    const x = box.x + box.width / 2
    const y = box.y + box.height / 2
    await page.mouse.move(x, y)
    await page.mouse.down()
    for (let step = 1; step <= 10; step++) {
      // left and down: narrower but taller, like Word's corner handle
      await page.mouse.move(x - step * 20, y + step * 6)
      await page.waitForTimeout(30)
    }
    await expect.poll(tableWidth, { timeout: 3_000 }).toBeLessThan(before - 150)
    await expect.poll(tableHeight, { timeout: 3_000 }).toBeGreaterThan(heightBefore + 40)
    await page.mouse.up()
    await page.waitForTimeout(500)
    const after = await tableWidth()
    expect(after).toBeLessThan(before - 150)
    // the committed w:trHeight paints the dragged height, within a pixel per row
    expect(Math.abs((await tableHeight()) - (heightBefore + 60))).toBeLessThan(4)
    await page.screenshot({ path: join(dir, '..', 'corner-resize-after.png') })
    // every column scaled by the same factor (the content-autofit columns of
    // 'A' / 'B' / 'C' start out slightly unequal, so compare ratios, not widths)
    const widths = await columnWidths()
    const scale = after / before
    widths.forEach((w, i) => expect(Math.abs(w - columnsBefore[i] * scale)).toBeLessThan(3))
  } finally {
    await closeAndSaveVideo(launched)
    rmSync(dir, { recursive: true, force: true })
  }
})
