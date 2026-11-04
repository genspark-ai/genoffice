/**
 * Word-style continuous text selection across table cells (genoffice#2329):
 * dragging from one cell into another extends the text selection across the
 * boundary instead of jumping to a cell-selection rectangle, and deleting or
 * typing over it clears the covered text while every cell it crossed survives.
 */
import { test, expect } from '@playwright/test'
import { mkdtempSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

const WORDS = ['Alpha', 'Bravo', 'Charlie', 'Delta']

function docx(): Promise<Buffer> {
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
  const cell = (t: string) => `<w:tc><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`
  const rows = [WORDS.slice(0, 2), WORDS.slice(2)]
    .map((r) => `<w:tr>${r.map(cell).join('')}</w:tr>`)
    .join('')
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${p('before')}<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>${rows}</w:tbl>${p('after')}</w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

/** the four cell texts in reading order */
const cellTexts = (page: import('@playwright/test').Page): Promise<string[]> =>
  page.evaluate(() =>
    Array.from(
      document.querySelectorAll('.doc-page table.doc-table td, .doc-page table.doc-table th'),
    ).map((c) => (c.textContent ?? '').replace(/\u200b/g, '').trim()),
  )

interface DragProbe {
  selected: string
  startCell: number
  endCell: number
}

async function dragAcrossDiagonal(page: import('@playwright/test').Page): Promise<void> {
  const table = page.locator('.doc-page table.doc-table').first()
  await table.waitFor()
  await page.waitForTimeout(800)
  const first = await table.locator('td').first().boundingBox()
  const last = await table.locator('td').last().boundingBox()
  if (!first || !last) throw new Error('missing cells')
  // start inside the first cell's text, end inside the last cell's text
  await page.mouse.move(first.x + 5, first.y + 5)
  await page.mouse.down()
  for (let step = 1; step <= 12; step++) {
    await page.mouse.move(
      first.x + 5 + ((last.x + 5 - (first.x + 5)) * step) / 12,
      first.y + 5 + ((last.y + 5 - (first.y + 5)) * step) / 12,
    )
    await page.waitForTimeout(30)
  }
  await page.mouse.up()
  await page.waitForTimeout(300)
}

async function dragProbe(page: import('@playwright/test').Page): Promise<DragProbe> {
  return page.evaluate(() => {
    const sel = window.getSelection()
    const cells = Array.from(
      document.querySelectorAll('.doc-page table.doc-table td, .doc-page table.doc-table th'),
    )
    const cellIndex = (node: Node | null): number => {
      for (let el = node; el; el = el.parentNode) {
        const at = cells.indexOf(el)
        if (at !== -1) return at
      }
      return -1
    }
    return {
      selected: sel?.toString() ?? '',
      startCell: cellIndex(sel?.rangeCount ? sel.getRangeAt(0).startContainer : null),
      endCell: cellIndex(sel?.rangeCount ? sel.getRangeAt(0).endContainer : null),
    }
  })
}

test('a drag from one cell into another selects text across them', async () => {
  test.setTimeout(150_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-cross-cell-drag-')))
  const docPath = join(dir, 'table.docx')
  writeFileSync(docPath, await docx())
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-table-cross-cell-drag',
    openFile: docPath,
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    await dragAcrossDiagonal(page)
    const probe = await dragProbe(page)
    // continuous text across the boundary, anchored in the first cell and
    // ending in the last — a cell-selection rectangle would leave the DOM
    // selection collapsed inside one cell
    expect(probe.selected.length).toBeGreaterThan(WORDS[0].length)
    expect(probe.startCell).toBe(0)
    expect(probe.endCell).toBe(WORDS.length - 1)
    expect(probe.selected).toContain('Alpha')
    expect(probe.selected).toContain('Delta')
    await page.screenshot({ path: join(dir, '..', 'cross-cell-selected.png') })

    await page.keyboard.press('Backspace')
    await page.waitForTimeout(300)
    const texts = await cellTexts(page)
    // Word's delete: the covered text goes, every cell it crossed stays
    expect(texts).toHaveLength(4)
    expect(texts[0]).not.toBe(WORDS[0])
    expect(texts[3]).not.toBe(WORDS[3])
    expect(texts[1]).toBe('')
    expect(texts[2]).toBe('')
    await page.screenshot({ path: join(dir, '..', 'cross-cell-deleted.png') })
  } finally {
    await closeAndSaveVideo(launched)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('typing over a cross-cell selection replaces it in the first cell', async () => {
  test.setTimeout(150_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-cross-cell-type-')))
  const docPath = join(dir, 'table.docx')
  writeFileSync(docPath, await docx())
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-table-cross-cell-type',
    openFile: docPath,
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    await dragAcrossDiagonal(page)
    await page.keyboard.type('X')
    await page.waitForTimeout(300)
    const texts = await cellTexts(page)
    expect(texts).toHaveLength(4)
    expect(texts[0].startsWith('X')).toBe(true)
    expect(texts[3]).not.toBe(WORDS[3])
    await page.screenshot({ path: join(dir, '..', 'cross-cell-typed.png') })
  } finally {
    await closeAndSaveVideo(launched)
    rmSync(dir, { recursive: true, force: true })
  }
})
