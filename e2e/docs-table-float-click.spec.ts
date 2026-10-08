import { test, expect } from '@playwright/test'
import { mkdtempSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * A floating table (w:tblpPr) is followed by a positioned paragraph whose
 * full-width box used to paint over the float and take every click meant for
 * a cell. Clicking a cell must place the caret inside that cell.
 */

interface AidocsWindow {
  __aidocs?: {
    editor?: {
      state: {
        selection: { $from: { depth: number; node(d: number): { type: { name: string } } } }
      }
    }
  }
}

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

function cell(text: string): string {
  return `<w:tc><w:tcPr><w:tcW w:w="2200" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`
}

function floatingTable(): string {
  const row = (cells: string[]) =>
    `<w:tr><w:trPr><w:trHeight w:val="900" w:hRule="atLeast"/></w:trPr>${cells.map(cell).join('')}</w:tr>`
  return (
    `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/>` +
    `<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="text" w:horzAnchor="page" w:tblpX="1842" w:tblpY="1"/>` +
    `<w:tblOverlap w:val="never"/><w:tblW w:w="8800" w:type="dxa"/>` +
    `<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="2200"/><w:gridCol w:w="2200"/><w:gridCol w:w="2200"/><w:gridCol w:w="2200"/></w:tblGrid>` +
    row(['Name', 'Gender', 'D.O.B.', 'Passport No.']) +
    row(['Alice', 'F', '01/01/1990', 'X0000000']) +
    `</w:tbl>`
  )
}

async function docxWithFloatingTable(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  const spaced = (text: string) =>
    `<w:p><w:pPr><w:spacing w:before="180" w:after="180"/></w:pPr><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body>${spaced('The following are the details')}${floatingTable()}${spaced(' ')}${spaced('Signature:')}</w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

test.describe('docs floating table click', () => {
  let dir: string
  let docPath: string

  test.beforeEach(async () => {
    dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-tblp-')))
    docPath = join(dir, 'floating.docx')
    writeFileSync(docPath, await docxWithFloatingTable())
  })

  test.afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  test('clicking a cell of a floating table puts the caret in that cell', async () => {
    test.setTimeout(120_000)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-table-float-click',
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
      const table = page.locator('.doc-page .doc-table.doc-table-float-left').first()
      await expect(table).toBeVisible()
      const nameCell = table.locator('tr').nth(1).locator('td, th').nth(0)
      await nameCell.scrollIntoViewIfNeeded()
      const box = (await nameCell.boundingBox())!
      // the empty top band of the cell, well away from the glyphs
      await page.mouse.click(box.x + box.width / 2, box.y + 6)
      const path = await page.evaluate(() => {
        const $from = (window as unknown as AidocsWindow).__aidocs!.editor!.state.selection.$from
        const names: string[] = []
        for (let d = $from.depth; d >= 0; d--) names.push($from.node(d).type.name)
        return names
      })
      expect(path).toContain('docTableCell')
      await page.keyboard.type('Z')
      await expect(nameCell).toHaveText(/Z/)
    } finally {
      await closeAndSaveVideo(launched, 'docs-table-float-click').catch(() => {})
    }
  })
})
