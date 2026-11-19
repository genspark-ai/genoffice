import { test, expect } from '@playwright/test'
import { mkdtempSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import type { Locator, Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * Ruler indent and margin handles (Word): dragging the first-line, hanging,
 * left and right markers writes the paragraph's w:ind as one undo step,
 * snapped to 1/16 in unless Option is held; dragging the margin boundary
 * changes the section's page margins; in a table the markers measure from
 * the cell and the margin handles hide.
 */

interface AidocsWindow {
  __aidocs?: { editor?: unknown }
}

const PAGE_WIDTH_TWIPS = 12240

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
  const para = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
  const table =
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Cell B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${para(
      'First paragraph with enough words to wrap onto a second line when it is indented a lot on the ruler by the tester.',
    )}${para('Second paragraph.')}${table}${para('After the table.')}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function dragBy(page: Page, handle: Locator, dx: number, alt = false): Promise<void> {
  const box = (await handle.boundingBox())!
  const x = box.x + box.width / 2
  const y = box.y + box.height / 2
  await page.mouse.move(x, y)
  await page.mouse.down()
  if (alt) await page.keyboard.down('Alt')
  await page.mouse.move(x + dx / 2, y, { steps: 3 })
  await page.mouse.move(x + dx, y, { steps: 3 })
  await page.mouse.up()
  if (alt) await page.keyboard.up('Alt')
}

test.describe('docs ruler indent and margin handles', () => {
  let dir: string
  let docPath: string

  test.beforeEach(async () => {
    dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-ruler-')))
    docPath = join(dir, 'ruler.docx')
    writeFileSync(docPath, await minimalDocx())
  })

  test.afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  test('markers write paragraph indents, margin handle writes page margins', async () => {
    test.setTimeout(150_000)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-ruler-indent',
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
      await page.locator('.ribbon-tab', { hasText: /^View$/ }).click()
      await page.getByRole('button', { name: 'Ruler', exact: true }).click()
      const ruler = page.locator('.ruler')
      await expect(ruler).toBeVisible()

      const first = page.locator('.doc-page p', { hasText: 'First paragraph' }).first()
      await first.click({ position: { x: 20, y: 8 } })
      const rulerBox = (await ruler.boundingBox())!
      const pxPerInch = (rulerBox.width / PAGE_WIDTH_TWIPS) * 1440

      const marker = (label: string) => page.getByRole('slider', { name: label, exact: true })
      await expect(marker('First Line Indent')).toHaveAttribute('aria-valuenow', '1440')
      await expect(marker('Right Indent')).toHaveAttribute('aria-valuenow', '10800')

      // first-line marker: only the first line moves, one undo step
      await dragBy(page, marker('First Line Indent'), pxPerInch)
      await expect(first).toHaveCSS('text-indent', '96px')
      await expect(first).toHaveCSS('margin-inline-start', '0px')
      const mod = process.platform === 'darwin' ? 'Meta' : 'Control'
      await page.keyboard.press(`${mod}+z`)
      await expect(first).toHaveCSS('text-indent', '0px')
      await page.keyboard.press(`${mod}+Shift+z`)
      await expect(first).toHaveCSS('text-indent', '96px')

      // hanging marker: wrapped lines move, the first line stays put
      await dragBy(page, marker('Hanging Indent'), pxPerInch / 2)
      await expect(first).toHaveCSS('margin-inline-start', '48px')
      await expect(first).toHaveCSS('text-indent', '48px')
      await expect(marker('First Line Indent')).toHaveAttribute('aria-valuenow', '2880')

      // left box: both move together
      await dragBy(page, marker('Left Indent'), pxPerInch / 2)
      await expect(first).toHaveCSS('margin-inline-start', '96px')
      await expect(first).toHaveCSS('text-indent', '48px')

      // right marker snaps to 1/16 in; Option drags unsnapped
      await dragBy(page, marker('Right Indent'), -pxPerInch)
      await expect(first).toHaveCSS('margin-inline-end', '96px')
      await dragBy(page, marker('Right Indent'), -pxPerInch / 24, true)
      await expect(first).toHaveCSS('margin-inline-end', '100px')

      // the second paragraph is untouched
      const second = page.locator('.doc-page p', { hasText: 'Second paragraph' }).first()
      await expect(second).toHaveCSS('margin-inline-start', '0px')

      // margin boundary: the section's left margin grows, the markers ride along
      const leftMargin = page.locator('.ruler-margin-handle[data-side="left"]')
      await expect(leftMargin).toHaveCSS('cursor', 'col-resize')
      await dragBy(page, leftMargin, pxPerInch / 2)
      await expect(marker('Left Margin')).toHaveAttribute('aria-valuenow', '2160')
      await expect(ruler.locator('.ruler-zone').first()).toHaveCSS('width', '144px')
      await expect(page.locator('.doc-page').first()).toHaveCSS('padding-left', '144px')
      await expect(marker('Left Indent')).toHaveAttribute('aria-valuenow', '3600')

      // in a table the markers measure from the cell and the margin handles hide
      await page.locator('.doc-page td', { hasText: 'Cell B' }).first().click()
      await expect(page.locator('.ruler-margin-handle')).toHaveCount(0)
      const cellLeft = Number(await marker('Left Indent').getAttribute('aria-valuenow'))
      expect(cellLeft).toBeGreaterThan(2160 + 4000)
      await first.click({ position: { x: 20, y: 8 } })
      await expect(page.locator('.ruler-margin-handle')).toHaveCount(2)
    } finally {
      await closeAndSaveVideo(launched, 'docs-ruler-indent')
    }
  })
})
