import { test, expect } from '@playwright/test'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, screenshotPath, waitForPageWithUrl } from './helpers'

/**
 * Rich header editing: double-click the page-1 header strip, type bold text and a
 * PAGE field through the contextual tab, unlink section 2's header from section 1
 * and give it its own text, widen the header distance, save. The docx then carries
 * a bold run + PAGE field in header1.xml, an own headerReference on section 2 and
 * the new w:header margin.
 */

interface AidocsWindow {
  __aidocs?: { editor?: unknown; save?: () => Promise<unknown> }
}

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const SECT_PR =
  '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'

/** two sections, each long enough to span pages, section 1 carrying a plain header */
async function twoSectionDocx(): Promise<Buffer> {
  const para = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
  const filler = (label: string) =>
    Array.from({ length: 70 }, (_, i) => para(`${label} paragraph ${i + 1}`)).join('')
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  zip.file(
    'word/_rels/document.xml.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>',
  )
  zip.file(
    'word/header1.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:hdr ${W}><w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>Plain header</w:t></w:r></w:p></w:hdr>`,
  )
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${W} ${R}><w:body>${filler('First')}` +
      `<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rId5"/>${SECT_PR}</w:sectPr></w:pPr></w:p>` +
      `${filler('Second')}<w:sectPr>${SECT_PR}</w:sectPr></w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

/** one section whose header is a logo + company-name layout table over a ruled blank paragraph */
async function tableHeaderDocx(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  zip.file(
    'word/_rels/document.xml.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>',
  )
  zip.file(
    'word/_rels/header1.xml.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/logo.png"/></Relationships>',
  )
  zip.file('word/media/logo.png', Buffer.from(PNG_1X1, 'base64'))
  const pic =
    '<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><wp:extent cx="1512000" cy="375622"/><wp:docPr id="1" name="Picture 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="logo.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1512000" cy="375622"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
  zip.file(
    'word/header1.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:hdr ${W} ${R}><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2835"/><w:gridCol w:w="7370"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2835" w:type="dxa"/></w:tcPr><w:p>${pic}</w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="7370" w:type="dxa"/></w:tcPr><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Letterhead Company Ltd</w:t></w:r></w:p><w:p><w:r><w:t>LETTERHEAD SCIENCE AND TECHNOLOGY</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="8" w:space="1" w:color="000000"/></w:pBdr></w:pPr></w:p></w:hdr>`,
  )
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${W} ${R}><w:body><w:p><w:r><w:t>Body placeholder</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId5"/>${SECT_PR}</w:sectPr></w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

const PNG_1X1 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='

async function part(path: string, name: string): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(path))
  return (await zip.file(name)?.async('string')) ?? ''
}

test.describe('docs header/footer rich editing', () => {
  let dir: string
  let docPath: string

  test.beforeEach(async () => {
    dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-hf-')))
    docPath = join(dir, 'sections.docx')
    writeFileSync(docPath, await twoSectionDocx())
  })

  test.afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  test('bold header text, page field, unlinked section header and header distance save', async () => {
    test.setTimeout(180_000)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-header-footer-edit',
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
      const header = page.locator('.page-wrap > .page-hf-header')
      await expect(header).toContainText('Plain header')
      // the open-time respell kick focuses the body and shields keystrokes for a moment
      await page.waitForTimeout(1500)
      await header.dblclick()
      const editor = header.locator('.page-hf-editor')
      await expect(editor).toBeVisible()
      // the contextual tab appears and the ribbon acts on the strip editor
      const hfTab = page.getByRole('button', { name: 'Header & Footer', exact: true })
      await expect(hfTab).toBeVisible()
      await page.keyboard.press('End')
      await page.keyboard.type(' ')
      // Home-tab formatting acts on the strip editor; TipTap hands focus back a frame after the click
      await page.getByRole('button', { name: 'Home', exact: true }).click()
      const boldButton = page.locator('.ribbon [data-tip^="Bold"]').first()
      await boldButton.click()
      await expect(editor).toBeFocused()
      await page.keyboard.type('Bold')
      await boldButton.click()
      await expect(editor).toBeFocused()
      await page.keyboard.type(' ')
      await expect(editor.locator('strong')).toHaveText('Bold')
      await hfTab.click()
      await page.getByRole('button', { name: 'Page Number', exact: true }).click()
      await page.getByRole('button', { name: 'Current Position' }).click()
      await expect(editor.locator('.hf-field[data-hf-num="1"]')).toHaveCount(1)
      await page.screenshot({ path: screenshotPath('docs-header-footer-edit') })
      // header distance: 720 twips = 1.27 cm -> 2 cm (explicit unit: the field follows the unit preference)
      const dist = page.getByLabel('Header from Top')
      await dist.fill('2 cm')
      await dist.press('Enter')
      await page.getByRole('button', { name: 'Close Header and Footer' }).click()
      await expect(editor).toHaveCount(0)
      await expect(header).toContainText('Plain header Bold 1')

      // section 2 inherits section 1's header: unlink it on its first page and retype
      const gapHeader = page.locator('.page-gap-hf.page-hf-header[data-hf-section="1"]').first()
      await gapHeader.scrollIntoViewIfNeeded()
      await expect(gapHeader).toContainText('Plain header Bold')
      await gapHeader.dblclick()
      const gapEditor = gapHeader.locator('.page-hf-editor')
      await expect(gapEditor).toBeVisible()
      const link = page.getByRole('button', { name: 'Link to Previous' })
      await expect(link).toHaveClass(/active/)
      await link.click()
      await expect(link).not.toHaveClass(/active/)
      // unlinking re-paginates; the editor follows the strip into its new widget and takes focus back
      await expect(gapHeader.locator('.page-hf-editor')).toBeFocused()
      await page.keyboard.press('ControlOrMeta+a')
      await page.keyboard.type('Second section header')
      await page.getByRole('button', { name: 'Close Header and Footer' }).click()
      await expect(page.locator('.page-wrap > .page-hf-header')).toContainText(
        'Plain header Bold 1',
      )

      expect(await page.evaluate(() => (window as unknown as AidocsWindow).__aidocs!.save!())).toBe(
        true,
      )
      await expect
        .poll(() => part(docPath, 'word/header1.xml'), { timeout: 15_000 })
        .toContain('<w:b/>')
      const hdr1 = await part(docPath, 'word/header1.xml')
      expect(hdr1).toMatch(/<w:rPr><w:b\/>(<w:bCs\/>)?<\/w:rPr><w:t[^>]*>Bold<\/w:t>/)
      expect(hdr1).toContain('<w:instrText xml:space="preserve"> PAGE </w:instrText>')
      const doc = await part(docPath, 'word/document.xml')
      // section 1 (break paragraph) keeps rId5, the trailing sectPr got its own reference and the wider header margin
      const sectPrs = doc.match(/<w:sectPr>[\s\S]*?<\/w:sectPr>/g) ?? []
      expect(sectPrs).toHaveLength(2)
      expect(sectPrs[0]).toContain('r:id="rId5"')
      expect(sectPrs[0]).toContain('w:header="1134"')
      expect(sectPrs[1]).toMatch(/<w:headerReference w:type="default" r:id="rId\d+"\/>/)
      const rId = /<w:headerReference w:type="default" r:id="(rId\d+)"\/>/.exec(sectPrs[1])![1]
      const rels = await part(docPath, 'word/_rels/document.xml.rels')
      const target = new RegExp(`Id="${rId}"[^>]*Target="([^"]+)"`).exec(rels)?.[1]
      expect(target).toBeTruthy()
      expect(await part(docPath, `word/${target}`)).toContain('Second section header')

      // Link to Previous with the editor still open on section 2: the strip being
      // discarded must not land in section 1's part, and the own reference goes
      await expect(gapHeader).toContainText('Second section header')
      await gapHeader.dblclick()
      await expect(gapHeader.locator('.page-hf-editor')).toBeVisible()
      await page.keyboard.press('End')
      await page.keyboard.type(' junk')
      page.once('dialog', (d) => void d.accept())
      await page.getByRole('button', { name: 'Link to Previous' }).click()
      await expect(gapHeader.locator('.page-hf-editor')).toHaveCount(0)
      await expect(gapHeader).toContainText('Plain header Bold')
      await expect(gapHeader).not.toContainText('junk')
      await expect(page.locator('.page-wrap > .page-hf-header')).toContainText(
        'Plain header Bold 1',
      )
      expect(await page.evaluate(() => (window as unknown as AidocsWindow).__aidocs!.save!())).toBe(
        true,
      )
      const trailingSectPr = async () =>
        (await part(docPath, 'word/document.xml')).match(/<w:sectPr>[\s\S]*?<\/w:sectPr>/g)!.at(-1)!
      await expect
        .poll(async () => (await trailingSectPr()).includes('headerReference'), { timeout: 15_000 })
        .toBe(false)
      const hdr1After = await part(docPath, 'word/header1.xml')
      expect(hdr1After).not.toContain('junk')
      expect(hdr1After).not.toContain('Second section header')
      expect(hdr1After).toMatch(/<w:t[^>]*>Bold<\/w:t>/)
    } finally {
      await closeAndSaveVideo(launched, 'docs-header-footer-edit')
    }
  })

  test('a layout-table header edits in place: cell text, a picture from file, saved into the table', async () => {
    test.setTimeout(150_000)
    const tablePath = join(dir, 'letterhead.docx')
    writeFileSync(tablePath, await tableHeaderDocx())
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-header-footer-edit-table',
      openFile: tablePath,
    })
    const { app } = launched
    try {
      const page = await waitForPageWithUrl(app, '://docs/')
      await page.waitForFunction(
        () => Boolean((window as unknown as AidocsWindow).__aidocs?.editor),
        undefined,
        { timeout: 30_000 },
      )
      const header = page.locator('.page-wrap > .page-hf-header')
      await expect(header).toContainText('Letterhead Company Ltd')
      await page.waitForTimeout(1500)
      await header.dblclick()
      const editor = header.locator('.page-hf-editor')
      await expect(editor).toBeVisible()
      // the table row is part of the editor: logo cell and text cell stay in place
      const row = editor.locator('.page-hf-row')
      await expect(row).toBeVisible()
      await expect(row.locator('img.doc-inline-img')).toHaveCount(1)
      const title = row.locator('.page-hf-cell').nth(1).locator('.doc-textbox-para').first()
      await expect(title).toHaveText('Letterhead Company Ltd')
      await title.click()
      await page.keyboard.press('End')
      await page.keyboard.type(' Group')
      await expect(title).toHaveText('Letterhead Company Ltd Group')
      // Backspace at the cell start stays inside the cell (no merge with the logo cell)
      await page.keyboard.press('Home')
      await page.keyboard.press('Backspace')
      await expect(row.locator('.page-hf-cell')).toHaveCount(2)
      await expect(title).toHaveText('Letterhead Company Ltd Group')
      // Home / End move inside the cell without the document scrolling away
      await page.keyboard.press('End')
      await page.waitForTimeout(600)
      expect(await page.evaluate(() => document.querySelector('.editor-scroll')!.scrollTop)).toBe(0)
      // the logo selects with Word-style handles; a corner drag resizes it in proportion
      const logo = row.locator('.page-hf-cell').nth(0).locator('img.doc-inline-img')
      const logoBefore = (await logo.boundingBox())!
      await logo.click()
      const seHandle = row.locator('.doc-inline-pic.ProseMirror-selectednode .img-resize-handle-se')
      await expect(seHandle).toBeVisible()
      await expect(page.getByRole('button', { name: 'Picture Format', exact: true })).toBeVisible()
      const seBox = (await seHandle.boundingBox())!
      await page.mouse.move(seBox.x + seBox.width / 2, seBox.y + seBox.height / 2)
      await page.mouse.down()
      // stays inside the 189px logo cell (max-width would clip a wider picture)
      await page.mouse.move(seBox.x + 25, seBox.y + 6, { steps: 8 })
      await page.mouse.up()
      await expect
        .poll(async () => (await logo.boundingBox())!.width)
        .toBeGreaterThan(logoBefore.width + 15)
      const logoAfter = (await logo.boundingBox())!
      expect(logoAfter.height / logoAfter.width).toBeCloseTo(
        logoBefore.height / logoBefore.width,
        1,
      )
      // the picture stayed in its cell
      await expect(row.locator('.page-hf-cell').nth(0).locator('img.doc-inline-img')).toHaveCount(1)
      // back to the text cell: the contextual tab follows the caret
      await title.click()
      await page.keyboard.press('End')
      await expect(page.getByRole('button', { name: 'Picture Format', exact: true })).toHaveCount(0)
      // Picture from File on the contextual tab (the native picker is stubbed in the main process)
      const sealPath = join(dir, 'seal.png')
      writeFileSync(sealPath, Buffer.from(PNG_1X1, 'base64'))
      await app.evaluate(({ dialog }, file) => {
        dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [file] })
      }, sealPath)
      // the contextual tab is already active while editing (clicking it again collapses the ribbon)
      await expect(page.getByRole('button', { name: 'Header & Footer', exact: true })).toBeVisible()
      await page.getByRole('button', { name: 'Picture', exact: true }).click()
      await expect(row.locator('.page-hf-cell').nth(1).locator('img.doc-inline-img')).toHaveCount(1)
      await page.screenshot({ path: screenshotPath('docs-header-footer-edit-table') })
      await page.getByRole('button', { name: 'Close Header and Footer' }).click()
      await expect(editor).toHaveCount(0)
      await expect(header).toContainText('Letterhead Company Ltd Group')
      await expect(header.locator('.page-hf-row img')).toHaveCount(2)

      expect(await page.evaluate(() => (window as unknown as AidocsWindow).__aidocs!.save!())).toBe(
        true,
      )
      await expect
        .poll(() => part(tablePath, 'word/header1.xml'), { timeout: 15_000 })
        .toContain('Letterhead Company Ltd Group')
      const hdr = await part(tablePath, 'word/header1.xml')
      // the logo cell keeps its original bytes at its dragged size; the edited cell carries the new picture
      expect(hdr).toContain('<a:blip r:embed="rId1"/>')
      expect(hdr).not.toContain('<wp:extent cx="1512000" cy="375622"/>')
      const logoCx = Number(/<wp:extent cx="(\d+)"/.exec(hdr)?.[1])
      expect(logoCx).toBeGreaterThan(1512000 * 1.1)
      expect((hdr.match(/<w:tbl>/g) ?? []).length).toBe(1)
      expect((hdr.match(/<a:blip /g) ?? []).length).toBe(2)
      expect(hdr).not.toContain('aidocs-new-pic-')
      const rels = await part(tablePath, 'word/_rels/header1.xml.rels')
      expect((rels.match(/relationships\/image"/g) ?? []).length).toBe(2)
    } finally {
      await closeAndSaveVideo(launched, 'docs-header-footer-edit-table')
    }
  })
})
