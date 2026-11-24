/**
 * Fields, links and Mailings as Word has them, the parts left after the first
 * rounds: linked pictures in a table cell and a text box are listed under
 * Edit > Links… and updated, relinked and broken there; Lock, Unlock and
 * Unlink Fields act on a table of contents and on the page field of a footer;
 * Update Field writes the standard's counting formats; Preview Results asks
 * the Fill-in questions; the AutoText a document carries is offered while
 * typing and under Insert > AutoText; Define styles acts while Track Changes
 * records.
 */
import { test, expect, type ElectronApplication, type Locator, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { pathToFileURL } from 'node:url'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu, withWordPart } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-fields-links-leftovers'
const IS_MAC = process.platform === 'darwin'
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

const PNG_WIDE =
  'iVBORw0KGgoAAAANSUhEUgAAACgAAAAUCAIAAABwJOjsAAAAJ0lEQVR42mM8ISfHMBCAiWGAwKjFoxaPWjxq8ajFoxaPWjxqMf0BAF+MASwgP5VpAAAAAElFTkSuQmCC'
const PNG_SQUARE =
  'iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAIAAAAC64paAAAAHUlEQVR42mOUkzvBQC5gYqAAjGoe1TyqeVQz5ZoBOL0BLMMULg8AAAAASUVORK5CYII='

const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'

const STYLES =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${W}>` +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/></w:style>' +
  '</w:styles>'

const run = (text: string, rPr = '') =>
  `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`
const fld = (type: string, attrs = '') => `<w:r><w:fldChar w:fldCharType="${type}"${attrs}/></w:r>`
const instr = (code: string) =>
  `<w:r><w:instrText xml:space="preserve"> ${code} </w:instrText></w:r>`
const field = (code: string, result: string, attrs = '') =>
  fld('begin', attrs) + instr(code) + fld('separate') + result + fld('end')
const heading = (text: string) =>
  `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>${run(text)}</w:p>`

/** clicks a menu bar item once the renderer has told the menu it is available */
async function clickMenu(app: ElectronApplication, path: string[]) {
  const click = () =>
    app.evaluate(({ Menu }, path) => {
      let items = Menu.getApplicationMenu()?.items
      let item: Electron.MenuItem | undefined
      for (const label of path) {
        item = items?.find((candidate) => candidate.label === label)
        items = item?.submenu?.items
      }
      if (!item?.enabled) return false
      item.click()
      return true
    }, path)
  await expect.poll(click, { timeout: 15_000 }).toBe(true)
}

const pickFile = (app: ElectronApplication, file: string) =>
  app.evaluate(({ dialog }, picked) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [picked] })
  }, file)

const part = async (file: string, name: string): Promise<string> =>
  (await (await JSZip.loadAsync(readFileSync(file))).file(name)?.async('string')) ?? ''
const documentXml = (file: string) => part(file, 'word/document.xml')

/** the outside files the pictures of a saved document are linked to, in document order */
async function linkedTargets(file: string): Promise<string[]> {
  const xml = await documentXml(file)
  const rels = await part(file, 'word/_rels/document.xml.rels')
  return [...xml.matchAll(/<a:blip\b[^>]*\br:link="([^"]+)"/g)].map((m) => {
    const rel = new RegExp(`<Relationship\\b[^>]*Id="${m[1]}"[^>]*>`).exec(rels)?.[0] ?? ''
    return rel.includes('TargetMode="External"') ? (/Target="([^"]*)"/.exec(rel)?.[1] ?? '') : ''
  })
}

const mediaOf = async (file: string): Promise<string[]> => {
  const zip = await JSZip.loadAsync(readFileSync(file))
  const names = Object.keys(zip.files).filter(
    (name) => name.startsWith('word/media/') && !zip.files[name]!.dir,
  )
  return Promise.all(names.map((name) => zip.file(name)!.async('base64')))
}

async function open(docPath: string) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await page.waitForFunction(
    () => Boolean((window as unknown as { __aidocs?: { editor?: unknown } }).__aidocs?.editor),
    undefined,
    { timeout: 30_000 },
  )
  return { launched, app: launched.app, page }
}

async function openLinks(app: ElectronApplication, page: Page) {
  await clickMenu(app, ['Edit', 'Links…'])
  const dialog = page.locator('.links-dialog')
  await expect(dialog).toBeVisible({ timeout: 15_000 })
  return dialog
}

/** a picture in a line of text that is embedded and names an outside file, as Word writes one */
const linkedPicture = (id: number, linkId: string) =>
  `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="762000" cy="381000"/>` +
  `<wp:docPr id="${id}" name="Picture ${id}"/>` +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
  `<pic:blipFill><a:blip r:embed="rIdImg" r:link="${linkId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="762000" cy="381000"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'

const cell = (inner: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`
const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/><w:tblBorders>' +
  ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map((side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`)
    .join('') +
  '</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
  `<w:tr>${cell(`<w:p>${run('Logo ')}${linkedPicture(11, 'rIdLogo')}</w:p>`)}${cell(`<w:p>${run('Free cell')}</w:p>`)}</w:tr></w:tbl>`
const TEXTBOX =
  '<w:p><w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
  '<mc:Choice Requires="wps"><w:drawing>' +
  '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>300000</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="3000000" cy="1200000"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapTopAndBottom/>' +
  '<wp:docPr id="21" name="Text Box 21"/><wp:cNvGraphicFramePr/>' +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="1200000"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>' +
  `<wps:txbx><w:txbxContent><w:p>${run('Boxed ')}${linkedPicture(12, 'rIdBoxed')}</w:p></w:txbxContent></wps:txbx>` +
  '<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r></w:p>'

/** adds relationships to the main document part */
async function withRels(docx: Buffer, entries: string): Promise<Buffer> {
  const zip = await JSZip.loadAsync(docx)
  const path = 'word/_rels/document.xml.rels'
  zip.file(
    path,
    (await zip.file(path)!.async('string')).replace(
      '</Relationships>',
      `${entries}</Relationships>`,
    ),
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

const externalImage = (id: string, file: string) =>
  `<Relationship Id="${id}" Type="${REL}/image" Target="${pathToFileURL(file).href}" TargetMode="External"/>`

test('linked pictures in a table cell and a text box are listed, updated, relinked by an insert and broken', async () => {
  test.setTimeout(360_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-links-outside-')))
  const docPath = join(dir, 'host.docx')
  const logoPath = join(dir, 'logo.png')
  const boxedPath = join(dir, 'boxed.png')
  const photoPath = join(dir, 'photo.png')
  // the files hold other pictures than the ones embedded: nothing reads them until asked
  writeFileSync(logoPath, Buffer.from(PNG_WIDE, 'base64'))
  writeFileSync(boxedPath, Buffer.from(PNG_SQUARE, 'base64'))
  writeFileSync(photoPath, Buffer.from(PNG_SQUARE, 'base64'))
  writeFileSync(
    docPath,
    await withRels(
      await docxWith(para('Host first') + TABLE + TEXTBOX + para('Host last') + SECT_PR),
      externalImage('rIdLogo', logoPath) + externalImage('rIdBoxed', boxedPath),
    ),
  )
  let { launched, app, page } = await open(docPath)
  try {
    await expect(page.locator('.doc-page p', { hasText: 'Host first' })).toBeVisible({
      timeout: 30_000,
    })
    const cellPicture = page.locator('.doc-page td img.doc-inline-img')
    await expect(cellPicture).toHaveCount(1)
    // drawn from the bytes the document holds, not from the file its link names
    await expect(cellPicture).toHaveAttribute('src', /^data:image\/png;base64,/)
    await expect(cellPicture).not.toHaveAttribute('src', new RegExp(PNG_WIDE.slice(0, 40)))

    let dialog = await openLinks(app, page)
    let rows = dialog.locator('button.links-row')
    await expect(rows).toHaveCount(2)
    await expect(rows.nth(0)).toContainText('logo.png')
    await expect(rows.nth(0)).toContainText('Picture')
    await expect(rows.nth(1)).toContainText('boxed.png')
    await page.screenshot({ path: test.info().outputPath('links-outside-body.png') })
    // both at once: Update Now reads the two files
    await rows.nth(0).click()
    await rows.nth(1).click({ modifiers: ['Shift'] })
    await dialog.locator('button[name="update"]').click()
    await expect(rows.nth(0)).toContainText('Updated')
    await expect(rows.nth(1)).toContainText('Updated')
    await expect(cellPicture).toHaveAttribute('src', new RegExp(PNG_WIDE.slice(0, 40)))
    await dialog.locator('button[name="close"]').click()
    await saveFromMenu(app)
    await expect
      .poll(async () => (await mediaOf(docPath)).includes(PNG_WIDE), { timeout: 30_000 })
      .toBe(true)
    expect(await mediaOf(docPath)).toContain(PNG_SQUARE)
    expect(await linkedTargets(docPath)).toEqual([
      pathToFileURL(logoPath).href,
      pathToFileURL(boxedPath).href,
    ])
    expect(await documentXml(docPath)).not.toContain('aidocs-')

    // Break Link on the picture of the cell: the one in the text box keeps its link
    dialog = await openLinks(app, page)
    rows = dialog.locator('button.links-row')
    await rows.nth(0).click()
    page.once('dialog', (confirm) => void confirm.accept())
    await dialog.locator('button[name="break"]').click()
    await expect(rows).toHaveCount(1)
    await expect(rows.nth(0)).toContainText('boxed.png')
    await dialog.locator('button[name="close"]').click()

    // a picture inserted with Link to File into the other cell keeps its link
    await page.locator('.doc-page td', { hasText: 'Free cell' }).click()
    await page.keyboard.press('End')
    await pickFile(app, photoPath)
    await clickMenu(app, ['Insert', 'Pictures', 'Picture from File…'])
    const options = page.locator('.insert-picture-options-dialog')
    await expect(options).toContainText('photo.png')
    await options.locator('input[name="link"]').check()
    await options.locator('button[type="submit"]').click()
    await expect(options).toBeHidden()
    await expect(page.locator('.doc-page td img.doc-inline-img')).toHaveCount(2)
    dialog = await openLinks(app, page)
    rows = dialog.locator('button.links-row')
    await expect(rows).toHaveCount(2)
    await expect(rows.nth(0)).toContainText('photo.png')
    await expect(rows.nth(1)).toContainText('boxed.png')
    await dialog.locator('button[name="close"]').click()
    await saveFromMenu(app)
    await expect
      .poll(() => linkedTargets(docPath), { timeout: 30_000 })
      .toEqual([pathToFileURL(photoPath).href, pathToFileURL(boxedPath).href])
    const xml = await documentXml(docPath)
    expect(xml).not.toContain('aidocs-')
    expect(xml.match(/<a:blip\b/g)).toHaveLength(3)

    // read again from the file: the two links are still listed
    await closeAndSaveVideo(launched, VIDEO_DIR)
    ;({ launched, app, page } = await open(docPath))
    await expect(page.locator('.doc-page td img.doc-inline-img')).toHaveCount(2, {
      timeout: 30_000,
    })
    dialog = await openLinks(app, page)
    await expect(dialog.locator('button.links-row')).toHaveCount(2)
    await expect(dialog.locator('button.links-row').nth(0)).toContainText('photo.png')
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR).catch(() => {})
    rmSync(dir, { recursive: true, force: true })
  }
})

const pageRef = (anchor: string, page: string) => field(`PAGEREF ${anchor} \\h`, run(page))
const tocLine = (text: string, anchor: string, page: string, lead = '', tail = '') =>
  '<w:p><w:pPr><w:pStyle w:val="TOC1"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9000"/></w:tabs></w:pPr>' +
  lead +
  `<w:hyperlink w:anchor="${anchor}" w:history="1">${run(text)}<w:r><w:tab/></w:r>${pageRef(anchor, page)}</w:hyperlink>` +
  tail +
  '</w:p>'
const TOC =
  tocLine(
    'Stale One',
    '_Toc1',
    '4',
    fld('begin') + instr('TOC \\o "1-3" \\h \\z \\u') + fld('separate'),
  ) + tocLine('Stale Two', '_Toc2', '5', '', fld('end'))

/** the document with a footer part that reads "Page <PAGE>" */
async function withFooter(docx: Buffer): Promise<Buffer> {
  const zip = await JSZip.loadAsync(docx)
  zip.file(
    'word/footer1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ${W}><w:p>${run('Page ')}${field('PAGE', run('1'))}</w:p></w:ftr>`,
  )
  const rels = 'word/_rels/document.xml.rels'
  zip.file(
    rels,
    (await zip.file(rels)!.async('string')).replace(
      '</Relationships>',
      `<Relationship Id="rIdFooter" Type="${REL}/footer" Target="footer1.xml"/></Relationships>`,
    ),
  )
  const types = '[Content_Types].xml'
  zip.file(
    types,
    (await zip.file(types)!.async('string')).replace(
      '</Types>',
      '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>',
    ),
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

type FieldEditor = {
  state: { selection: { from: number; to: number; node?: { type: { name: string } } } }
}
const selection = (page: Page) =>
  page.evaluate(() => {
    const { selection } = (window as never as { __aidocs: { editor: FieldEditor } }).__aidocs.editor
      .state
    return { from: selection.from, node: selection.node?.type.name ?? null }
  })

test('Lock, Unlock and Unlink Fields act on a table of contents and on the page field of a footer', async () => {
  test.skip(!IS_MAC, 'the field chords are Word for Mac’s; other platforms keep their own')
  test.setTimeout(360_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-field-tables-')))
  const docPath = join(dir, 'report.docx')
  const body =
    TOC +
    heading('Alpha') +
    para('Body of alpha') +
    heading('Beta') +
    para('Body of beta') +
    SECT_PR.replace('<w:pgSz', '<w:footerReference w:type="default" r:id="rIdFooter"/><w:pgSz')
  writeFileSync(
    docPath,
    await withFooter(await withWordPart(await docxWith(body), 'styles', STYLES)),
  )
  let { launched, app, page } = await open(docPath)
  try {
    const lines = page.locator('.doc-page .doc-toc-line')
    await expect(lines).toHaveCount(2, { timeout: 30_000 })
    await expect(lines.nth(0)).toContainText('Stale One')
    // the open-time respell kick focuses the body and shields keystrokes for a moment
    await page.waitForTimeout(1500)

    // Previous Field from the end of the document stops at the table
    await page.locator('.doc-page p', { hasText: 'Body of beta' }).click()
    await page.keyboard.press('Shift+F11')
    await expect.poll(() => selection(page)).toEqual({ from: 0, node: 'docProtected' })

    // locked: Update Table and Update Field leave it as it is
    await page.keyboard.press('Meta+Digit3')
    await page.locator('.ribbon-tab', { hasText: /^References$/ }).click()
    await page
      .locator('[data-ribbon-body]')
      .getByRole('button', { name: 'Update Table', exact: true })
      .first()
      .click()
    await expect(page.locator('.app-toast.show')).toContainText('Locked, not updated')
    await expect(lines.nth(0)).toContainText('Stale One')
    await page.keyboard.press('Shift+F11')
    await page.keyboard.press('F11')
    await page.locator('.doc-page p', { hasText: 'Body of beta' }).click()
    await page.keyboard.press('Shift+F11')
    await page.keyboard.press('F9')
    await expect(page.locator('.status-bar')).toContainText('Locked, not updated')
    await expect(lines.nth(1)).toContainText('Stale Two')
    await saveFromMenu(app)
    await expect.poll(() => documentXml(docPath), { timeout: 30_000 }).toContain('w:fldLock="true"')
    expect(await documentXml(docPath)).toMatch(
      /<w:fldChar w:fldCharType="begin" w:fldLock="true"\/><\/w:r><w:r><w:instrText[^>]*> TOC /,
    )

    // read again from the file it is still locked; unlocked, Update Field rebuilds it from the headings
    await closeAndSaveVideo(launched, VIDEO_DIR)
    ;({ launched, app, page } = await open(docPath))
    const again = page.locator('.doc-page .doc-toc-line')
    await expect(again).toHaveCount(2, { timeout: 30_000 })
    await page.waitForTimeout(1500)
    await page.locator('.doc-page p', { hasText: 'Body of beta' }).click()
    await page.keyboard.press('Shift+F11')
    await page.keyboard.press('F9')
    await expect(page.locator('.status-bar')).toContainText('Locked, not updated')
    await expect(again.nth(0)).toContainText('Stale One')
    await page.keyboard.press('Meta+Digit4')
    await page.keyboard.press('F9')
    await expect(again.nth(0)).toContainText('Alpha')
    await expect(again.nth(1)).toContainText('Beta')
    await page.screenshot({ path: test.info().outputPath('toc-updated.png') })
    await saveFromMenu(app)
    await expect.poll(() => documentXml(docPath), { timeout: 30_000 }).not.toContain('w:fldLock')
    expect(await documentXml(docPath)).not.toContain('Stale One')

    // Unlink Fields: the lines stay, the field is gone
    await page.locator('.doc-page p', { hasText: 'Body of beta' }).click()
    await page.keyboard.press('Shift+F11')
    await expect.poll(() => selection(page)).toEqual({ from: 0, node: 'docProtected' })
    await page.keyboard.press('Meta+Digit6')
    await page.locator('.doc-page p', { hasText: 'Body of beta' }).click()
    // nothing is left for Previous Field to stop at
    await page.keyboard.press('Shift+F11')
    await expect.poll(async () => (await selection(page)).node).toBeNull()
    await saveFromMenu(app)
    await expect.poll(() => documentXml(docPath), { timeout: 30_000 }).not.toContain('TOC \\o')
    const unlinked = await documentXml(docPath)
    expect(unlinked).not.toContain('PAGEREF')
    expect(unlinked).not.toContain('w:hyperlink')
    expect(unlinked.match(/<w:pStyle w:val="TOC1"\/>/g)).toHaveLength(2)

    // the footer: its page number is live until it is locked
    const footer = page.locator('.page-wrap > .page-hf-footer').first()
    await footer.scrollIntoViewIfNeeded()
    await footer.dblclick()
    const editor = footer.locator('.page-hf-editor')
    await expect(editor).toBeVisible()
    const live = editor.locator('.hf-field')
    await expect(live).toHaveCount(1)
    await live.click()
    await page.keyboard.press('Meta+Digit3')
    await expect(live).toHaveCount(0)
    await expect(editor).toHaveText('Page 1')
    await page.getByRole('button', { name: 'Close Header and Footer' }).click()
    await expect(editor).toHaveCount(0)
    await saveFromMenu(app)
    await expect
      .poll(() => part(docPath, 'word/footer1.xml'), { timeout: 30_000 })
      .toContain('w:fldLock="true"')
    expect(await part(docPath, 'word/footer1.xml')).toMatch(
      /w:fldLock="true"\/><\/w:r><w:r><w:instrText[^>]*>\s*PAGE\s*<\/w:instrText>.*?<w:t[^>]*>1<\/w:t>/,
    )

    // read again: the locked number is text of the footer, and Unlock Fields makes it live
    await closeAndSaveVideo(launched, VIDEO_DIR)
    ;({ launched, app, page } = await open(docPath))
    const reopened = page.locator('.page-wrap > .page-hf-footer').first()
    await expect(reopened).toContainText('Page 1', { timeout: 30_000 })
    await page.waitForTimeout(1500)
    await reopened.scrollIntoViewIfNeeded()
    await reopened.dblclick()
    const strip = reopened.locator('.page-hf-editor')
    await expect(strip).toBeVisible()
    await expect(strip.locator('.hf-field')).toHaveCount(0)
    await page.keyboard.press('End')
    await page.keyboard.press('Meta+Digit4')
    await expect(strip.locator('.hf-field')).toHaveCount(1)
    await page.getByRole('button', { name: 'Close Header and Footer' }).click()
    await saveFromMenu(app)
    await expect
      .poll(() => part(docPath, 'word/footer1.xml'), { timeout: 30_000 })
      .not.toContain('w:fldLock')
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR).catch(() => {})
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Update Field writes the counting formats of the standard, where a heading is, and the page of a bookmark in a table', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-field-formats-')))
  const docPath = join(dir, 'fields.docx')
  const waiting = (code: string) => field(code, run('?'), ' w:dirty="true"')
  const body =
    heading('Alpha') +
    `<w:p>${run('Count ')}${waiting('NUMPAGES \\* CardText')}${run(' page ')}${waiting('PAGE \\* CHINESENUM3')}` +
    `${run(' dash ')}${waiting('PAGE \\* ArabicDash')}` +
    `${run(' cell ')}${waiting('PAGEREF InCell')}${run(' legal ')}${waiting('PAGE \\* KANJINUM3')}</w:p>` +
    `<w:p>${run('The heading is ')}${waiting('STYLEREF 1 \\p')}</w:p>` +
    '<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>' +
    `<w:tr><w:tc><w:tcPr><w:tcW w:w="8000" w:type="dxa"/></w:tcPr><w:p><w:bookmarkStart w:id="5" w:name="InCell"/>${run('In a cell')}<w:bookmarkEnd w:id="5"/></w:p></w:tc></w:tr></w:tbl>` +
    para('Last line') +
    SECT_PR
  writeFileSync(docPath, await withWordPart(await docxWith(body), 'styles', STYLES))
  const { launched, app, page } = await open(docPath)
  try {
    const line = page.locator('.doc-page p', { hasText: 'Count ' })
    await expect(line).toBeVisible({ timeout: 30_000 })
    await page.waitForTimeout(1500)
    await line.click()
    await page.keyboard.press('F9')
    // one page: "one", the ideograph for one, "- 1 -", page 1; the last format has no rule here
    await expect(line).toHaveText(
      `Count one page ${String.fromCodePoint(0x4e00)} dash - 1 - cell 1 legal ?`,
    )
    await expect(page.locator('.doc-page')).toContainText('The heading is above')
    await saveFromMenu(app)
    await expect.poll(() => documentXml(docPath), { timeout: 30_000 }).toContain('- 1 -')
    const xml = await documentXml(docPath)
    const dirty = [
      ...xml.matchAll(/w:dirty="(?:true|1)"\/><\/w:r><w:r><w:instrText[^>]*>([^<]*)</g),
    ]
    // only the format this editor does not compute still asks Word for an update
    expect(dirty.map((m) => m[1].trim())).toEqual(['PAGE \\* KANJINUM3'])
    expect(xml).toMatch(new RegExp(`<w:t[^>]*>${String.fromCodePoint(0x4e00)}</w:t>`))
    expect(xml).toMatch(/STYLEREF 1 \\p <\/w:instrText>.*?<w:t[^>]*>above<\/w:t>/)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR).catch(() => {})
    rmSync(dir, { recursive: true, force: true })
  }
})

const tab = (page: Page, name: string) =>
  page.locator('.ribbon-tab', { hasText: new RegExp(`^${name}$`) }).click()
const big = (page: Page, name: string | RegExp) =>
  page.locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: name })

test('Preview Results asks the Fill-in question of each record once and shows the answer', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-preview-asks-')))
  const docPath = join(dir, 'letter.docx')
  const listPath = join(dir, 'people.csv')
  writeFileSync(docPath, await docxWith(para('Dear') + para('Regards') + SECT_PR))
  writeFileSync(listPath, 'Name,City\nAda,London\nAlan,Wilmslow\n')
  const { launched, app, page } = await open(docPath)
  try {
    const greeting = page.locator('.doc-page p', { hasText: 'Dear' }).first()
    await expect(greeting).toBeVisible({ timeout: 30_000 })
    await pickFile(app, listPath)
    await tab(page, 'Mailings')
    await big(page, 'Select Recipients').click()
    const record = page.locator('[data-ribbon-body] [data-testid="mailings-record-status"]')
    await expect(record).toHaveText(/Record 1 of 2/)
    await greeting.click()
    await page.keyboard.press('End')
    await page.keyboard.type(' ')
    await big(page, 'Insert Merge Field').click()
    await page.locator('.rb-field-item', { hasText: /^Name$/ }).click()
    await page.keyboard.type(': ')
    await big(page, 'Rules').click()
    await page.locator('.layout-menu button', { hasText: /^Fill-in…$/ }).click()
    const rule = page.locator('.rule-dialog')
    await rule.locator('input[name="prompt"]').fill('Which gift?')
    await rule.locator('input[name="value"]').fill('a card')
    await rule.locator('button.primary').click()
    await expect(rule).toBeHidden()

    const preview = page.locator('.doc-page .doc-merge-preview')
    const question = page.locator('.modal', { hasText: 'Which gift?' })
    await big(page, 'Preview Results').click()
    // the question comes with its default in the box, as in Finish & Merge
    await expect(question).toBeVisible()
    await expect(question).toContainText('Record 1 of 2')
    await expect(question.locator('input')).toHaveValue('a card')
    await page.screenshot({ path: test.info().outputPath('preview-asks.png') })
    await question.locator('input').fill('a pen')
    await question.locator('input').press('Enter')
    await expect(question).toHaveCount(0)
    await expect(preview).toHaveText(['Ada', 'a pen'])

    // the next record is asked for itself; closing the question takes the default
    const next = page.locator('[data-ribbon-body] button[aria-label="Next"]')
    await next.click()
    await expect(question).toBeVisible()
    await expect(question).toContainText('Record 2 of 2')
    await page.keyboard.press('Escape')
    await expect(question).toHaveCount(0)
    await expect(preview).toHaveText(['Alan', 'a card'])
    // back on the first record nothing is asked again
    await page.locator('[data-ribbon-body] button[aria-label="Previous"]').click()
    await expect(record).toHaveText(/Record 1 of 2/)
    await expect(preview).toHaveText(['Ada', 'a pen'])
    await expect(question).toHaveCount(0)

    // a new preview starts over
    await big(page, 'Preview Results').click()
    await expect(preview).toHaveCount(0)
    await big(page, 'Preview Results').click()
    await expect(question).toBeVisible()
    await question.locator('input').press('Enter')
    await expect(preview).toHaveText(['Ada', 'a card'])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR).catch(() => {})
    rmSync(dir, { recursive: true, force: true })
  }
})

const docPart = (name: string, gallery: string, body: string) =>
  `<w:docPart><w:docPartPr><w:name w:val="${name}"/><w:category><w:name w:val="General"/>` +
  `<w:gallery w:val="${gallery}"/></w:category><w:behaviors><w:behavior w:val="content"/></w:behaviors>` +
  `</w:docPartPr><w:docPartBody>${body}</w:docPartBody></w:docPart>`

/** a document that carries AutoText in its glossary part, as one made from a template does */
async function docxWithAutoText(): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await docxWith(para('Letter body') + para('') + SECT_PR))
  zip.file(
    'word/glossary/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:glossaryDocument ${W}><w:docParts>` +
      docPart(
        'Sign off',
        'autoTxt',
        `<w:p>${run('Kind regards, ')}${run('Finance team', '<w:b/>')}</w:p>`,
      ) +
      docPart(
        'Office address',
        'autoTxt',
        `<w:p>${run('1 Main Street')}</w:p><w:p>${run('Springfield')}</w:p>`,
      ) +
      '</w:docParts></w:glossaryDocument>',
  )
  const rels = 'word/_rels/document.xml.rels'
  zip.file(
    rels,
    (await zip.file(rels)!.async('string')).replace(
      '</Relationships>',
      `<Relationship Id="rIdGlossary" Type="${REL}/glossaryDocument" Target="glossary/document.xml"/></Relationships>`,
    ),
  )
  const types = '[Content_Types].xml'
  zip.file(
    types,
    (await zip.file(types)!.async('string')).replace(
      '</Types>',
      '<Override PartName="/word/glossary/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"/></Types>',
    ),
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

type CaretEditor = { commands: { focus(at: string): boolean } }
const caretAtDocEnd = (page: Page) =>
  page.evaluate(() => {
    const editor = (window as never as { __aidocs: { editor: CaretEditor } }).__aidocs.editor
    editor.commands.focus('end')
    return new Promise<void>((resolve) => requestAnimationFrame(() => resolve()))
  })

test('the AutoText a document carries is offered while typing and under Insert > AutoText', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-doc-autotext-')))
  const docPath = join(dir, 'letter.docx')
  writeFileSync(docPath, await docxWithAutoText())
  const { launched, app, page } = await open(docPath)
  try {
    await expect(page.locator('.doc-page p', { hasText: 'Letter body' })).toBeVisible({
      timeout: 30_000,
    })
    await page.waitForTimeout(1500)
    await caretAtDocEnd(page)
    // the start of the entry's name brings up the tip; Enter takes the entry with its formatting
    await page.keyboard.type('Sign')
    const tip = page.locator('.doc-autocomplete-tip')
    await expect(tip).toContainText('Kind regards, Finance team')
    await page.screenshot({ path: test.info().outputPath('autotext-tip.png') })
    await page.keyboard.press('Enter')
    const signed = page.locator('.doc-page p', { hasText: 'Kind regards, Finance team' })
    await expect(signed).toHaveCount(1)
    await expect(signed.locator('strong')).toHaveText('Finance team')
    await expect(page.locator('.doc-page')).not.toContainText('Sign')

    // Insert > AutoText lists the document's entries; they are not this computer's to delete
    await caretAtDocEnd(page)
    await page.keyboard.press('Enter')
    await clickMenu(app, ['Insert', 'AutoText', 'AutoText…'])
    const dialog = page.locator('.auto-text-dialog')
    await expect(dialog).toBeVisible()
    await expect(dialog.locator('select.auto-text-list option')).toHaveText([
      'Sign off',
      'Office address',
    ])
    await dialog.locator('select.auto-text-list').selectOption({ label: 'Office address' })
    await expect(dialog.locator('.auto-text-preview')).toContainText('1 Main Street')
    await expect(dialog.getByRole('button', { name: 'Delete', exact: true })).toBeDisabled()
    await page.screenshot({ path: test.info().outputPath('insert-autotext.png') })
    await dialog.locator('button[type="submit"]').click()
    await expect(dialog).toBeHidden()
    await expect(page.locator('.doc-page p', { hasText: '1 Main Street' })).toHaveCount(1)
    await expect(page.locator('.doc-page p', { hasText: 'Springfield' })).toHaveCount(1)
    await saveFromMenu(app)
    await expect.poll(() => documentXml(docPath), { timeout: 30_000 }).toContain('Springfield')
    expect(await documentXml(docPath)).toContain('Finance team')
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR).catch(() => {})
    rmSync(dir, { recursive: true, force: true })
  }
})

/** the native inputs sit under their drawn boxes: the label takes the click */
const labelOf = (dialog: Locator, text: string) =>
  dialog.locator('label.protect-check', { hasText: text })

test('Define styles based on your formatting acts while Track Changes records, as a tracked formatting change', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-define-tracked-')))
  const docPath = join(dir, 'letter.docx')
  writeFileSync(docPath, await docxWith(para('Opening line') + para('') + SECT_PR))
  const { launched, app, page } = await open(docPath)
  try {
    await expect(page.locator('.doc-page p', { hasText: 'Opening line' })).toBeVisible({
      timeout: 30_000,
    })
    await clickMenu(app, ['Tools', 'AutoCorrect Options…'])
    const options = page.getByRole('dialog', { name: 'AutoCorrect Options' })
    await expect(options).toBeVisible()
    await options.getByRole('tab', { name: 'AutoFormat As You Type', exact: true }).click()
    await labelOf(options, 'Define styles based on your formatting').click()
    await expect(options.getByLabel('Define styles based on your formatting')).toBeChecked()
    await options.getByRole('button', { name: 'Close', exact: true }).click()
    await expect(options).toBeHidden()

    await clickMenu(app, ['Tools', 'Track Changes', 'Highlight Changes…'])
    const tracking = page.locator('.hlc-dialog')
    await expect(tracking).toBeVisible()
    await tracking.locator('input[name="track-changes"]').check()
    await tracking.locator('button.primary').click()
    await expect(tracking).toBeHidden()

    await caretAtDocEnd(page)
    await page.keyboard.press('ControlOrMeta+b')
    await page.keyboard.type('Quarterly results')
    await page.keyboard.press('Enter')
    await page.keyboard.type('Body under the heading')
    const made = page.locator('.doc-page :is(h1, h2, h3, h4, h5, h6)', {
      hasText: 'Quarterly results',
    })
    await expect(made).toHaveCount(1)
    await page.screenshot({ path: test.info().outputPath('define-styles-tracked.png') })
    await saveFromMenu(app)
    // typing that crosses a second is recorded as two insertions: the line is read through its runs
    const textOf = (xml: string) =>
      [...xml.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]).join('')
    const headingOf = (xml: string) =>
      (xml.match(/<w:p[ >].*?<\/w:p>/gs) ?? []).find((p) => textOf(p) === 'Quarterly results') ?? ''
    await expect
      .poll(async () => headingOf(await documentXml(docPath)), { timeout: 30_000 })
      .not.toBe('')
    const paragraph = headingOf(await documentXml(docPath))
    // the heading style, the paragraph as it was before it, and the typed text still an insertion
    expect(paragraph).toMatch(/<w:pStyle w:val="Heading\d"\/>/)
    expect(paragraph).toContain('<w:pPrChange')
    expect(paragraph).toContain('<w:ins ')
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR).catch(() => {})
    rmSync(dir, { recursive: true, force: true })
  }
})
