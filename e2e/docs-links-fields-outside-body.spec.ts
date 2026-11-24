/**
 * Links and pictures outside the body flow: the linked pictures of a header
 * (a floating logo and a picture in its line of text) are listed under
 * Edit > Links… and updated, relinked and broken there, each change read back
 * from the saved header part; Change Picture works on a picture in a text box.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { pathToFileURL } from 'node:url'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-links-fields-outside-body'
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const NS =
  `${W} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

const PNG_1PX =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='
const PNG_WIDE =
  'iVBORw0KGgoAAAANSUhEUgAAACgAAAAUCAIAAABwJOjsAAAAJ0lEQVR42mM8ISfHMBCAiWGAwKjFoxaPWjxq8ajFoxaPWjxqMf0BAF+MASwgP5VpAAAAAElFTkSuQmCC'
const PNG_SQUARE =
  'iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAIAAAAC64paAAAAHUlEQVR42mOUkzvBQC5gYqAAjGoe1TyqeVQz5ZoBOL0BLMMULg8AAAAASUVORK5CYII='

const SECT_PR = (refs = '') =>
  `<w:sectPr>${refs}<w:pgSz w:w="12240" w:h="15840"/>` +
  '<w:pgMar w:top="2160" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`

const graphic = (id: number, embed: string, link?: string) =>
  `<wp:extent cx="762000" cy="381000"/><wp:docPr id="${id}" name="Picture ${id}"/><wp:cNvGraphicFramePr/>` +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
  `<pic:blipFill><a:blip r:embed="${embed}"${link ? ` r:link="${link}"` : ''}/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="762000" cy="381000"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
  '</a:graphicData></a:graphic>'
const inlinePicture = (id: number, embed: string, link?: string) =>
  `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">${graphic(id, embed, link)}</wp:inline></w:drawing></w:r>`

/** a floating logo alone in its paragraph, then a picture in a line of text: both embedded and linked */
const HEADER =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${NS}>` +
  '<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>5486400</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="page"><wp:posOffset>228600</wp:posOffset></wp:positionV>' +
  `${graphic(31, 'rIdEmb', 'rIdLogo').replace('<wp:docPr', '<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr')}</wp:anchor></w:drawing></w:r></w:p>` +
  `<w:p>${run('Company ')}${inlinePicture(32, 'rIdEmb', 'rIdMark')}</w:p></w:hdr>`

const external = (id: string, file: string) =>
  `<Relationship Id="${id}" Type="${REL}/image" Target="${pathToFileURL(file).href}" TargetMode="External"/>`

/** the document with the header part, its relationships and the reference to it */
async function withHeader(docx: Buffer, rels: string): Promise<Buffer> {
  const zip = await JSZip.loadAsync(docx)
  zip.file('word/header1.xml', HEADER)
  zip.file(
    'word/_rels/header1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdEmb" Type="${REL}/image" Target="media/image1.png"/>${rels}</Relationships>`,
  )
  const docRels = 'word/_rels/document.xml.rels'
  zip.file(
    docRels,
    (await zip.file(docRels)!.async('string')).replace(
      '</Relationships>',
      `<Relationship Id="rIdHeader" Type="${REL}/header" Target="header1.xml"/></Relationships>`,
    ),
  )
  const types = '[Content_Types].xml'
  zip.file(
    types,
    (await zip.file(types)!.async('string')).replace(
      '</Types>',
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>',
    ),
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

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

/** each picture of the saved header: the file it is linked to (null: none) and the bytes it holds */
async function headerPictures(file: string) {
  const zip = await JSZip.loadAsync(readFileSync(file))
  const xml = (await zip.file('word/header1.xml')?.async('string')) ?? ''
  const rels = (await zip.file('word/_rels/header1.xml.rels')?.async('string')) ?? ''
  const rel = (id: string | undefined) =>
    id ? (new RegExp(`<Relationship\\b[^>]*Id="${id}"[^>]*>`).exec(rels)?.[0] ?? '') : ''
  const out: Array<{ link: string | null; bytes: string }> = []
  for (const blip of xml.matchAll(/<a:blip\b[^>]*>/g)) {
    const link = rel(/r:link="([^"]+)"/.exec(blip[0])?.[1])
    const media = /Target="([^"]+)"/.exec(rel(/r:embed="([^"]+)"/.exec(blip[0])?.[1]))?.[1]
    out.push({
      link: link.includes('TargetMode="External"')
        ? (/Target="([^"]*)"/.exec(link)?.[1] ?? '')
        : null,
      bytes: media ? ((await zip.file(`word/${media}`)?.async('base64')) ?? '') : '',
    })
  }
  return out
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

test('the linked pictures of a header are listed under Edit > Links, updated, relinked and broken', async () => {
  test.setTimeout(360_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-header-links-')))
  const docPath = join(dir, 'letterhead.docx')
  const logoPath = join(dir, 'logo.png')
  const markPath = join(dir, 'mark.png')
  const otherPath = join(dir, 'other.png')
  // the files hold other pictures than the one embedded: nothing reads them until asked
  writeFileSync(logoPath, Buffer.from(PNG_WIDE, 'base64'))
  writeFileSync(markPath, Buffer.from(PNG_SQUARE, 'base64'))
  writeFileSync(otherPath, Buffer.from(PNG_SQUARE, 'base64'))
  const source = await withHeader(
    await docxWith(
      para('Body first') +
        para('Body last') +
        SECT_PR('<w:headerReference w:type="default" r:id="rIdHeader"/>'),
    ),
    external('rIdLogo', logoPath) + external('rIdMark', markPath),
  )
  writeFileSync(docPath, source)
  let { launched, app, page } = await open(docPath)
  try {
    await expect(page.locator('.doc-page p', { hasText: 'Body last' })).toBeVisible({
      timeout: 30_000,
    })
    // the body has no link: the header's two open the dialog
    let dialog = await openLinks(app, page)
    let rows = dialog.locator('button.links-row')
    await expect(rows).toHaveCount(2)
    await expect(rows.nth(0)).toContainText('logo.png')
    await expect(rows.nth(0)).toContainText('Picture')
    await expect(rows.nth(1)).toContainText('mark.png')
    await page.screenshot({ path: test.info().outputPath('header-links.png') })
    // nothing was read yet: the file holds the one picture it embeds
    expect(await headerPictures(docPath)).toEqual([
      { link: pathToFileURL(logoPath).href, bytes: PNG_1PX },
      { link: pathToFileURL(markPath).href, bytes: PNG_1PX },
    ])

    // both at once: Update Now reads the two files
    await rows.nth(0).click()
    await rows.nth(1).click({ modifiers: ['Shift'] })
    await dialog.locator('button[name="update"]').click()
    await expect(rows.nth(0)).toContainText('Updated')
    await expect(rows.nth(1)).toContainText('Updated')
    await dialog.locator('button[name="close"]').click()
    await saveFromMenu(app)
    await expect
      .poll(() => headerPictures(docPath), { timeout: 30_000 })
      .toEqual([
        { link: pathToFileURL(logoPath).href, bytes: PNG_WIDE },
        { link: pathToFileURL(markPath).href, bytes: PNG_SQUARE },
      ])
    const header = await part(docPath, 'word/header1.xml')
    expect(header).not.toContain('aidocs-')
    expect(header).toContain('<w:t xml:space="preserve">Company </w:t>')
    expect(header.match(/<w:drawing>/g)).toHaveLength(2)

    // Change Source on the logo names the other file and takes its picture
    dialog = await openLinks(app, page)
    rows = dialog.locator('button.links-row')
    await rows.nth(0).click()
    await pickFile(app, otherPath)
    await dialog.locator('button[name="change"]').click()
    await expect(rows.nth(0)).toContainText('other.png')
    await expect(rows.nth(0)).toContainText('Updated')
    // Break Link on the picture in the line: the logo keeps its link
    await rows.nth(1).click()
    page.once('dialog', (confirm) => void confirm.accept())
    await dialog.locator('button[name="break"]').click()
    await expect(rows).toHaveCount(1)
    await expect(rows.nth(0)).toContainText('other.png')
    await dialog.locator('button[name="close"]').click()
    await saveFromMenu(app)
    await expect
      .poll(() => headerPictures(docPath), { timeout: 30_000 })
      .toEqual([
        { link: pathToFileURL(otherPath).href, bytes: PNG_SQUARE },
        { link: null, bytes: PNG_SQUARE },
      ])
    expect(await part(docPath, 'word/header1.xml')).not.toContain('aidocs-')

    // read again from the file: one link is left
    await closeAndSaveVideo(launched, VIDEO_DIR)
    ;({ launched, app, page } = await open(docPath))
    await expect(page.locator('.doc-page p', { hasText: 'Body last' })).toBeVisible({
      timeout: 30_000,
    })
    dialog = await openLinks(app, page)
    await expect(dialog.locator('button.links-row')).toHaveCount(1)
    await expect(dialog.locator('button.links-row').nth(0)).toContainText('other.png')
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR).catch(() => {})
    rmSync(dir, { recursive: true, force: true })
  }
})

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
  `<wps:txbx><w:txbxContent><w:p>${run('Boxed ')}${inlinePicture(12, 'rIdImg')}</w:p></w:txbxContent></wps:txbx>` +
  '<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r></w:p>'

test('Change Picture works on a picture in a text box: the box keeps it in its line', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-box-picture-')))
  const docPath = join(dir, 'boxed.docx')
  writeFileSync(
    docPath,
    await docxWith(para('Host first') + TEXTBOX + para('Host last') + SECT_PR()),
  )
  const { launched, app, page } = await open(docPath)
  try {
    await expect(page.locator('.doc-page p', { hasText: 'Host last' })).toBeVisible({
      timeout: 30_000,
    })
    const picture = page.locator('.doc-page .doc-textbox img.doc-inline-img')
    await expect(picture).toHaveCount(1)
    const src = await picture.getAttribute('src')
    // handed to the page as the clipboard's content: the system clipboard is shared with
    // whatever else runs on the machine
    await page.evaluate(async () => {
      const canvas = document.createElement('canvas')
      canvas.width = canvas.height = 20
      const ctx = canvas.getContext('2d')!
      ctx.fillStyle = 'rgb(0, 120, 220)'
      ctx.fillRect(0, 0, 20, 20)
      const png = await new Promise<Blob>((resolve) => canvas.toBlob((blob) => resolve(blob!)))
      navigator.clipboard.read = async () => [new ClipboardItem({ 'image/png': png })]
    })
    await picture.click()
    await picture.click({ button: 'right' })
    const menu = page.locator('.ctx-menu')
    await menu.waitFor({ timeout: 15_000 })
    const item = menu.locator('.ctx-item', { hasText: 'Change Picture' })
    await expect(item).toBeEnabled()
    await item.hover()
    const items = menu.locator('.ctx-submenu .ctx-item')
    await expect(items.locator('.ctx-label')).toHaveText([
      'From a File…',
      'From Icons…',
      'From Clipboard…',
    ])
    await page.screenshot({ path: test.info().outputPath('box-change-picture-menu.png') })
    await items.nth(2).click()
    await expect.poll(() => picture.getAttribute('src'), { timeout: 15_000 }).not.toBe(src)
    await expect(picture).toHaveCount(1)
    await expect(page.locator('.doc-page .doc-textbox').first()).toContainText('Boxed')
    // the box hands its content to the document when the caret leaves it
    await page.locator('.doc-page p', { hasText: 'Host last' }).click()

    await saveFromMenu(app)
    const media = async () =>
      (await part(docPath, 'word/_rels/document.xml.rels')).match(/media\//g) ?? []
    await expect.poll(async () => (await media()).length, { timeout: 30_000 }).toBeGreaterThan(1)
    const xml = await part(docPath, 'word/document.xml')
    const box = /<w:txbxContent>.*?<\/w:txbxContent>/s.exec(xml)?.[0] ?? ''
    expect(box).toContain('Boxed ')
    expect(box.match(/<w:drawing>/g)).toHaveLength(1)
    expect(box).not.toContain('r:embed="rIdImg"')
    expect(xml).not.toContain('aidocs-new-pic-')
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR).catch(() => {})
    rmSync(dir, { recursive: true, force: true })
  }
})
