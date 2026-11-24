/**
 * Object formatting that used to be read past: a shape sized as a share of
 * the page (wp14:sizeRelH / sizeRelV), a gradient or transparent text fill
 * (w14:textFill), Auto Crop on a picture in a line of text, and the density a
 * BMP states. Each is driven from the UI and checked in the saved file.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { crc32, deflateSync } from 'node:zlib'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-relative-size-text-fill'

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

// US Letter with one-inch margins: the text area is 6.5 in, 624 px
const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'

async function docx(body: string, media?: { name: string; bytes: Buffer }): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="bmp" ContentType="image/bmp"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  if (media) {
    zip.file(
      'word/_rels/document.xml.rels',
      `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${media.name}"/></Relationships>`,
    )
    zip.file(`word/media/${media.name}`, media.bytes)
  }
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body>${body}${SECT_PR}</w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

const para = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
const run = (text: string, rPr = '') =>
  `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`
const emu = (px: number) => px * 9525

const relSize = (axis: 'H' | 'V', from: string, units: number) => {
  const pct = axis === 'H' ? 'pctWidth' : 'pctHeight'
  return `<wp14:sizeRel${axis} relativeFrom="${from}"><wp14:${pct}>${units}</wp14:${pct}></wp14:sizeRel${axis}>`
}

/** a paragraph anchoring one text box of w x h px at `y` px below it; `tail` is its relative size */
const textBox = (id: number, text: string, y: number, w: number, h: number, tail: string) =>
  `<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" ` +
  `relativeHeight="${251659264 + id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
  `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>` +
  `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${emu(y)}</wp:posOffset></wp:positionV>` +
  `<wp:extent cx="${emu(w)}" cy="${emu(h)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
  `<wp:wrapNone/><wp:docPr id="${id}" name="Text Box ${id}"/><wp:cNvGraphicFramePr/>` +
  `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
  `<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="DDEBF7"/></a:solidFill></wps:spPr>` +
  `<wps:txbx><w:txbxContent>${para(text)}</w:txbxContent></wps:txbx>` +
  `<wps:bodyPr><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic>${tail}</wp:anchor></w:drawing></w:r></w:p>`

function png(w: number, h: number, box: { x: number; y: number; w: number; h: number }): Buffer {
  const raw = Buffer.alloc((w * 4 + 1) * h)
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const red = x >= box.x && x < box.x + box.w && y >= box.y && y < box.y + box.h
      raw.set(red ? [200, 0, 0, 255] : [255, 255, 255, 255], y * (w * 4 + 1) + 1 + x * 4)
    }
  }
  const chunk = (type: string, data: Buffer) => {
    const body = Buffer.concat([Buffer.from(type, 'ascii'), data])
    const out = Buffer.alloc(body.length + 8)
    out.writeUInt32BE(data.length, 0)
    body.copy(out, 4)
    out.writeUInt32BE(crc32(body), body.length + 4)
    return out
  }
  const head = Buffer.alloc(13)
  head.writeUInt32BE(w, 0)
  head.writeUInt32BE(h, 4)
  head.set([8, 6, 0, 0, 0], 8)
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', head),
    chunk('IDAT', deflateSync(raw)),
    chunk('IEND', Buffer.alloc(0)),
  ])
}

/** a 24-bit BMP of w x h that states `dpi` in its info header */
function bmp(w: number, h: number, dpi: number): Buffer {
  const row = Math.ceil((w * 3) / 4) * 4
  const out = Buffer.alloc(54 + row * h, 0x80)
  out.write('BM', 0, 'ascii')
  out.writeUInt32LE(out.length, 2)
  out.writeUInt32LE(0, 6)
  out.writeUInt32LE(54, 10)
  out.writeUInt32LE(40, 14)
  out.writeInt32LE(w, 18)
  out.writeInt32LE(h, 22)
  out.writeUInt16LE(1, 26)
  out.writeUInt16LE(24, 28)
  out.writeUInt32LE(0, 30)
  out.writeUInt32LE(row * h, 34)
  const ppm = Math.round(dpi / 0.0254)
  out.writeInt32LE(ppm, 38)
  out.writeInt32LE(ppm, 42)
  out.writeUInt32LE(0, 46)
  out.writeUInt32LE(0, 50)
  return out
}

const picture = (kind: 'inline', w: number, h: number) =>
  `<w:r><w:drawing><wp:${kind} distT="0" distB="0" distL="0" distR="0">` +
  `<wp:extent cx="${emu(w)}" cy="${emu(h)}"/><wp:docPr id="9" name="Picture 9" descr="kept words"/>` +
  `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
  `<pic:pic><pic:nvPicPr><pic:cNvPr id="9" name="Picture 9"/><pic:cNvPicPr/></pic:nvPicPr>` +
  `<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
  `</a:graphicData></a:graphic></wp:${kind}></w:drawing></w:r>`

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

const documentXml = async (docPath: string) =>
  (await (
    await JSZip.loadAsync(readFileSync(docPath))
  )
    .file('word/document.xml')
    ?.async('string')) ?? ''

async function openDoc(docPath: string, text: string) {
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
  await expect(page.locator('.doc-page p', { hasText: text }).first()).toBeVisible({
    timeout: 30_000,
  })
  return { launched, page }
}

/** opens Size and Position on the first text box: its frame is selected, then right-clicked at its edge */
async function openShapeSize(page: Page) {
  await page.evaluate(() => {
    const { editor } = (
      window as never as {
        __aidocs: {
          editor: {
            state: {
              doc: {
                descendants(f: (node: { type: { name: string } }, pos: number) => void): void
              }
            }
            chain(): { focus(): { setNodeSelection(pos: number): { run(): boolean } } }
          }
        }
      }
    ).__aidocs
    let at = -1
    editor.state.doc.descendants((node, pos) => {
      if (at < 0 && node.type.name === 'docProtected') at = pos
    })
    // focused, so that an open still about to put its caret at the start leaves the selection alone
    editor.chain().focus().setNodeSelection(at).run()
  })
  const box = page.locator('.doc-page .doc-textbox').first()
  await box.click({ button: 'right', position: { x: 2, y: 2 } })
  const menu = page.locator('.ctx-menu')
  await menu.waitFor({ timeout: 15_000 })
  const item = menu.locator('.ctx-item', { hasText: /^Size and Position/ })
  await expect(item).toBeEnabled()
  await item.click()
}

/** the item of the right-click menu on `target`, once it can be used */
async function contextItem(page: Page, target: ReturnType<Page['locator']>, label: RegExp) {
  await target.click({ button: 'right' })
  const menu = page.locator('.ctx-menu')
  await menu.waitFor({ timeout: 15_000 })
  const item = menu.locator('.ctx-item', { hasText: label })
  await expect(item).toBeEnabled()
  await item.click()
}

test('a shape takes a share of the page from Size and Position, keeps it in the file, and follows the margins', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-rel-size-')))
  const docPath = join(dir, 'relative.docx')
  // the second box states half the text area but caches 200 px, as files in the wild do
  const stale = textBox(
    2,
    'Stale',
    150,
    200,
    60,
    relSize('H', 'margin', 50000) + relSize('V', 'margin', 0),
  )
  writeFileSync(
    docPath,
    await docx(
      para('Above') +
        textBox(1, 'Sized', 20, 200, 80, relSize('H', 'margin', 0) + relSize('V', 'margin', 0)) +
        stale +
        para('Below'),
    ),
  )
  const { launched, page } = await openDoc(docPath, 'Above')
  try {
    const { app } = launched
    const boxes = page.locator('.doc-page .doc-textbox')
    await expect(boxes).toHaveCount(2)
    // a file's cached size is what is drawn, whatever share it states
    const zoom = await page.evaluate(() => {
      const el = document.querySelector<HTMLElement>('.doc-page .doc-textbox')!
      return el.getBoundingClientRect().width / el.offsetWidth
    })
    const widths = async () =>
      (
        await page.evaluate(() =>
          Array.from(document.querySelectorAll<HTMLElement>('.doc-page .doc-textbox')).map(
            (el) => el.getBoundingClientRect().width,
          ),
        )
      ).map((w) => Math.round(w / zoom))
    await expect.poll(widths).toEqual([200, 200])

    await openShapeSize(page)
    const dialog = page.locator('.shape-size-dialog')
    await expect(dialog).toBeVisible()
    const width = dialog.locator('[data-side="width"]')
    await expect(width.locator('input[value="absolute"]')).toBeChecked()
    await expect(width.locator('input[name="widthPct"]')).toBeDisabled()
    await width.locator('input[value="relative"]').check()
    await width.locator('input[name="widthPct"]').fill('50')
    await expect(width.locator('select[name="widthFrom"]')).toHaveValue('margin')
    // the height stays absolute at a new value
    const height = dialog.locator('[data-side="height"]')
    const heightBox = height.getByRole('textbox').first()
    await heightBox.fill('72pt')
    await heightBox.blur()
    await page.screenshot({ path: test.info().outputPath('shape-size-dialog.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect.poll(widths).toEqual([312, 200])

    // reopened, the dialog shows the share
    await openShapeSize(page)
    await expect(dialog.locator('[data-side="width"] input[value="relative"]')).toBeChecked()
    await expect(dialog.locator('input[name="widthPct"]')).toHaveValue('50')
    await dialog.getByRole('button', { name: 'Cancel' }).click()

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const xml = await documentXml(docPath)
          const first = xml.slice(0, xml.indexOf('Stale'))
          return [
            first.includes(relSize('H', 'margin', 50000) + relSize('V', 'margin', 0)),
            first.includes(`<wp:extent cx="${emu(312)}" cy="${emu(96)}"/>`),
            // the box nobody touched keeps its bytes, stale cache and all
            xml.includes(stale),
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([true, true, true])

    // narrower margins: both shares are worked out again against the wider text area (7.5 in)
    const margins = page
      .locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: 'Margins' })
      .first()
    if (!(await margins.isVisible()))
      await page.locator('.ribbon-tab', { hasText: /^Layout$/ }).click()
    await margins.click()
    await page
      .locator('[data-rb-panel] button', { hasText: /^Narrow/ })
      .first()
      .click()
    await expect.poll(widths, { timeout: 20_000 }).toEqual([360, 360])
    await clickMenu(app, ['Edit', 'Undo'])
    await expect.poll(widths, { timeout: 20_000 }).toEqual([312, 312])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

const gradFill =
  '<w14:textFill><w14:gradFill><w14:gsLst>' +
  '<w14:gs w14:pos="0"><w14:srgbClr w14:val="FF0000"/></w14:gs>' +
  '<w14:gs w14:pos="100000"><w14:srgbClr w14:val="0000FF"/></w14:gs>' +
  '</w14:gsLst><w14:lin w14:ang="0" w14:scaled="0"/></w14:gradFill></w14:textFill>'
const faintFill =
  '<w14:textFill><w14:solidFill><w14:srgbClr w14:val="C00000"><w14:alpha w14:val="60000"/></w14:srgbClr></w14:solidFill></w14:textFill>'
const otherFill =
  '<w14:textFill><w14:solidFill><w14:srgbClr w14:val="00B050"/></w14:solidFill></w14:textFill>'

const drawn = (page: Page, text: string) =>
  page.evaluate((text) => {
    const span = Array.from(document.querySelectorAll<HTMLElement>('.doc-page p span')).find(
      (el) => el.textContent === text,
    )
    if (!span) return null
    const style = getComputedStyle(span)
    return {
      fill: style.getPropertyValue('-webkit-text-fill-color'),
      image: style.backgroundImage,
      clip: style.getPropertyValue('-webkit-background-clip'),
    }
  }, text)

const selectText = (page: Page, from: number, to: number) =>
  page.evaluate(
    ([from, to]) => {
      const { editor } = (
        window as never as {
          __aidocs: {
            editor: { chain(): { focus(): { setTextSelection(r: object): { run(): boolean } } } }
          }
        }
      ).__aidocs
      editor.chain().focus().setTextSelection({ from, to }).run()
    },
    [from, to],
  )

test('a gradient, a transparent and a differing text fill are drawn; Text Effects puts a gradient on and takes it off', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-text-fill-')))
  const docPath = join(dir, 'fill.docx')
  const fileRuns =
    run('blend', `<w:color w:val="FF0000"/>${gradFill}`) +
    run(' ') +
    run('faint', `<w:color w:val="C00000"/>${faintFill}`) +
    run(' ') +
    run('green', `<w:color w:val="000000"/>${otherFill}`)
  writeFileSync(docPath, await docx(`<w:p>${run('Plain word ')}${fileRuns}</w:p>${para('Last')}`))
  const { launched, page } = await openDoc(docPath, 'Plain word')
  try {
    const { app } = launched
    await expect
      .poll(() => drawn(page, 'blend'))
      .toEqual({
        fill: 'rgba(0, 0, 0, 0)',
        image: 'linear-gradient(90deg, rgb(255, 0, 0) 0%, rgb(0, 0, 255) 100%)',
        clip: 'text',
      })
    expect((await drawn(page, 'faint'))?.fill).toBe('rgba(192, 0, 0, 0.4)')
    // Word shows the fill, not the w:color beside it
    expect((await drawn(page, 'green'))?.fill).toBe('rgb(0, 176, 80)')
    await page.screenshot({ path: test.info().outputPath('text-fill-page.png') })

    // "word" (7..11) takes a gradient of the dialog
    await selectText(page, 7, 11)
    await clickMenu(app, ['Format', 'Text Effects…'])
    const dialog = page.locator('.text-fx-dialog')
    await expect(dialog).toBeVisible()
    const gradients = dialog.getByRole('radiogroup', { name: 'Gradient' })
    await expect(gradients.getByRole('radio', { name: 'None' })).toBeChecked()
    await dialog.locator('[data-fx="fill"] .text-fx-swatch[aria-label="#4472C4"]').click()
    await gradients.getByRole('radio', { name: 'Linear Down' }).click()
    await expect(gradients.getByRole('radio', { name: 'Linear Down' })).toBeChecked()
    await page.screenshot({ path: test.info().outputPath('text-effects-gradient.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect
      .poll(async () => (await drawn(page, 'word'))?.image)
      .toBe('linear-gradient(rgb(180, 199, 231) 0%, rgb(68, 114, 196) 100%)')

    await saveFromMenu(app)
    const written =
      '<w:color w:val="4472C4"/><w14:textFill><w14:gradFill><w14:gsLst>' +
      '<w14:gs w14:pos="0"><w14:srgbClr w14:val="B4C7E7"/></w14:gs>' +
      '<w14:gs w14:pos="100000"><w14:srgbClr w14:val="4472C4"/></w14:gs>' +
      '</w14:gsLst><w14:lin w14:ang="5400000" w14:scaled="0"/></w14:gradFill></w14:textFill>'
    await expect
      .poll(
        async () => {
          const xml = await documentXml(docPath)
          return [
            xml.includes(written),
            // the fills that came with the file are kept as they were
            xml.includes(`<w:color w:val="FF0000"/>${gradFill}`),
            xml.includes(`<w:color w:val="C00000"/>${faintFill}`),
            xml.includes(`<w:color w:val="000000"/>${otherFill}`),
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([true, true, true, true])

    // the dialog shows the gradient it put on, and takes it off again
    await selectText(page, 7, 11)
    await clickMenu(app, ['Format', 'Text Effects…'])
    await expect(gradients.getByRole('radio', { name: 'Linear Down' })).toBeChecked()
    await gradients.getByRole('radio', { name: 'None' }).click()
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect.poll(async () => (await drawn(page, 'word'))?.image).toBe('none')
    await clickMenu(app, ['File', 'Save'])
    await expect
      .poll(
        async () => {
          const xml = await documentXml(docPath)
          return [
            xml.includes(written),
            xml.includes(
              '<w:color w:val="4472C4"/><w14:textFill><w14:solidFill><w14:srgbClr w14:val="4472C4"/></w14:solidFill></w14:textFill>',
            ),
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([false, true])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
  }

  // reopened, the fills read back the way they were drawn
  const again = await openDoc(docPath, 'Plain word')
  try {
    await expect
      .poll(async () => (await drawn(again.page, 'blend'))?.image)
      .toBe('linear-gradient(90deg, rgb(255, 0, 0) 0%, rgb(0, 0, 255) 100%)')
    expect((await drawn(again.page, 'faint'))?.fill).toBe('rgba(192, 0, 0, 0.4)')
  } finally {
    await closeAndSaveVideo(again.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Auto Crop trims a picture in a line of text, and a BMP is scaled from the size its header states', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-inline-crop-')))
  const docPath = join(dir, 'inline.docx')
  writeFileSync(
    docPath,
    await docx(
      `<w:p>${run('Before ')}${picture('inline', 200, 150)}${run(' after')}</w:p>${para('Last')}`,
      { name: 'image1.png', bytes: png(40, 30, { x: 10, y: 5, w: 20, h: 15 }) },
    ),
  )
  const { launched, page } = await openDoc(docPath, 'Before')
  try {
    const { app } = launched
    const img = page.locator('.doc-page[contenteditable="true"] img').first()
    await expect(img).toBeVisible({ timeout: 30_000 })
    const shown = async () => {
      const size = await img
        .evaluate((el: HTMLElement) => [el.offsetWidth, el.offsetHeight])
        .catch(() => null)
      return size
    }
    await expect.poll(shown).toEqual([200, 150])
    await contextItem(page, img, /^Auto Crop$/)
    // the white border goes; what is left keeps its scale, in its line of text
    await expect.poll(shown).toEqual([100, 75])
    await expect(page.locator('.doc-page p', { hasText: 'Before' }).first()).toContainText('after')
    await contextItem(page, img, /^Auto Crop$/)
    await expect(page.getByText('The picture has no border to crop')).toBeVisible()

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const zip = await JSZip.loadAsync(readFileSync(docPath))
          const xml = (await zip.file('word/document.xml')?.async('string')) ?? ''
          const media = Object.keys(zip.files).filter((f) => /^word\/media\/.+\.png$/.test(f))
          const sizes = await Promise.all(
            media.map(async (f) => {
              const bytes = await zip.file(f)!.async('nodebuffer')
              return `${bytes.readUInt32BE(16)}x${bytes.readUInt32BE(20)}`
            }),
          )
          return [
            /<wp:inline\b[^>]*><wp:extent cx="952500" cy="714375"\/>/.test(xml),
            // the drawing is the file's own, pointed at the new bytes
            xml.includes('descr="kept words"'),
            sizes.includes('20x15'),
            /Before[\s\S]*<wp:inline[\s\S]*after/.test(xml),
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([true, true, true, true])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
  }

  // 60 x 30 px at 300 dpi is 19.2 x 9.6 px on the page; shown at 96 x 48 that is 500%
  const bmpPath = join(dir, 'bmp.docx')
  writeFileSync(
    bmpPath,
    await docx(`${para('One')}<w:p>${picture('inline', 96, 48)}</w:p>${para('Last')}`, {
      name: 'image1.bmp',
      bytes: bmp(60, 30, 300),
    }),
  )
  const second = await openDoc(bmpPath, 'One')
  try {
    const img = second.page.locator('.doc-page[contenteditable="true"] img').first()
    await expect(img).toBeVisible({ timeout: 30_000 })
    await img.click()
    await contextItem(second.page, img, /^Size and Position/)
    const size = second.page.locator('.size-position-dialog')
    // at 96 dpi this picture would read 160%
    await expect(size.locator('input[name="scaleWidth"]')).toHaveValue('500')
    await expect(size.locator('input[name="scaleHeight"]')).toHaveValue('500')
    await second.page.screenshot({ path: test.info().outputPath('bmp-scale.png') })
  } finally {
    await closeAndSaveVideo(second.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
