/**
 * The picture and table right-click items that used to be greyed out, and the
 * two small commands that were: View Alt Text, Auto Crop, Size and Position,
 * Format Picture, the table's Borders and Shading, Edit > Toggle Drawing and
 * Design > Set as Default.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdirSync, mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { crc32, deflateSync } from 'node:zlib'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

/** a white PNG of w x h with one red box on it */
function borderedPng(
  w: number,
  h: number,
  box: { x: number; y: number; w: number; h: number },
): Buffer {
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

const pngSize = (bytes: Buffer) => [bytes.readUInt32BE(16), bytes.readUInt32BE(20)]

const emu = (px: number) => px * 9525

/** a paragraph that is one in-line picture, size in px */
const picturePara = (w: number, h: number) =>
  `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
  `<wp:extent cx="${emu(w)}" cy="${emu(h)}"/><wp:docPr id="1" name="Picture 1"/>` +
  `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
  `<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="Picture 1"/><pic:cNvPicPr/></pic:nvPicPr>` +
  `<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
  `</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`

const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
  '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'

async function fixture(): Promise<Buffer> {
  const zip = await JSZip.loadAsync(
    await docxWith(para('One') + picturePara(200, 150) + TABLE + para('Last')),
  )
  zip.file('word/media/image1.png', borderedPng(40, 30, { x: 10, y: 5, w: 20, h: 15 }))
  return zip.generateAsync({ type: 'nodebuffer' })
}

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

async function pictureItem(page: Page, label: RegExp) {
  const img = page.locator('.doc-page[contenteditable="true"] img').first()
  await img.click({ button: 'right' })
  const menu = page.locator('.ctx-menu')
  await menu.waitFor({ timeout: 15_000 })
  const item = menu.locator('.ctx-item', { hasText: label })
  await expect(item).toBeEnabled()
  await item.click()
}

const ready = (page: Page) =>
  page.waitForFunction(
    () => Boolean((window as unknown as { __aidocs?: { editor?: unknown } }).__aidocs?.editor),
    undefined,
    { timeout: 30_000 },
  )

test('a picture is cropped to its content, described, sized and turned from its right-click menu', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-picture-items-')))
  const docPath = join(dir, 'picture.docx')
  writeFileSync(docPath, await fixture())
  let launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-picture-alt-size-crop',
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await ready(page)
    const img = page.locator('.doc-page[contenteditable="true"] img').first()
    await expect(img).toBeVisible({ timeout: 30_000 })
    // polled: the picture's element is replaced when its bytes or size change
    const shown = async () => {
      const box = await img.boundingBox().catch(() => null)
      return box ? [Math.round(box.width), Math.round(box.height)] : null
    }
    await expect.poll(shown).toEqual([200, 150])

    // Auto Crop: the white border goes, what is left keeps its scale
    await pictureItem(page, /^Auto Crop$/)
    await expect.poll(shown).toEqual([100, 75])
    // nothing left to crop: the picture stays and the status bar says so
    await pictureItem(page, /^Auto Crop$/)
    await expect(page.getByText('The picture has no border to crop')).toBeVisible()
    await expect.poll(shown).toEqual([100, 75])

    await pictureItem(page, /^View Alt Text/)
    const alt = page.locator('.alt-text-dialog')
    await expect(alt.locator('textarea')).toHaveValue('')
    await alt.locator('textarea').fill('A red kite & its "tail"')
    await page.screenshot({ path: test.info().outputPath('alt-text.png') })
    await alt.locator('button.primary').click()
    await expect(alt).toBeHidden()

    await pictureItem(page, /^Size and Position/)
    const size = page.locator('.size-position-dialog')
    const widthBox = size.getByLabel('Width', { exact: true })
    await widthBox.fill('144pt')
    await widthBox.blur()
    // the aspect ratio is locked: the height follows
    await expect(size.getByLabel('Height', { exact: true })).toHaveValue(/^(1\.5"|3\.81 cm)$/)
    // unlocked, the height changes alone; locked again, the width follows at the new ratio
    const lock = size.locator('input[name="lock"]')
    const heightBox = size.getByLabel('Height', { exact: true })
    await lock.uncheck()
    await heightBox.fill('72pt')
    await heightBox.blur()
    await expect(widthBox).toHaveValue(/^(2"|5\.08 cm)$/)
    await lock.check()
    await heightBox.fill('108pt')
    await heightBox.blur()
    await expect(widthBox).toHaveValue(/^(3"|7\.62 cm)$/)
    await widthBox.fill('144pt')
    await widthBox.blur()
    await expect(heightBox).toHaveValue(/^(1"|2\.54 cm)$/)
    await size.locator('input[name="rotation"]').fill('90')
    await page.screenshot({ path: test.info().outputPath('size-position.png') })
    await size.locator('button.primary').click()
    await expect(size).toBeHidden()

    await pictureItem(page, /^Format Picture/)
    await expect(page.locator('.ribbon-tab.active')).toHaveText('Picture Format')

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const zip = await JSZip.loadAsync(readFileSync(docPath))
          const xml = await zip.file('word/document.xml')!.async('string')
          const media = Object.keys(zip.files).filter((f) => /^word\/media\/.+\.png$/.test(f))
          const sizes = await Promise.all(
            media.map(async (f) => pngSize(await zip.file(f)!.async('nodebuffer')).join('x')),
          )
          return [
            /<wp:docPr\b[^>]* descr="A red kite &amp; its &quot;tail&quot;"/.test(xml),
            /<wp:extent cx="1828800" cy="914400"/.test(xml),
            /<a:xfrm[^>]* rot="5400000"/.test(xml),
            sizes.includes('20x15'),
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([true, true, true, true])
  } finally {
    await closeAndSaveVideo(launched, 'docs-picture-alt-size-crop')
  }

  launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-picture-alt-size-crop',
    openFile: docPath,
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await ready(page)
    await expect(page.locator('.doc-page[contenteditable="true"] img').first()).toBeVisible({
      timeout: 30_000,
    })
    await pictureItem(page, /^View Alt Text/)
    await expect(page.locator('.alt-text-dialog textarea')).toHaveValue('A red kite & its "tail"')
  } finally {
    await closeAndSaveVideo(launched, 'docs-picture-alt-size-crop')
    rmSync(dir, { recursive: true, force: true })
  }
})

test("the table's Borders and Shading, Toggle Drawing and Set as Default run", async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-picture-items-')))
  const docPath = join(dir, 'table.docx')
  writeFileSync(docPath, await fixture())
  // one profile for both launches: the default look is kept per user
  const userDataDir = join(dir, 'profile')
  mkdirSync(userDataDir)
  let launched = await launchShell({
    onboardingSeen: true,
    userDataDir,
    videoDir: 'docs-picture-alt-size-crop',
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await ready(page)
    const cell = page.locator('.doc-page td', { hasText: 'cell' }).first()
    await expect(cell).toBeVisible({ timeout: 30_000 })
    // the caret is in the cell before the menu opens: the table items go by the selection
    await cell.click()
    await cell.click({ button: 'right' })
    const menu = page.locator('.ctx-menu')
    await menu.waitFor({ timeout: 15_000 })
    await menu.locator('.ctx-item', { hasText: /^Borders and Shading/ }).click()
    const borders = page.locator('.borders-shading-dialog')
    await expect(borders).toBeVisible()
    // the dialog has the focus before Escape closes it
    await expect(borders.locator(':focus')).toHaveCount(1)
    await page.keyboard.press('Escape')
    await expect(borders).toBeHidden()

    // Edit > Toggle Drawing: the pen on, then the text caret again
    const ink = page.locator('.ink-overlay.ink-active')
    await clickMenu(app, ['Edit', 'Toggle Drawing'])
    await expect(ink).toHaveCount(1)
    await expect(page.locator('.ribbon-tab.active')).toHaveText('Draw')
    await clickMenu(app, ['Edit', 'Toggle Drawing'])
    await expect(ink).toHaveCount(0)

    // Design > Set as Default keeps the look: a new blank document starts with it
    await page.locator('.doc-page p', { hasText: 'One' }).first().click()
    await page.locator('.ribbon-tab', { hasText: /^Design$/ }).click()
    await page.locator('.ribbon button', { hasText: /^Paragraph Spacing$/ }).click()
    await page.locator('button', { hasText: /^Double$/ }).click()
    const setDefault = page.locator('.ribbon button', { hasText: /^Set as Default$/ })
    await expect(setDefault).toBeEnabled()
    await setDefault.click()
    await expect(
      page.getByText(
        'New blank documents will start with this theme, style set and paragraph spacing',
      ),
    ).toBeVisible()
    const kept = await page.evaluate(() => localStorage.getItem('aidocs.designDefault'))
    expect(JSON.parse(kept ?? '{}').spacing).toEqual({ afterTwips: 160, line: 2 })
  } finally {
    await closeAndSaveVideo(launched, 'docs-picture-alt-size-crop')
  }

  launched = await launchShell({
    onboardingSeen: true,
    userDataDir,
    videoDir: 'docs-picture-alt-size-crop',
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await ready(page)
    await expect(page.locator('.doc-page td', { hasText: 'cell' }).first()).toBeVisible({
      timeout: 30_000,
    })
    await clickMenu(app, ['File', 'New Document'])
    // the blank document, wherever it opened: 160 twips after each paragraph, double spacing
    const blankSpacing = async () => {
      for (const candidate of app.windows()) {
        if (!candidate.url().includes('://docs/')) continue
        const found = await candidate
          .evaluate(() => {
            const paper = document.querySelector('.doc-page[contenteditable="true"]')
            const first = paper?.querySelector('p')
            if (!paper || !first || (paper.textContent ?? '').trim() !== '') return null
            const style = getComputedStyle(first)
            return [
              Math.round(parseFloat(style.marginBottom) * 10) / 10,
              Math.round((parseFloat(style.lineHeight) / parseFloat(style.fontSize)) * 10) / 10 >=
                2,
            ]
          })
          .catch(() => null)
        if (found) return found
      }
      return null
    }
    await expect.poll(blankSpacing, { timeout: 30_000 }).toEqual([10.7, true])
  } finally {
    await closeAndSaveVideo(launched, 'docs-picture-alt-size-crop')
    rmSync(dir, { recursive: true, force: true })
  }
})
