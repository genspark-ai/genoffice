/**
 * Pictures, shapes and text boxes where the shipped commands stopped short:
 * Align on the Shape Format and Picture Format tabs, a group that follows the
 * pointer whole while it is dragged, Cut and Copy of the whole group, View Alt
 * Text on a picture in a line of text, and Scale and the preset places in
 * Size and Position. Each is checked on the page and in the saved file.
 */
import { test, expect, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { crc32, deflateSync } from 'node:zlib'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import {
  docxWith,
  para,
  pictureRun,
  rects,
  saveFromMenu,
  shapeParagraph,
  type Rect,
} from './docs-objects-fixture'

const emu = (px: number) => px * 9525

/** a plain PNG of w x h */
function png(w: number, h: number): Buffer {
  const raw = Buffer.alloc((w * 4 + 1) * h)
  for (let y = 0; y < h; y++)
    for (let x = 0; x < w; x++) raw.set([200, 0, 0, 255], y * (w * 4 + 1) + 1 + x * 4)
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

/** a run that is one in-line picture, size in px */
const inlinePicture = (id: number, w: number, h: number, descr = '') =>
  `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
  `<wp:extent cx="${emu(w)}" cy="${emu(h)}"/><wp:docPr id="${id}" name="Picture ${id}"${descr ? ` descr="${descr}"` : ''}/>` +
  `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
  `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
  `<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
  `</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`

const documentXml = async (file: string): Promise<string> => {
  const zip = await JSZip.loadAsync(readFileSync(file))
  return (await zip.file('word/document.xml')?.async('string')) ?? ''
}

const ready = (page: Page) =>
  page.waitForFunction(
    () => Boolean((window as unknown as { __aidocs?: { editor?: unknown } }).__aidocs?.editor),
    undefined,
    { timeout: 30_000 },
  )

const PICTURES = '.doc-page[contenteditable="true"] img'

/** opens the right-click menu of the nth picture and runs one of its items */
async function pictureItem(page: Page, nth: number, label: RegExp) {
  await page.locator(PICTURES).nth(nth).click({ button: 'right' })
  const menu = page.locator('.ctx-menu')
  await menu.waitFor({ timeout: 15_000 })
  const item = menu.locator('.ctx-item', { hasText: label })
  await expect(item).toBeEnabled()
  await item.click()
}

test('a picture in a line of text is described; Size and Position scales a picture and puts it at a preset place', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-object-gaps-')))
  const docPath = join(dir, 'pictures.docx')
  const zip = await JSZip.loadAsync(
    await docxWith(
      para('One') +
        `<w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r>${inlinePicture(1, 80, 60, 'Old words')}<w:r><w:t xml:space="preserve"> after.</w:t></w:r></w:p>` +
        `<w:p>${inlinePicture(2, 200, 150)}</w:p>` +
        para('Last'),
    ),
  )
  zip.file('word/media/image1.png', png(40, 30))
  writeFileSync(docPath, await zip.generateAsync({ type: 'nodebuffer' }))
  /** the corner of the picture alone in its paragraph, measured from the paper's */
  const onPaper = async (page: Page): Promise<number[]> => {
    const [paper] = await rects(page, '.doc-page[contenteditable="true"]')
    const [picture] = await rects(page, '.doc-page[contenteditable="true"] .doc-protected img')
    return [picture.left - paper.left, picture.top - paper.top]
  }
  let placed: number[] = []
  let launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-floating-object-gaps',
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await ready(page)
    const pictures = page.locator(PICTURES)
    await expect(pictures).toHaveCount(2, { timeout: 30_000 })
    const shown = async (nth: number) => {
      const box = await pictures
        .nth(nth)
        .boundingBox()
        .catch(() => null)
      return box ? [Math.round(box.width), Math.round(box.height)] : null
    }
    await expect.poll(() => shown(0)).toEqual([80, 60])

    // the picture among text: its alternative text is the file's, and what is typed replaces it
    await pictureItem(page, 0, /^View Alt Text/)
    const alt = page.locator('.alt-text-dialog')
    await expect(alt.locator('textarea')).toHaveValue('Old words')
    await alt.locator('textarea').fill('A red square & its "edge"')
    await page.screenshot({ path: test.info().outputPath('inline-alt-text.png') })
    await alt.locator('button.primary').click()
    await expect(alt).toBeHidden()

    // the picture alone in its paragraph: 200 x 150 shown, made at 40 x 30
    await pictureItem(page, 1, /^Size and Position/)
    const size = page.locator('.size-position-dialog')
    const scaleH = size.locator('input[name="scaleHeight"]')
    const scaleW = size.locator('input[name="scaleWidth"]')
    const fromOriginal = size.locator('input[name="fromOriginal"]')
    await expect(fromOriginal).toBeChecked()
    await expect(scaleH).toHaveValue('500')
    await expect(scaleW).toHaveValue('500')
    // not from the original: the size the dialog opened with is 100%
    await fromOriginal.uncheck()
    await expect(scaleW).toHaveValue('100')
    await scaleW.fill('50')
    // the aspect ratio is locked: the height follows, and the absolute size with it
    await expect(scaleH).toHaveValue('50')
    await expect(size.getByLabel('Width', { exact: true })).toHaveValue(/^(1\.04"|2\.65 cm)$/)
    await fromOriginal.check()
    await expect(scaleW).toHaveValue('250')
    // unlocked, one side scales alone
    await size.locator('input[name="lock"]').uncheck()
    await scaleH.fill('300')
    await expect(scaleW).toHaveValue('250')
    // a preset place: against the right margin, at the bottom of the text area
    await expect(size.locator('input[name="place"][value="offset"]')).toBeChecked()
    await size.locator('input[name="place"][value="preset"]').check()
    await size.locator('select[name="presetH"]').selectOption('right')
    await size.locator('select[name="presetV"]').selectOption('bottom')
    await page.screenshot({ path: test.info().outputPath('size-position-scale.png') })
    await size.locator('button.primary').click()
    await expect(size).toBeHidden()
    await expect.poll(() => shown(1)).toEqual([100, 90])
    // the picture stands against the right edge of the text area
    await expect
      .poll(async () => {
        const [line] = await rects(page, '.doc-page[contenteditable="true"] p')
        const now = await rects(page, PICTURES)
        return Math.abs(now[1].left + now[1].width - (line.left + line.width))
      })
      .toBeLessThan(2)
    // and at the bottom of the text area, well below the text
    await expect
      .poll(async () => {
        const lines = await rects(page, '.doc-page[contenteditable="true"] p')
        const last = lines[lines.length - 1]
        return (await rects(page, PICTURES))[1].top - (last.top + last.height)
      })
      .toBeGreaterThan(200)
    placed = await onPaper(page)
    await page.screenshot({ path: test.info().outputPath('picture-at-preset.png') })

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const xml = await documentXml(docPath)
          return [
            /<wp:docPr id="1" name="Picture 1" descr="A red square &amp; its &quot;edge&quot;"\/>/.test(
              xml,
            ),
            (xml.match(/descr=/g) ?? []).length,
            /<wp:extent cx="952500" cy="857250"/.test(xml),
            /<wp:positionH relativeFrom="margin"><wp:align>right<\/wp:align>/.test(xml),
            /<wp:positionV relativeFrom="margin"><wp:align>bottom<\/wp:align>/.test(xml),
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([true, 1, true, true, true])
  } finally {
    await closeAndSaveVideo(launched, 'docs-floating-object-gaps')
  }

  launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-floating-object-gaps',
    openFile: docPath,
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await ready(page)
    await expect(page.locator(PICTURES)).toHaveCount(2, { timeout: 30_000 })
    await pictureItem(page, 0, /^View Alt Text/)
    const alt = page.locator('.alt-text-dialog')
    await expect(alt.locator('textarea')).toHaveValue('A red square & its "edge"')
    await alt.locator('button.primary').click()
    await expect(alt).toBeHidden()
    // the picture that was moved is the one alone in its paragraph: it reads back at its preset place and scale
    const alone = await page.evaluate(
      (selector) =>
        Array.from(document.querySelectorAll(selector)).findIndex((el) =>
          el.closest('.doc-protected'),
        ),
      PICTURES,
    )
    // on the page it is where it was before the save
    await expect
      .poll(async () => {
        const now = await onPaper(page)
        return Math.abs(now[0] - placed[0]) <= 2 && Math.abs(now[1] - placed[1]) <= 2
      })
      .toBe(true)
    await pictureItem(page, alone, /^Size and Position/)
    const size = page.locator('.size-position-dialog')
    await expect(size.locator('input[name="place"][value="preset"]')).toBeChecked()
    await expect(size.locator('select[name="presetH"]')).toHaveValue('right')
    await expect(size.locator('select[name="presetV"]')).toHaveValue('bottom')
    await expect(size.locator('input[name="scaleWidth"]')).toHaveValue('250')
    await expect(size.locator('input[name="scaleHeight"]')).toHaveValue('300')
  } finally {
    await closeAndSaveVideo(launched, 'docs-floating-object-gaps')
    rmSync(dir, { recursive: true, force: true })
  }
})

const OBJECTS = '.doc-page .doc-textbox, .doc-page .doc-img-wrap, .doc-page img.doc-inline-img'

/** the same objects by what a drag slides: a picture alone in its paragraph moves inside its slot until it is dropped */
const SHOWN =
  '.doc-page .doc-textbox, .doc-page .doc-img-wrap .doc-protected-img, .doc-page img.doc-inline-img'

/** two shapes, a picture alone in its paragraph and a picture beside text, none beside another */
const docxWithObjects = (): Promise<Buffer> =>
  docxWith(
    para('First') +
      shapeParagraph(1, 40, 10, 120, 60) +
      ['Second', 'Third', 'Fourth', 'Fifth'].map(para).join('') +
      shapeParagraph(2, 260, 10, 80, 90) +
      ['Sixth', 'Seventh', 'Eighth', 'Ninth', 'Tenth'].map(para).join('') +
      `<w:p>${pictureRun(3, 30, 10, 90, 70)}</w:p>` +
      ['Eleventh', 'Twelfth', 'Thirteenth', 'Fourteenth', 'Fifteenth'].map(para).join('') +
      `<w:p><w:r><w:t>Text beside a picture that floats in its paragraph. </w:t></w:r>${pictureRun(4, 300, 10, 140, 50)}</w:p>` +
      para('Last'),
  )

/** shows a ribbon tab; a click on the one already showing would fold the ribbon away */
async function showTab(page: Page, tab: RegExp) {
  const active = page.locator('.ribbon-tab.active', { hasText: tab })
  if ((await active.count()) === 0) await page.locator('.ribbon-tab', { hasText: tab }).click()
  await expect(active).toBeVisible()
}

/** runs one item of the Align menu the active contextual tab shows */
async function alignFromTab(page: Page, tab: RegExp, items: RegExp[]) {
  await showTab(page, tab)
  const button = page.locator('.ribbon-body .rb-align-objects')
  await expect(button).toBeEnabled()
  await button.click()
  for (const item of items) await page.locator('.layout-menu button', { hasText: item }).click()
}

const moved = (now: Rect[], from: Rect[]): number[][] =>
  now.map((r, i) => [Math.round(r.left - from[i].left), Math.round(r.top - from[i].top)])

test('Align is on the Shape Format and Picture Format tabs; a dragged group follows the pointer whole and Cut takes all of it', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-object-gaps-')))
  const docPath = join(dir, 'objects.docx')
  writeFileSync(docPath, await docxWithObjects())
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-floating-object-gaps',
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await ready(page)
    const objects = page.locator(OBJECTS)
    await expect(objects).toHaveCount(4, { timeout: 30_000 })
    const paper = async () => (await rects(page, '.doc-page[contenteditable="true"]'))[0]

    // Shape Format: two shapes line up on the left edge of the leftmost
    await objects.nth(0).click()
    await objects.nth(1).click({ modifiers: ['Shift'] })
    await expect(page.locator('.doc-multi-selected')).toHaveCount(1)
    const before = await rects(page, OBJECTS)
    await showTab(page, /^Shape Format$/)
    await page.locator('.ribbon-body .rb-align-objects').click()
    await page.screenshot({ path: test.info().outputPath('shape-format-align.png') })
    await page.locator('.layout-menu button', { hasText: /^Align Left$/ }).click()
    await expect
      .poll(async () => {
        const now = await rects(page, OBJECTS)
        return [Math.round(now[0].left - before[0].left), Math.round(now[1].left - before[0].left)]
      })
      .toEqual([0, 0])

    // Picture Format, a picture alone in its paragraph: one object lines up against the page
    await objects.nth(2).click()
    await alignFromTab(page, /^Picture Format$/, [/^Align Center$/])
    await expect
      .poll(async () => {
        const now = (await rects(page, OBJECTS))[2]
        const sheet = await paper()
        return Math.abs(now.left + now.width / 2 - (sheet.left + sheet.width / 2))
      })
      .toBeLessThan(2)

    // Picture Format, a picture beside text: against the margin once that is what Align lines up with
    await objects.nth(3).click()
    await showTab(page, /^Picture Format$/)
    await page.locator('.ribbon-body .rb-align-objects').click()
    await page.locator('.layout-menu button', { hasText: /^Align to Margin$/ }).click()
    await expect(
      page.locator('.layout-menu button', { hasText: /^Align to Margin$/ }),
    ).toHaveAttribute('aria-checked', 'true')
    await page.screenshot({ path: test.info().outputPath('picture-format-align.png') })
    await page.locator('.layout-menu button', { hasText: /^Align Left$/ }).click()
    await expect
      .poll(async () => {
        const now = (await rects(page, OBJECTS))[3]
        const [line] = await rects(page, '.doc-page[contenteditable="true"] p')
        return Math.abs(now.left - line.left)
      })
      .toBeLessThan(2)

    // the four as one group
    await objects.nth(0).click()
    for (const i of [1, 2, 3]) await objects.nth(i).click({ modifiers: ['Shift'] })
    await expect(page.locator('.doc-multi-selected')).toHaveCount(3)
    await showTab(page, /^Layout$/)
    await page.locator('.rb-big', { hasText: /^Group$/ }).click()
    await page.locator('.layout-menu button', { hasText: /^Group$/ }).click()
    await page.locator('.doc-page p', { hasText: 'Last' }).click()
    await objects.nth(1).click()
    await expect(page.locator('.doc-multi-selected')).toHaveCount(3)

    // while the button is down every member is under the pointer's move, not only the one pressed
    const start = await rects(page, OBJECTS)
    const shownAt = await rects(page, SHOWN)
    const grab = start[2]
    await page.mouse.move(grab.left + grab.width / 2, grab.top + grab.height / 2)
    await page.mouse.down()
    await page.mouse.move(grab.left + grab.width / 2 + 60, grab.top + grab.height / 2 + 40, {
      steps: 6,
    })
    await expect
      .poll(async () =>
        moved(await rects(page, SHOWN), shownAt).every(
          ([dx, dy]) => Math.abs(dx - 60) <= 2 && Math.abs(dy - 40) <= 2,
        ),
      )
      .toBe(true)
    await page.screenshot({ path: test.info().outputPath('group-mid-drag.png') })
    await page.mouse.up()
    await expect
      .poll(async () => {
        const by = moved(await rects(page, OBJECTS), start)
        return (
          Math.abs(by[2][0]) > 20 &&
          by.every((m) => Math.abs(m[0] - by[2][0]) <= 2 && Math.abs(m[1] - by[2][1]) <= 2)
        )
      })
      .toBe(true)
    await expect(page.locator('.doc-multi-selected')).toHaveCount(3)

    // Cut takes the whole group out of the document, and out of the saved file
    await objects.nth(0).click({ button: 'right' })
    await page.locator('.ctx-menu .ctx-item', { hasText: /^Cut/ }).click()
    await expect(objects).toHaveCount(0)
    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const xml = await documentXml(docPath)
          return (xml.match(/<w:drawing>/g) ?? []).length
        },
        { timeout: 30_000 },
      )
      .toBe(0)
    // the clipboard holds more than the object that was pressed (a shape): the group's pictures come back
    await page.locator('.doc-page p', { hasText: 'Last' }).click({ button: 'right' })
    await page
      .locator('.ctx-menu .ctx-item', { hasText: /^Paste/ })
      .first()
      .click()
    await expect
      .poll(() => page.locator('.doc-page[contenteditable="true"] img').count())
      .toBeGreaterThanOrEqual(2)
  } finally {
    await closeAndSaveVideo(launched, 'docs-floating-object-gaps')
    rmSync(dir, { recursive: true, force: true })
  }
})
