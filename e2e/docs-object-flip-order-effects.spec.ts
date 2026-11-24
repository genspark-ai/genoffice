/**
 * Arrange on shapes and floating pictures, through the real ribbon: a
 * mirrored shape is drawn mirrored and Flip turns it over in the file; a
 * floating picture among text is stacked by its rank and restacked; the text
 * fill takes a transparency; Convert Table to Text keeps the shapes of the
 * cells; a group placed by alignment comes apart where it was drawn.
 */
import { test, expect, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { ElectronApplication } from 'playwright'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import {
  docxWith,
  groupParagraph,
  groupShape,
  para,
  rects,
  saveFromMenu,
  shapeParagraph,
} from './docs-objects-fixture'

const VIDEO = 'docs-object-flip-order-effects'
const emu = (px: number) => px * 9525
const BOXES = '.doc-page .doc-textbox'
const PICTURES = '.doc-page[contenteditable="true"] img.doc-inline-img'

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

const control = (page: Page, tip: string) =>
  page.locator(`[data-ribbon-body] button[data-tip="${tip}"]`)

async function pick(page: Page, tip: string, item: string) {
  await control(page, tip).click()
  await page
    .locator('[data-ribbon-body] [data-rb-panel] button', { hasText: new RegExp(`^${item}$`) })
    .click()
  await expect(page.locator('[data-ribbon-body] [data-rb-panel]')).toHaveCount(0)
}

const clickMenu = (app: ElectronApplication, path: string[]) =>
  app.evaluate(({ Menu }, labels) => {
    let items = Menu.getApplicationMenu()?.items ?? []
    for (const [i, label] of labels.entries()) {
      const item = items.find((candidate) => candidate.label === label)
      if (i === labels.length - 1) item?.click()
      else items = item?.submenu?.items ?? []
    }
  }, path)

/** the SVG a box is drawn with, and the transform on the box itself */
const drawn = (page: Page, nth: number) =>
  page.evaluate(
    ([selector, index]) => {
      const el = document.querySelectorAll<HTMLElement>(selector as string)[index as number]
      const svg = decodeURIComponent(
        /url\("data:image\/svg\+xml,([^"]+)"\)/.exec(el.style.backgroundImage)?.[1] ?? '',
      )
      return {
        mirror: /<g transform="[^"]*scale\(([^)]*)\)/.exec(svg)?.[1] ?? null,
        own: el.style.transform,
      }
    },
    [BOXES, nth] as const,
  )

/** every box as [left, top, width, height] from the paper's corner */
async function onPaper(page: Page, selector = BOXES): Promise<number[][]> {
  const [paper] = await rects(page, '.doc-page')
  return (await rects(page, selector)).map((r) => [
    r.left - paper.left,
    r.top - paper.top,
    r.width,
    r.height,
  ])
}

const near = (a: number[][], b: number[][], by: number): boolean =>
  a.length === b.length && a.every((box, i) => box.every((v, k) => Math.abs(v - b[i][k]) <= by))

/** a shape that is not its own mirror image; a pentagon arrow leaves its word room, a triangle has none */
const lopsided = (id: number, y: number, attrs: string, text = '') =>
  shapeParagraph(id, 300, y, 140, 70)
    .replace('<a:xfrm>', `<a:xfrm${attrs}>`)
    .replace('prst="rect"', `prst="${text ? 'homePlate' : 'rtTriangle'}"`)
    .replace(
      '<wps:bodyPr/>',
      text
        ? `<wps:txbx><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/>`
        : '<wps:bodyPr/>',
    )

test('a shape the file mirrors is drawn mirrored, and Flip turns a shape over on the page and in the file', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-flip-')))
  const docPath = join(dir, 'flip.docx')
  writeFileSync(
    docPath,
    await docxWith(
      para('One') +
        lopsided(1, 10, ' flipH="1"') +
        ['Two', 'Three', 'Four', 'Five', 'Six'].map(para).join('') +
        lopsided(2, 10, '', 'Label') +
        ['Seven', 'Eight', 'Nine', 'Ten', 'Eleven'].map(para).join(''),
    ),
  )
  let launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO, openFile: docPath })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await ready(page)
    const boxes = page.locator(BOXES)
    await expect(boxes).toHaveCount(2, { timeout: 30_000 })
    // the file's flipH: the outline through a mirror; the second shape as it is
    expect(await drawn(page, 0)).toEqual({ mirror: '-1 1', own: '' })
    expect(await drawn(page, 1)).toEqual({ mirror: null, own: '' })
    const label = page.locator(`${BOXES} >> nth=1`).locator('p, .doc-textbox-para').first()
    const wordAt = async () => {
      const [paper] = await rects(page, '.doc-page')
      const r = await label.boundingBox()
      return r ? [Math.round(r.x - paper.left), Math.round(r.y - paper.top)] : []
    }
    const labelWas = await wordAt()
    const placed = await onPaper(page)
    await page.screenshot({ path: test.info().outputPath('opened.png') })

    await boxes.nth(1).click({ position: { x: 4, y: 60 } })
    await expect(page.locator('.ribbon-tab.active')).toHaveText('Shape Format')
    await control(page, 'Rotate').click()
    const panel = page.locator('[data-ribbon-body] [data-rb-panel]')
    await expect(panel.locator('button', { hasText: /^Flip Horizontal$/ })).toBeEnabled()
    await expect(panel.locator('button', { hasText: /^Flip Vertical$/ })).toBeEnabled()
    await page.keyboard.press('Escape')
    await pick(page, 'Rotate', 'Flip Horizontal')
    await expect.poll(async () => (await drawn(page, 1)).mirror).toBe('-1 1')
    await pick(page, 'Rotate', 'Flip Vertical')
    await expect.poll(async () => (await drawn(page, 1)).mirror).toBe('-1 -1')
    // the outline turns over; the text and the boxes stay where they were
    expect(await wordAt()).toEqual(labelWas)
    const now = await onPaper(page)
    expect(now, JSON.stringify([placed, now])).toHaveLength(placed.length)
    expect(near(now, placed, 0.5), JSON.stringify([placed, now])).toBe(true)
    await page.screenshot({ path: test.info().outputPath('flipped.png') })

    // the shape the file mirrored, mirrored back
    await boxes.nth(0).click({ position: { x: 130, y: 60 } })
    await pick(page, 'Rotate', 'Flip Horizontal')
    await expect.poll(async () => (await drawn(page, 0)).mirror).toBe(null)

    await saveFromMenu(app)
    await expect
      .poll(async () => (await documentXml(docPath)).includes('flipV="1"'), { timeout: 30_000 })
      .toBe(true)
    const xml = await documentXml(docPath)
    expect(xml.match(/<a:xfrm[^>]*>/g)).toEqual(['<a:xfrm>', '<a:xfrm flipH="1" flipV="1">'])
  } finally {
    await closeAndSaveVideo(launched, VIDEO)
  }

  launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO, openFile: docPath })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(page.locator(BOXES)).toHaveCount(2, { timeout: 30_000 })
    expect(await drawn(page, 0)).toEqual({ mirror: null, own: '' })
    expect(await drawn(page, 1)).toEqual({ mirror: '-1 -1', own: '' })
    await page.screenshot({ path: test.info().outputPath('reopened.png') })
  } finally {
    await closeAndSaveVideo(launched, VIDEO)
    rmSync(dir, { recursive: true, force: true })
  }
})

const floatingPicture = (id: number, rank: number, x: number, y: number) =>
  `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" ` +
  `relativeHeight="${251658240 + rank}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
  `<wp:simplePos x="0" y="0"/>` +
  `<wp:positionH relativeFrom="column"><wp:posOffset>${emu(x)}</wp:posOffset></wp:positionH>` +
  `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${emu(y)}</wp:posOffset></wp:positionV>` +
  `<wp:extent cx="${emu(120)}" cy="${emu(90)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
  `<wp:wrapNone/><wp:docPr id="${id}" name="Picture ${id}"/><wp:cNvGraphicFramePr/>` +
  `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
  `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
  `<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(120)}" cy="${emu(90)}"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
  `</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`

const stacking = (page: Page): Promise<number[]> =>
  page.evaluate(
    (selector) =>
      Array.from(document.querySelectorAll<HTMLElement>(selector), (img) =>
        Number(getComputedStyle(img).zIndex),
      ),
    PICTURES,
  )

const ranksIn = (xml: string): number[] =>
  Array.from(xml.matchAll(/relativeHeight="(\d+)"/g), (m) => Number(m[1]) - 251658240)

test('floating pictures among text are stacked by their rank, and Bring Forward / Send Backward restack them', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-rank-')))
  const docPath = join(dir, 'rank.docx')
  // the first in the text has the higher rank: the file says it lies over the second
  writeFileSync(
    docPath,
    await docxWith(
      para('One') +
        `<w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r>${floatingPicture(1, 3072, 200, 10)}` +
        `${floatingPicture(2, 1024, 260, 50)}<w:r><w:t xml:space="preserve"> after.</w:t></w:r></w:p>` +
        ['Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight'].map(para).join(''),
    ),
  )
  let launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO, openFile: docPath })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await ready(page)
    const pictures = page.locator(PICTURES)
    await expect(pictures).toHaveCount(2, { timeout: 30_000 })
    const [first, second] = await stacking(page)
    expect(first).toBeGreaterThan(second)
    const placed = await onPaper(page, PICTURES)
    await page.screenshot({ path: test.info().outputPath('opened.png') })

    // the second picture, pressed where the first does not cover it
    await pictures.nth(1).click({ position: { x: 100, y: 80 } })
    await expect(page.locator('.ribbon-tab.active')).toHaveText('Picture Format')
    await expect(control(page, 'Bring Forward')).toBeEnabled()
    await expect(control(page, 'Send Backward')).toBeEnabled()
    await pick(page, 'Bring Forward', 'Bring Forward')
    await expect
      .poll(async () => {
        const [a, b] = await stacking(page)
        return b > a
      })
      .toBe(true)
    expect(near(await onPaper(page, PICTURES), placed, 0.5)).toBe(true)
    await page.screenshot({ path: test.info().outputPath('brought-forward.png') })

    await saveFromMenu(app)
    await expect
      .poll(async () => ranksIn(await documentXml(docPath)), { timeout: 30_000 })
      .toEqual([3072, 3073])

    // the save read the file back: the picture is pressed again
    await pictures.nth(1).click({ position: { x: 100, y: 80 } })
    await pick(page, 'Send Backward', 'Send to Back')
    await expect
      .poll(async () => {
        const [a, b] = await stacking(page)
        return a > b
      })
      .toBe(true)
    await saveFromMenu(app)
    await expect
      .poll(async () => ranksIn(await documentXml(docPath)), { timeout: 30_000 })
      .toEqual([3072, 3071])
    await pictures.nth(1).click({ position: { x: 100, y: 80 } })
    await pick(page, 'Bring Forward', 'Bring to Front')
    await saveFromMenu(app)
    await expect
      .poll(async () => ranksIn(await documentXml(docPath)), { timeout: 30_000 })
      .toEqual([3072, 3073])
  } finally {
    await closeAndSaveVideo(launched, VIDEO)
  }

  launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO, openFile: docPath })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(page.locator(PICTURES)).toHaveCount(2, { timeout: 30_000 })
    const [first, second] = await stacking(page)
    expect(second).toBeGreaterThan(first)
  } finally {
    await closeAndSaveVideo(launched, VIDEO)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Text Effects sets how see-through the text fill is, and the file says so', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-fill-alpha-')))
  const docPath = join(dir, 'fill.docx')
  writeFileSync(docPath, await docxWith(para('Plain fancy words') + para('Last')))
  const open = async (app: ElectronApplication, page: Page) => {
    await page.evaluate(() => {
      const { editor } = (
        window as never as {
          __aidocs: {
            editor: { chain(): { focus(): { setTextSelection(r: object): { run(): boolean } } } }
          }
        }
      ).__aidocs
      editor.chain().focus().setTextSelection({ from: 7, to: 12 }).run()
    })
    await clickMenu(app, ['Format', 'Text Effects…'])
    const dialog = page.locator('.text-fx-dialog')
    await expect(dialog).toBeVisible()
    return dialog
  }
  const fillOf = (page: Page) =>
    page.evaluate(() => {
      const span = document.querySelector<HTMLElement>('.doc-page p span[style*="text-fill-color"]')
      return span
        ? [span.textContent, span.style.getPropertyValue('-webkit-text-fill-color')]
        : null
    })
  let launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO, openFile: docPath })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'Plain fancy words' })).toBeVisible({
      timeout: 30_000,
    })
    let dialog = await open(app, page)
    const percent = dialog.locator('input[name="fillTransparency"]')
    await expect(percent).toHaveValue('0')
    await dialog.locator('[data-fx="fill"] .text-fx-swatch[aria-label="#4472C4"]').click()
    await percent.fill('40')
    await expect(dialog.locator('[data-fx="fill"] input[type="range"]')).toHaveValue('40')
    await page.screenshot({ path: test.info().outputPath('text-fill-transparency.png') })
    await dialog.locator('button.primary').click()
    await expect.poll(() => fillOf(page)).toEqual(['fancy', 'rgba(68, 114, 196, 0.6)'])

    await saveFromMenu(app)
    await expect
      .poll(async () => (await documentXml(docPath)).includes('<w14:textFill>'), {
        timeout: 30_000,
      })
      .toBe(true)
    // w14:alpha of a fill is how transparent it is
    expect(await documentXml(docPath)).toContain(
      '<w14:textFill><w14:solidFill><w14:srgbClr w14:val="4472C4"><w14:alpha w14:val="40000"/></w14:srgbClr></w14:solidFill></w14:textFill>',
    )

    // the dialog opens on what the text has; a gradient has no transparency of its own here
    dialog = await open(app, page)
    await expect(dialog.locator('input[name="fillTransparency"]')).toHaveValue('40')
    await dialog.locator('[data-fx="fill"] .text-fx-choice', { hasText: /^Linear Down$/ }).click()
    await expect(dialog.locator('input[name="fillTransparency"]')).toHaveCount(0)
    await dialog.locator('[data-fx="fill"] .text-fx-choice', { hasText: /^None$/ }).click()
    await dialog.locator('input[name="fillTransparency"]').fill('0')
    await dialog.locator('button.primary').click()
    await expect.poll(() => fillOf(page)).toBeNull()
  } finally {
    await closeAndSaveVideo(launched, VIDEO)
  }

  launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO, openFile: docPath })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'Plain fancy words' })).toBeVisible({
      timeout: 30_000,
    })
    // the save before the second visit to the dialog is what the file holds
    await expect.poll(() => fillOf(page)).toEqual(['fancy', 'rgba(68, 114, 196, 0.6)'])
  } finally {
    await closeAndSaveVideo(launched, VIDEO)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Convert Table to Text takes the shapes anchored in the cells along with the text', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-table-shapes-')))
  const docPath = join(dir, 'table.docx')
  const drawing = /<w:r><w:drawing>[\s\S]*<\/w:drawing><\/w:r>/.exec(
    shapeParagraph(7, 60, 4, 50, 30),
  )![0]
  const cell = (inner: string) =>
    `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`
  const text = (t: string, extra = '') => `<w:p><w:r><w:t>${t}</w:t></w:r>${extra}</w:p>`
  writeFileSync(
    docPath,
    await docxWith(
      para('One') +
        '<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
        `<w:tr>${cell(text('A1', drawing))}${cell(text('B1'))}</w:tr>` +
        `<w:tr>${cell(text('A2'))}${cell(text('B2'))}</w:tr></w:tbl>` +
        ['Two', 'Three', 'Four'].map(para).join(''),
    ),
  )
  let launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO, openFile: docPath })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await ready(page)
    await expect(page.locator(BOXES)).toHaveCount(1, { timeout: 30_000 })
    await page.screenshot({ path: test.info().outputPath('table.png') })
    await page.locator('.doc-page table td').nth(1).locator('p').first().click()
    await clickMenu(app, ['Table', 'Convert', 'Convert Table to Text…'])
    const dialog = page.locator('.convert-table-dialog')
    await dialog.locator('input[name="separator"][value="tab"]').check()
    await dialog.locator('button.primary').click()
    await expect(page.locator('.doc-page table')).toHaveCount(0)
    // the shape is still on the page, by the line its cell's text went to
    await expect(page.locator(BOXES)).toHaveCount(1)
    const line = await page.locator('.doc-page > p', { hasText: 'A1' }).boundingBox()
    const box = await page.locator(BOXES).boundingBox()
    expect(Math.abs(box!.y - (line!.y + 4))).toBeLessThan(3)
    await page.screenshot({ path: test.info().outputPath('converted.png') })

    await saveFromMenu(app)
    await expect
      .poll(async () => !(await documentXml(docPath)).includes('<w:tbl>'), { timeout: 30_000 })
      .toBe(true)
    const xml = await documentXml(docPath)
    const host =
      xml.match(/<w:p[ >](?:(?!<\/w:p>)[\s\S])*name="Rectangle 7"[\s\S]*?<\/w:p>/)?.[0] ?? ''
    expect(host).toContain('>A1</w:t>')
    expect(host).toContain('>B1</w:t>')
    expect(xml.match(/name="Rectangle 7"/g)).toHaveLength(1)
  } finally {
    await closeAndSaveVideo(launched, VIDEO)
  }

  launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO, openFile: docPath })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(page.locator(BOXES)).toHaveCount(1, { timeout: 30_000 })
    await expect(page.locator('.doc-page table')).toHaveCount(0)
    const line = await page.locator('.doc-page').getByText('A1').first().boundingBox()
    const box = await page.locator(BOXES).boundingBox()
    expect(Math.abs(box!.y - (line!.y + 4))).toBeLessThan(3)
  } finally {
    await closeAndSaveVideo(launched, VIDEO)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a group placed by alignment comes apart where the page drew it, and the file puts it there again', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-aligned-group-')))
  const docPath = join(dir, 'aligned.docx')
  const children =
    groupShape(1, [0, 0, 100, 60], 'C00000', { text: 'Box' }) +
    groupShape(2, [140, 20, 80, 80], '00B050', { prst: 'ellipse' }) +
    groupShape(3, [60, 110, 120, 40], '4472C4', { prst: 'triangle', attrs: ' rot="1800000"' })
  const centred = groupParagraph(60, [60, 20, 220, 150], children).replace(
    /<wp:positionH[^>]*>.*?<\/wp:positionH>/,
    '<wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>',
  )
  const right = groupParagraph(70, [60, 20, 220, 150], children).replace(
    /<wp:positionH[^>]*>.*?<\/wp:positionH>/,
    '<wp:positionH relativeFrom="page"><wp:align>right</wp:align></wp:positionH>',
  )
  writeFileSync(
    docPath,
    await docxWith(
      ['One', 'Two'].map(para).join('') +
        centred +
        ['Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten'].map(para).join('') +
        right +
        ['Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen']
          .map(para)
          .join(''),
    ),
  )
  const NODES = '.doc-page .doc-protected:has(.doc-textbox)'
  const ungroup = async (page: Page) => {
    const group = page.locator('.rb-big', { hasText: /^Group$/ })
    if (!(await group.isVisible()))
      await page.locator('.ribbon-tab', { hasText: /^Layout$/ }).click()
    await group.click()
    const item = page.locator('.layout-menu button', { hasText: /^Ungroup$/ })
    await expect(item).toBeEnabled()
    await item.click()
  }
  const textTops = (page: Page): Promise<number[]> =>
    page.evaluate(() => {
      const paper = document.querySelector('.doc-page')!.getBoundingClientRect().top
      return Array.from(document.querySelectorAll('.doc-page > p'), (p) =>
        p.textContent ? Math.round(p.getBoundingClientRect().top - paper) : null,
      ).filter((top): top is number => top !== null)
    })
  let placed: number[][] = []
  let launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO, openFile: docPath })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await ready(page)
    const boxes = page.locator(BOXES)
    await expect(boxes).toHaveCount(6, { timeout: 30_000 })
    await expect(page.locator(NODES)).toHaveCount(2)
    placed = await onPaper(page)
    const lines = await textTops(page)
    await page.screenshot({ path: test.info().outputPath('grouped.png') })

    await boxes.nth(1).click()
    await ungroup(page)
    await expect(page.locator(NODES)).toHaveCount(4)
    await boxes.nth(4).scrollIntoViewIfNeeded()
    await boxes.nth(4).click()
    await ungroup(page)
    await expect(page.locator(NODES)).toHaveCount(6)
    // nothing moved: neither the children nor the text
    await expect.poll(async () => near(await onPaper(page), placed, 1.5)).toBe(true)
    expect(await textTops(page)).toEqual(lines)
    await page.screenshot({ path: test.info().outputPath('ungrouped.png') })

    await saveFromMenu(app)
    await expect
      .poll(async () => !(await documentXml(docPath)).includes('<wpg:wgp'), { timeout: 30_000 })
      .toBe(true)
    const xml = await documentXml(docPath)
    expect(xml).not.toContain('<wp:align>')
    expect(xml.match(/<wp:positionH relativeFrom="column"><wp:posOffset>/g)).toHaveLength(6)
    writeFileSync(test.info().outputPath('saved.docx'), readFileSync(docPath))
  } finally {
    await closeAndSaveVideo(launched, VIDEO)
  }

  launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO, openFile: docPath })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(page.locator(BOXES)).toHaveCount(6, { timeout: 30_000 })
    await expect.poll(async () => near(await onPaper(page), placed, 1.5)).toBe(true)
    await page.screenshot({ path: test.info().outputPath('reopened.png') })
  } finally {
    await closeAndSaveVideo(launched, VIDEO)
    rmSync(dir, { recursive: true, force: true })
  }
})
