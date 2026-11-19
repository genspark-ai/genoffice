/**
 * Several floating objects selected together (Shift+click) and the Layout
 * tab's Align menu: the objects line up on screen through the real press and
 * ribbon path, and the positions are what the saved file carries.
 */
import { test, expect, type Page } from '@playwright/test'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import {
  boxRects,
  docxWith,
  para,
  pictureRun,
  rects,
  saveFromMenu,
  shapeParagraph,
} from './docs-objects-fixture'

const docxWithShapes = (): Promise<Buffer> =>
  docxWith(
    para('First') +
      shapeParagraph(1, 40, 30, 120, 60) +
      para('Second') +
      shapeParagraph(2, 260, 120, 80, 90) +
      para('Third') +
      shapeParagraph(3, 150, 300, 160, 40) +
      para('Last'),
  )

/**
 * A picture alone in its paragraph, and one sharing a paragraph with text, far enough apart
 * that neither is beside the other: side floats cannot overlap on the canvas.
 */
const docxWithPictures = (): Promise<Buffer> =>
  docxWith(
    para('First') +
      `<w:p>${pictureRun(1, 30, 20, 90, 70)}</w:p>` +
      ['Second', 'Third', 'Fourth', 'Fifth', 'Sixth'].map(para).join('') +
      `<w:p><w:r><w:t>Text beside a picture that floats in its paragraph. </w:t></w:r>${pictureRun(2, 300, 10, 140, 50)}</w:p>` +
      para('Last'),
  )

async function align(page: Page, item: string): Promise<void> {
  await page.locator('.rb-big', { hasText: /^Align$/ }).click()
  await page.locator('.layout-menu button', { hasText: new RegExp(`^${item}$`) }).click()
}

test('Shift+click selects several shapes and Align lines them up, on screen and in the saved file', async () => {
  test.setTimeout(120_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-object-align-')))
  const docPath = join(dir, 'shapes.docx')
  writeFileSync(docPath, await docxWithShapes())
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-object-align',
    openFile: docPath,
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    const boxes = page.locator('.doc-page .doc-textbox')
    await expect(boxes).toHaveCount(3, { timeout: 30_000 })
    const layoutTab = page.locator('.ribbon-tab', { hasText: /^Layout$/ })
    await layoutTab.click()
    const alignButton = page.locator('.rb-big', { hasText: /^Align$/ })
    // nothing floating is selected yet: the button opens, as in Word, on greyed align items
    await alignButton.click()
    await expect(page.locator('.layout-menu button', { hasText: /^Align Left$/ })).toBeDisabled()
    await page.keyboard.press('Escape')

    const start = await boxRects(page)
    await boxes.nth(0).click()
    await boxes.nth(1).click({ modifiers: ['Shift'] })
    await boxes.nth(2).click({ modifiers: ['Shift'] })
    await expect(page.locator('.doc-multi-selected')).toHaveCount(2)
    // selecting a shape shows its own tab; Align sits on Layout
    await layoutTab.click()
    await expect(alignButton).toBeEnabled()
    await page.screenshot({ path: test.info().outputPath('selected.png') })

    // against each other: the left edge of the leftmost, every top as it was
    await alignButton.click()
    await page.screenshot({ path: test.info().outputPath('align-menu.png') })
    await page.locator('.layout-menu button', { hasText: /^Align Left$/ }).click()
    await expect
      .poll(async () => (await boxRects(page)).map((r) => Math.round(r.left)))
      .toEqual(start.map(() => Math.round(start[0].left)))
    const lefted = await boxRects(page)
    lefted.forEach((r, i) => expect(Math.abs(r.top - start[i].top)).toBeLessThan(1.5))
    await expect(page.locator('.doc-multi-selected')).toHaveCount(2)

    // equal gaps between the three, the outer two staying put
    await align(page, 'Distribute Vertically')
    await expect
      .poll(async () => {
        const [a, b, c] = await boxRects(page)
        return Math.round(b.top - (a.top + a.height) - (c.top - (b.top + b.height)))
      })
      .toBe(0)
    const spread = await boxRects(page)
    expect(Math.abs(spread[0].top - lefted[0].top)).toBeLessThan(1.5)
    expect(Math.abs(spread[2].top - lefted[2].top)).toBeLessThan(1.5)
    await page.screenshot({ path: test.info().outputPath('distributed.png') })

    // against the page: each one centred on the paper
    await alignButton.click()
    await page.locator('.layout-menu button', { hasText: /^Align to Page$/ }).click()
    await page.locator('.layout-menu button', { hasText: /^Align Center$/ }).click()
    const paper = await page.evaluate(() => {
      const r = document.querySelector('.doc-page')!.getBoundingClientRect()
      return r.left + r.width / 2
    })
    await expect
      .poll(async () => (await boxRects(page)).map((r) => Math.round(r.left + r.width / 2 - paper)))
      .toEqual([0, 0, 0])
    await page.screenshot({ path: test.info().outputPath('centered.png') })

    // one undo takes the whole alignment back
    await app.evaluate(({ Menu }) => {
      const edit = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Edit')
      edit?.submenu?.items.find((candidate) => candidate.label.startsWith('Undo'))?.click()
    })
    await expect
      .poll(async () => (await boxRects(page)).map((r) => Math.round(r.left)))
      .toEqual(spread.map((r) => Math.round(r.left)))

    // the saved anchors carry the same left offset
    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const zip = await JSZip.loadAsync(readFileSync(docPath))
          const xml = (await zip.file('word/document.xml')?.async('string')) ?? ''
          const offsets = Array.from(
            xml.matchAll(/<wp:positionH[^>]*>\s*<wp:posOffset>(-?\d+)<\/wp:posOffset>/g),
          ).map((m) => Number(m[1]))
          return offsets.length === 3 && new Set(offsets).size
        },
        { timeout: 30_000 },
      )
      .toBe(1)
  } finally {
    await closeAndSaveVideo(launched, 'docs-object-align')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a picture alone in its paragraph and one beside text line up together and against the margin', async () => {
  test.setTimeout(120_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-object-align-')))
  const docPath = join(dir, 'pictures.docx')
  writeFileSync(docPath, await docxWithPictures())
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-object-align',
    openFile: docPath,
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    const PICTURES = '.doc-page .doc-img-wrap, .doc-page img.doc-inline-img'
    const pictures = page.locator(PICTURES)
    await expect(pictures).toHaveCount(2, { timeout: 30_000 })
    const start = await rects(page, PICTURES)
    await pictures.nth(0).click()
    await pictures.nth(1).click({ modifiers: ['Shift'] })
    await expect(page.locator('.doc-multi-selected')).toHaveCount(1)
    await page.locator('.ribbon-tab', { hasText: /^Layout$/ }).click()

    // against the margin: both right edges on the text area's right edge
    const alignButton = page.locator('.rb-big', { hasText: /^Align$/ })
    await alignButton.click()
    await page.locator('.layout-menu button', { hasText: /^Align to Margin$/ }).click()
    await page.locator('.layout-menu button', { hasText: /^Align Right$/ }).click()
    const margin = await page.evaluate(() => {
      const paper = document.querySelector('.doc-page')!
      const r = paper.getBoundingClientRect()
      const style = getComputedStyle(paper)
      return {
        left: r.left + parseFloat(style.paddingLeft),
        right: r.right - parseFloat(style.paddingRight),
      }
    })
    await expect
      .poll(async () =>
        (await rects(page, PICTURES)).map((r) => Math.round(r.left + r.width - margin.right)),
      )
      .toEqual([0, 0])
    await page.screenshot({ path: test.info().outputPath('pictures-right.png') })

    // and back to its left edge
    await align(page, 'Align Left')
    await expect
      .poll(async () => (await rects(page, PICTURES)).map((r) => Math.round(r.left - margin.left)))
      .toEqual([0, 0])

    // each alignment is one step to undo: two take the pictures back where they started
    const undo = () =>
      app.evaluate(({ Menu }) => {
        const edit = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Edit')
        edit?.submenu?.items.find((candidate) => candidate.label.startsWith('Undo'))?.click()
      })
    await undo()
    await undo()
    await expect
      .poll(async () =>
        (await rects(page, PICTURES)).map((r, i) => Math.abs(r.left - start[i].left) < 1.5),
      )
      .toEqual([true, true])

    // against each other: the higher picture's top
    await alignButton.click()
    await page.locator('.layout-menu button', { hasText: /^Align Selected Objects$/ }).click()
    await page.locator('.layout-menu button', { hasText: /^Align Top$/ }).click()
    await expect
      .poll(async () => {
        const [a, b] = await rects(page, PICTURES)
        return Math.abs(a.top - b.top) < 1
      })
      .toBe(true)
    const topped = await rects(page, PICTURES)
    expect(Math.abs(topped[0].top - Math.min(start[0].top, start[1].top))).toBeLessThan(1.5)
    await page.screenshot({ path: test.info().outputPath('pictures-top.png') })

    // Shift held on a grip of the selected picture still resizes it: the press is not a selection
    const before = (await rects(page, PICTURES))[0]
    const grip = page.locator('.img-resize-handle').last()
    const at = (await grip.boundingBox())!
    await page.keyboard.down('Shift')
    await page.mouse.move(at.x + at.width / 2, at.y + at.height / 2)
    await page.mouse.down()
    await page.mouse.move(at.x + at.width / 2 + 24, at.y + at.height / 2 + 24, { steps: 4 })
    await page.mouse.up()
    await page.keyboard.up('Shift')
    await expect
      .poll(async () => Math.abs((await rects(page, PICTURES))[0].width - before.width))
      .toBeGreaterThan(5)
    await expect(page.locator('.doc-multi-selected')).toHaveCount(1)

    // both pictures are still in the saved file, anchored
    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const zip = await JSZip.loadAsync(readFileSync(docPath))
          const xml = (await zip.file('word/document.xml')?.async('string')) ?? ''
          return [xml.match(/<wp:anchor /g)?.length ?? 0, xml.includes('<wp:inline')]
        },
        { timeout: 30_000 },
      )
      .toEqual([2, false])
  } finally {
    await closeAndSaveVideo(launched, 'docs-object-align')
    rmSync(dir, { recursive: true, force: true })
  }
})
