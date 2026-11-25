/**
 * Change Picture on the right-click menu takes its picture from the icons or
 * from the clipboard as well as from a file, and the picture keeps its width
 * and its wrapping. A picture that shares its paragraph with a shape turns
 * with Rotate, and Ungroup is offered on a group placed by alignment too.
 * Each is checked on the page and in the saved file.
 */
import { test, expect, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import {
  docxWith,
  groupParagraph,
  groupShape,
  para,
  pictureRun,
  rects,
  saveFromMenu,
  shapeParagraph,
} from './docs-objects-fixture'

const PICTURE = '.doc-page[contenteditable="true"] .doc-protected img'

const ready = (page: Page) =>
  page.waitForFunction(
    () => Boolean((window as unknown as { __aidocs?: { editor?: unknown } }).__aidocs?.editor),
    undefined,
    { timeout: 30_000 },
  )

const partsOf = async (file: string) => {
  const zip = await JSZip.loadAsync(readFileSync(file))
  return {
    xml: (await zip.file('word/document.xml')?.async('string')) ?? '',
    media: Object.keys(zip.files).filter((name) => name.startsWith('word/media/')),
  }
}

const size = async (page: Page, selector: string): Promise<number[]> => {
  const [r] = await rects(page, selector)
  return r ? [Math.round(r.width), Math.round(r.height)] : []
}

/** opens the picture's right-click menu on Change Picture */
async function changePicture(page: Page) {
  await page.locator(PICTURE).first().click({ button: 'right' })
  const menu = page.locator('.ctx-menu')
  await menu.waitFor({ timeout: 15_000 })
  await menu.locator('.ctx-item', { hasText: 'Change Picture' }).hover()
  return menu.locator('.ctx-submenu .ctx-item')
}

test('Change Picture replaces the picture from the clipboard and from the icons, at its width and with its wrapping', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-change-picture-')))
  const docPath = join(dir, 'picture.docx')
  writeFileSync(
    docPath,
    await docxWith(
      para('One') +
        `<w:p>${pictureRun(1, 40, 10, 160, 120)}</w:p>` +
        ['Two', 'Three', 'Four', 'Five', 'Six', 'Seven'].map(para).join(''),
    ),
  )
  let launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-change-picture-sources',
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await ready(page)
    await expect(page.locator(PICTURE)).toHaveCount(1, { timeout: 30_000 })
    await expect.poll(() => size(page, PICTURE)).toEqual([160, 120])
    const src = await page.locator(PICTURE).getAttribute('src')

    // text on the clipboard: the clipboard is not a source
    await app.evaluate(({ clipboard }) => clipboard.writeText('no picture here'))
    let sources = await changePicture(page)
    await expect(sources.locator('.ctx-label')).toHaveText([
      'From a File…',
      'From Icons…',
      'From Clipboard…',
    ])
    await expect(sources.nth(2)).toBeDisabled()
    await page.keyboard.press('Escape')
    await expect(page.locator('.ctx-menu')).toBeHidden()

    // a picture twice as wide as it is tall on the clipboard
    await app.evaluate(({ clipboard, nativeImage }) => {
      const pixels = Buffer.alloc(40 * 20 * 4, 0)
      for (let i = 0; i < pixels.length; i += 4) pixels.set([0, 120, 220, 255], i)
      clipboard.writeImage(nativeImage.createFromBitmap(pixels, { width: 40, height: 20 }))
    })
    sources = await changePicture(page)
    await expect(sources.nth(2)).toBeEnabled()
    await page.screenshot({ path: test.info().outputPath('change-picture-menu.png') })
    await sources.nth(2).click()
    await expect.poll(() => page.locator(PICTURE).getAttribute('src')).not.toBe(src)
    // the width stays, the height follows the new picture
    await expect.poll(() => size(page, PICTURE)).toEqual([160, 80])
    const pasted = await page.locator(PICTURE).getAttribute('src')
    const onPaper = async () => {
      const [paper] = await rects(page, '.doc-page[contenteditable="true"]')
      const [at] = await rects(page, PICTURE)
      return [Math.round(at.left - paper.left), Math.round(at.top - paper.top)]
    }
    const place = await onPaper()

    // an icon: one is picked at a time, and it is drawn from its vector
    sources = await changePicture(page)
    await sources.nth(1).click()
    const dialog = page.locator('.icons-dialog')
    await expect(dialog).toBeVisible()
    const cells = dialog.locator('.icons-dialog-cell')
    await expect(cells.first()).toBeVisible()
    await cells.nth(0).click()
    await cells.nth(3).click()
    await expect(dialog.locator('.icons-dialog-cell.active')).toHaveCount(1)
    await expect(cells.nth(3)).toHaveClass(/active/)
    await page.screenshot({ path: test.info().outputPath('change-picture-icons.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect.poll(() => page.locator(PICTURE).getAttribute('src')).not.toBe(pasted)
    expect(await page.locator(PICTURE).getAttribute('src')).toMatch(/^data:image\/svg\+xml/)
    await expect.poll(() => size(page, PICTURE)).toEqual([160, 160])
    // it stands where the picture stood
    expect(await onPaper()).toEqual(place)
    await page.screenshot({ path: test.info().outputPath('changed-to-icon.png') })

    await saveFromMenu(app)
    await expect
      .poll(async () => (await partsOf(docPath)).media.some((name) => name.endsWith('.svg')), {
        timeout: 30_000,
      })
      .toBe(true)
    const saved = await partsOf(docPath)
    // one anchor, still wrapped square, at the icon's size, with the vector beside its bitmap
    expect(saved.xml.match(/<wp:anchor /g)).toHaveLength(1)
    expect(saved.xml).toContain('<wp:wrapSquare')
    expect(saved.xml).toContain(`<wp:extent cx="${160 * 9525}" cy="${160 * 9525}"`)
    expect(saved.xml).toContain('svgBlip')
    expect(saved.xml).toContain(`<wp:posOffset>${40 * 9525}</wp:posOffset>`)

    await closeAndSaveVideo(launched)
    launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-change-picture-sources',
      openFile: docPath,
    })
    const again = await waitForPageWithUrl(launched.app, '://docs/')
    await ready(again)
    await expect(again.locator(PICTURE)).toHaveCount(1, { timeout: 30_000 })
    await expect.poll(() => size(again, PICTURE)).toEqual([160, 160])
    expect(await again.locator(PICTURE).getAttribute('src')).toMatch(/^data:image\/svg\+xml/)
  } finally {
    await closeAndSaveVideo(launched)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a picture beside a shape in one paragraph turns with Rotate; Ungroup is offered on a group placed by offsets or by alignment', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-picture-frame-')))
  const docPath = join(dir, 'frame.docx')
  const children =
    groupShape(1, [0, 0, 100, 60], 'C00000', { text: 'Box' }) +
    groupShape(2, [140, 20, 80, 80], '00B050', { prst: 'ellipse' })
  const group = (id: number) => groupParagraph(id, [60, 20, 220, 100], children)
  const aligned = group(60).replace(
    /<wp:positionH[^>]*>.*?<\/wp:positionH>/,
    '<wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>',
  )
  writeFileSync(
    docPath,
    await docxWith(
      para('One') +
        shapeParagraph(1, 300, 10, 120, 60).replace(
          '</w:p>',
          `${pictureRun(2, 40, 10, 160, 80)}</w:p>`,
        ) +
        ['Two', 'Three', 'Four', 'Five', 'Six'].map(para).join('') +
        group(50) +
        ['Seven', 'Eight', 'Nine', 'Ten', 'Eleven', 'Twelve'].map(para).join('') +
        aligned +
        ['Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen'].map(para).join(''),
    ),
  )
  const FRAME = '.doc-page .doc-textbox-picture'
  const layout = async (page: Page, button: RegExp) => {
    const big = page.locator('.rb-big', { hasText: button })
    if (!(await big.isVisible())) await page.locator('.ribbon-tab', { hasText: /^Layout$/ }).click()
    return big
  }
  let launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-change-picture-sources',
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await ready(page)
    await expect(page.locator(FRAME)).toHaveCount(1, { timeout: 30_000 })
    await expect.poll(() => size(page, FRAME)).toEqual([160, 80])

    await page.locator(FRAME).click()
    const rotate = await layout(page, /^Rotate$/)
    await expect(rotate).toBeEnabled()
    await rotate.click()
    const right = page.locator('.layout-menu button', { hasText: 'Rotate Right 90°' })
    await expect(right).toBeEnabled()
    await page.screenshot({ path: test.info().outputPath('frame-rotate-menu.png') })
    await right.click()
    // a quarter turn: the picture now stands 80 wide and 160 tall
    await expect.poll(() => size(page, FRAME)).toEqual([80, 160])
    await page.screenshot({ path: test.info().outputPath('frame-rotated.png') })

    // the group placed by offsets comes apart, and so does the one centred between the margins
    const boxes = page.locator('.doc-page .doc-textbox:not(.doc-textbox-picture)')
    await expect(boxes).toHaveCount(5)
    await boxes.nth(1).click()
    const groupButton = await layout(page, /^Group$/)
    await expect(groupButton).toBeEnabled()
    await groupButton.click()
    await expect(page.locator('.layout-menu button', { hasText: /^Ungroup$/ })).toBeEnabled()
    await page.keyboard.press('Escape')
    await boxes.nth(3).scrollIntoViewIfNeeded()
    await boxes.nth(3).click()
    const alignedGroup = await layout(page, /^Group$/)
    await expect(alignedGroup).toBeEnabled()
    await alignedGroup.click()
    await expect(page.locator('.layout-menu button', { hasText: /^Ungroup$/ })).toBeEnabled()
    await page.screenshot({ path: test.info().outputPath('aligned-group-menu.png') })
    await page.keyboard.press('Escape')

    await saveFromMenu(app)
    await expect
      .poll(async () => (await partsOf(docPath)).xml.includes('rot="5400000"'), { timeout: 30_000 })
      .toBe(true)
    const { xml } = await partsOf(docPath)
    // the turn is on the picture, not on the shape beside it
    expect(xml).toMatch(/<pic:spPr><a:xfrm rot="5400000">/)
    expect(xml.match(/rot="/g)).toHaveLength(1)
    expect(xml.match(/<wpg:wgp[ >]/g)).toHaveLength(2)

    await closeAndSaveVideo(launched)
    launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-change-picture-sources',
      openFile: docPath,
    })
    const again = await waitForPageWithUrl(launched.app, '://docs/')
    await ready(again)
    await expect(again.locator(FRAME)).toHaveCount(1, { timeout: 30_000 })
    await expect.poll(() => size(again, FRAME)).toEqual([80, 160])
  } finally {
    await closeAndSaveVideo(launched)
    rmSync(dir, { recursive: true, force: true })
  }
})
