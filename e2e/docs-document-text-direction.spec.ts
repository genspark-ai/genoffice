/**
 * Format > Document… > Document Grid > Direction: a section is made vertical
 * and horizontal again from the dialog, on the page and in the saved file.
 */
import { test, expect, type ElectronApplication, type Locator, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-document-text-direction'
const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>'
const LONG = 'A long line of text that runs down the page when the section is vertical. '

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

async function caretIn(paragraph: Locator) {
  await paragraph.click()
  await paragraph.page().keyboard.press('End')
}

const sectPrs = async (docPath: string) => {
  const zip = await JSZip.loadAsync(readFileSync(docPath))
  const xml = (await zip.file('word/document.xml')?.async('string')) ?? ''
  return xml.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g) ?? []
}

const flowOf = (sectPr: string) => /<w:textDirection w:val="(\w+)"\/>/.exec(sectPr)?.[1] ?? null

/** how a paragraph is drawn: its writing mode, and whether its box is taller than wide */
const drawn = (paragraph: Locator) =>
  paragraph.evaluate((el) => {
    const box = el.getBoundingClientRect()
    return { mode: getComputedStyle(el).writingMode, tall: box.height > box.width }
  })

async function openGridTab(app: ElectronApplication, page: Page) {
  const dialog = page.locator('.doc-dialog:not(.borders-shading-dialog)')
  await clickMenu(app, ['Format', 'Document…'])
  await dialog.locator('[role="tab"]').nth(2).click()
  return {
    dialog,
    horizontal: dialog.locator('input[name="textDirection"][value="horizontal"]'),
    vertical: dialog.locator('input[name="textDirection"][value="vertical"]'),
  }
}

test('the Document dialog turns the document vertical and back, on the page and in the file', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-text-direction-')))
  const docPath = join(dir, 'flow.docx')
  writeFileSync(docPath, await docxWith(para(LONG.repeat(3)) + para('banana') + SECT_PR))
  let launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const long = page.locator('.doc-page p', { hasText: 'A long line' }).first()
    await expect(long).toBeVisible({ timeout: 30_000 })
    await caretIn(long)
    expect(await drawn(long)).toEqual({ mode: 'horizontal-tb', tall: false })

    let grid = await openGridTab(app, page)
    await expect(grid.horizontal).toBeChecked()
    await expect(grid.vertical).toBeEnabled()
    await grid.vertical.check()
    await page.screenshot({ path: test.info().outputPath('direction-dialog.png') })
    await grid.dialog.locator('button.primary').click()
    await expect(grid.dialog).toBeHidden()
    await expect.poll(() => drawn(long)).toEqual({ mode: 'vertical-rl', tall: true })
    await expect(page.locator('.doc-page .doc-vert-block')).toHaveCount(2)
    // vertical lines fill the page from its right edge
    const paper = page.locator('.doc-page').first()
    await expect
      .poll(async () => {
        const [p, sheet] = await Promise.all([long.boundingBox(), paper.boundingBox()])
        return !!p && !!sheet && p.x + p.width / 2 > sheet.x + sheet.width / 2
      })
      .toBe(true)
    await page.screenshot({ path: test.info().outputPath('direction-vertical-page.png') })

    // the sheets that print are vertical too
    await page.evaluate(() =>
      (window as unknown as { __openPagePreview?: () => void }).__openPagePreview?.(),
    )
    await expect(page.locator('.pv-page')).toHaveCount(1, { timeout: 30_000 })
    const printed = page.locator('.pv-page p', { hasText: 'A long line' }).first()
    await expect.poll(() => drawn(printed)).toEqual({ mode: 'vertical-rl', tall: true })
    await expect
      .poll(async () => {
        const [p, sheet] = await Promise.all([
          printed.boundingBox(),
          page.locator('.pv-page').first().boundingBox(),
        ])
        return !!p && !!sheet && p.x + p.width / 2 > sheet.x + sheet.width / 2
      })
      .toBe(true)
    await page.screenshot({ path: test.info().outputPath('direction-vertical-preview.png') })
    await page.keyboard.press('Escape')
    await expect(page.locator('.pv-page')).toHaveCount(0)

    // the dialog reads the section back without a reopen
    grid = await openGridTab(app, page)
    await expect(grid.vertical).toBeChecked()
    await page.keyboard.press('Escape')
    await expect(grid.dialog).toBeHidden()

    await saveFromMenu(app)
    await expect
      .poll(async () => (await sectPrs(docPath)).map(flowOf), { timeout: 30_000 })
      .toEqual(['tbRl'])
    // CT_SectPr order: after cols, before docGrid
    expect((await sectPrs(docPath))[0]).toMatch(
      /<w:cols\b[^>]*\/><w:textDirection w:val="tbRl"\/><w:docGrid\b/,
    )
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
  }

  // the saved file opens vertical, and goes back to horizontal from the dialog
  launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const long = page.locator('.doc-page p', { hasText: 'A long line' }).first()
    await expect(long).toBeVisible({ timeout: 30_000 })
    await expect.poll(() => drawn(long)).toEqual({ mode: 'vertical-rl', tall: true })
    await caretIn(long)
    const grid = await openGridTab(app, page)
    await expect(grid.vertical).toBeChecked()
    await grid.horizontal.check()
    await grid.dialog.locator('button.primary').click()
    await expect(grid.dialog).toBeHidden()
    await expect.poll(() => drawn(long)).toEqual({ mode: 'horizontal-tb', tall: false })
    await expect(page.locator('.doc-page .doc-vert-block')).toHaveCount(0)
    await saveFromMenu(app)
    await expect
      .poll(async () => (await sectPrs(docPath)).map(flowOf), { timeout: 30_000 })
      .toEqual([null])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('the direction goes with Apply to and the rest of the dialog in one write', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-text-direction-')))
  const docPath = join(dir, 'sections.docx')
  writeFileSync(docPath, await docxWith(['apple', 'banana', 'cherry'].map(para).join('') + SECT_PR))
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const paragraph = (text: string) => page.locator('.doc-page p', { hasText: text }).first()
    await expect(paragraph('apple')).toBeVisible({ timeout: 30_000 })
    const saved = async () =>
      (await sectPrs(docPath)).map(
        (sectPr) => `${/<w:pgMar\b[^>]* w:left="(\d+)"/.exec(sectPr)?.[1]} ${flowOf(sectPr)}`,
      )

    // from the caret on, with a margin changed in the same OK: both land on the new section only
    await caretIn(paragraph('apple'))
    await clickMenu(app, ['Format', 'Document…'])
    const dialog = page.locator('.doc-dialog:not(.borders-shading-dialog)')
    const left = dialog.getByLabel('Left', { exact: true })
    await left.fill('144pt')
    await left.blur()
    await dialog.locator('[role="tab"]').nth(2).click()
    await dialog.locator('input[name="textDirection"][value="vertical"]').check()
    await dialog.locator('select[name="applyTo"]').selectOption('forward')
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect.poll(() => drawn(paragraph('banana')).then((d) => d.mode)).toBe('vertical-rl')
    expect((await drawn(paragraph('apple'))).mode).toBe('horizontal-tb')
    await saveFromMenu(app)
    await expect.poll(saved, { timeout: 30_000 }).toEqual(['1440 null', '2880 tbRl'])

    // this section only: the first section turns vertical, the second stays as it is
    await caretIn(paragraph('apple'))
    let grid = await openGridTab(app, page)
    await expect(grid.horizontal).toBeChecked()
    await expect(grid.dialog.locator('select[name="applyTo"]')).toHaveValue('section')
    await grid.vertical.check()
    await grid.dialog.locator('button.primary').click()
    await expect(grid.dialog).toBeHidden()
    await expect.poll(() => drawn(paragraph('apple')).then((d) => d.mode)).toBe('vertical-rl')
    await saveFromMenu(app)
    await expect.poll(saved, { timeout: 30_000 }).toEqual(['1440 tbRl', '2880 tbRl'])

    // the whole document back to horizontal: every section loses the tag, the margins stay
    grid = await openGridTab(app, page)
    await grid.horizontal.check()
    await grid.dialog.locator('select[name="applyTo"]').selectOption('whole')
    await grid.dialog.locator('button.primary').click()
    await expect(grid.dialog).toBeHidden()
    await expect(page.locator('.doc-page .doc-vert-block')).toHaveCount(0)
    await saveFromMenu(app)
    await expect.poll(saved, { timeout: 30_000 }).toEqual(['1440 null', '1440 null'])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
