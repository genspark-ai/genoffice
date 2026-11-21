/**
 * Page setup is one step in the undo history, in order with text edits:
 * the Layout tab's Margins, Orientation, Size and Columns, the Document
 * dialog and the Columns dialog.
 */
import { test, expect, type ElectronApplication, type Locator, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-page-setup-undo'
const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/></w:sectPr>'

const menuItem = (app: ElectronApplication, path: string[]) =>
  app.evaluate(({ Menu }, path) => {
    let items = Menu.getApplicationMenu()?.items
    let item: Electron.MenuItem | undefined
    for (const label of path) {
      item = items?.find((candidate) => candidate.label === label)
      items = item?.submenu?.items
    }
    return item ? { enabled: item.enabled } : null
  }, path)

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

const part = async (docPath: string, name: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file(name)?.async('string') ?? ''

/** every sectPr of the saved file, in document order */
const sectPrs = async (docPath: string) =>
  (await part(docPath, 'word/document.xml')).match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g) ?? []

const attr = (xml: string, tag: string, name: string) =>
  new RegExp(`<w:${tag}\\b[^>]* w:${name}="([^"]*)"`).exec(xml)?.[1] ?? null

/**
 * The first page as drawn, in hundredths of its own width so the zoom does not
 * matter: its height, and where its first paragraph starts and how wide it is.
 */
const pageBox = (page: Page) =>
  page.evaluate(() => {
    const sheet = document.querySelector('.doc-page')!.getBoundingClientRect()
    const first = document.querySelector('.doc-page p')!.getBoundingClientRect()
    const share = (px: number) => Math.round((px / sheet.width) * 100)
    return {
      landscape: sheet.width > sheet.height,
      height: share(sheet.height),
      left: share(first.left - sheet.left),
      column: share(first.width),
    }
  })

async function openDoc(docPath: string) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  const paragraph = (text: string) => page.locator('.doc-page p', { hasText: text }).first()
  await expect(paragraph('apple')).toBeVisible({ timeout: 30_000 })
  return { launched, app: launched.app, page, paragraph }
}

/** picks an item of one of the Layout tab's page setup menus */
async function layoutPick(page: Page, control: string, item: string | RegExp) {
  const button = page
    .locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: control })
    .first()
  // clicking the tab that is already open folds the ribbon away
  if (!(await button.isVisible()))
    await page.locator('.ribbon-tab', { hasText: /^Layout$/ }).click()
  await button.click()
  await page.locator('[data-rb-panel] button', { hasText: item }).first().click()
}

test('Layout tab page setup undoes and redoes in order with typing', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-page-setup-undo-')))
  const docPath = join(dir, 'layout.docx')
  writeFileSync(docPath, await docxWith(['apple', 'banana', 'cherry'].map(para).join('') + SECT_PR))
  const { launched, app, page, paragraph } = await openDoc(docPath)
  try {
    const start = await pageBox(page)
    expect(start.landscape).toBe(false)
    expect((await menuItem(app, ['Edit', 'Undo']))?.enabled).toBe(false)

    await caretIn(paragraph('apple'))
    await page.keyboard.type(' one')
    await layoutPick(page, 'Margins', /^Narrow/)
    await expect.poll(async () => (await pageBox(page)).left).toBeLessThan(start.left)
    const narrow = await pageBox(page)
    await caretIn(paragraph('apple one'))
    await page.keyboard.type(' two')
    await layoutPick(page, 'Orientation', 'Landscape')
    await expect.poll(async () => (await pageBox(page)).landscape).toBe(true)
    const landscape = await pageBox(page)
    await layoutPick(page, 'Size', /^A4/)
    await expect.poll(async () => (await pageBox(page)).height).toBeLessThan(landscape.height)
    const a4 = await pageBox(page)
    await layoutPick(page, 'Columns', /^Two$/)
    await expect.poll(async () => (await pageBox(page)).column).toBeLessThan(a4.column * 0.6)
    await page.screenshot({ path: test.info().outputPath('layout-changed.png') })

    // each change is its own step, the newest first
    await clickMenu(app, ['Edit', 'Undo'])
    await expect.poll(() => pageBox(page)).toEqual(a4)
    await clickMenu(app, ['Edit', 'Undo'])
    await expect.poll(() => pageBox(page)).toEqual(landscape)
    await clickMenu(app, ['Edit', 'Undo'])
    await expect.poll(() => pageBox(page)).toEqual(narrow)
    // the typing between two page setup changes is undone between them
    await expect(paragraph('apple one two')).toBeVisible()
    await clickMenu(app, ['Edit', 'Undo'])
    await expect(paragraph('apple one two')).toHaveCount(0)
    await expect.poll(() => pageBox(page)).toEqual(narrow)
    await clickMenu(app, ['Edit', 'Undo'])
    await expect.poll(() => pageBox(page)).toEqual(start)
    await expect(paragraph('apple one')).toBeVisible()
    await clickMenu(app, ['Edit', 'Undo'])
    await expect(paragraph('apple one')).toHaveCount(0)
    await expect.poll(async () => (await menuItem(app, ['Edit', 'Undo']))?.enabled).toBe(false)
    await page.screenshot({ path: test.info().outputPath('layout-undone.png') })

    // redo walks the same steps forward
    await clickMenu(app, ['Edit', 'Redo'])
    await expect(paragraph('apple one')).toBeVisible()
    await clickMenu(app, ['Edit', 'Redo'])
    await expect.poll(() => pageBox(page)).toEqual(narrow)
    await clickMenu(app, ['Edit', 'Redo'])
    await expect(paragraph('apple one two')).toBeVisible()
    await clickMenu(app, ['Edit', 'Redo'])
    await expect.poll(() => pageBox(page)).toEqual(landscape)

    // an edit made now discards the two page setup steps still to redo
    await caretIn(paragraph('banana'))
    await page.keyboard.type('!')
    await expect.poll(() => menuItem(app, ['Edit', 'Redo'])).toBeNull()
    await saveFromMenu(app)
    await expect
      .poll(async () => attr((await sectPrs(docPath))[0] ?? '', 'pgSz', 'orient'), {
        timeout: 30_000,
      })
      .toBe('landscape')
    const [sectPr] = await sectPrs(docPath)
    expect(attr(sectPr, 'pgSz', 'w')).toBe('15840')
    expect(attr(sectPr, 'pgMar', 'left')).toBe('720')
    expect(attr(sectPr, 'cols', 'num')).toBeNull()
    await expect.poll(() => pageBox(page)).toEqual(landscape)

    // the history outlives the save: undoing past it writes the earlier setup on the next save
    await clickMenu(app, ['Edit', 'Undo'])
    await clickMenu(app, ['Edit', 'Undo'])
    await expect.poll(() => pageBox(page)).toEqual(narrow)
    await saveFromMenu(app)
    await expect
      .poll(async () => attr((await sectPrs(docPath))[0] ?? '', 'pgSz', 'w'), { timeout: 30_000 })
      .toBe('12240')
    expect(attr((await sectPrs(docPath))[0], 'pgMar', 'left')).toBe('720')
    expect(await part(docPath, 'word/document.xml')).toContain('apple one two')
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('one OK in the Document dialog is one undo step', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-page-setup-undo-')))
  const docPath = join(dir, 'document.docx')
  writeFileSync(docPath, await docxWith(['apple', 'banana'].map(para).join('') + SECT_PR))
  const { launched, app, page, paragraph } = await openDoc(docPath)
  try {
    const start = await pageBox(page)
    const dialog = page.locator('.doc-dialog:not(.borders-shading-dialog)')
    const left = dialog.getByLabel('Left', { exact: true })
    const pages = dialog.locator('select', { has: page.locator('option[value="mirror"]') })
    const check = (label: string) =>
      dialog.locator('label.tco-check', { hasText: label }).locator('input')
    const firstPage = check('Different First Page')
    const oddEven = check('Different Odd')

    await caretIn(paragraph('apple'))
    await page.keyboard.type(' pie')
    await clickMenu(app, ['Format', 'Document…'])
    const opened = {
      left: await left.inputValue(),
      gutter: await dialog.getByLabel('Gutter', { exact: true }).inputValue(),
    }
    await left.fill('144pt')
    await left.blur()
    const gutter = dialog.getByLabel('Gutter', { exact: true })
    await gutter.fill('36pt')
    await gutter.blur()
    await pages.selectOption('mirror')
    await dialog.locator('[role="tab"]').nth(1).click()
    await oddEven.check()
    await firstPage.check()
    await page.screenshot({ path: test.info().outputPath('document-dialog.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect.poll(async () => (await pageBox(page)).left).toBeGreaterThan(start.left)
    const wide = await pageBox(page)

    // margins, mirror margins and the two header choices went in with one OK and come out with one Undo
    await clickMenu(app, ['Edit', 'Undo'])
    await expect.poll(() => pageBox(page)).toEqual(start)
    await expect(paragraph('apple pie')).toBeVisible()
    await clickMenu(app, ['Format', 'Document…'])
    await expect(pages).toHaveValue('normal')
    await dialog.locator('[role="tab"]').nth(1).click()
    await expect(oddEven).not.toBeChecked()
    await expect(firstPage).not.toBeChecked()
    await page.keyboard.press('Escape')
    await expect(dialog).toBeHidden()
    await saveFromMenu(app)
    await expect
      .poll(async () => (await part(docPath, 'word/document.xml')).includes('apple pie'), {
        timeout: 30_000,
      })
      .toBe(true)
    let [sectPr] = await sectPrs(docPath)
    expect(attr(sectPr, 'pgMar', 'left')).toBe('1440')
    expect(sectPr).not.toContain('<w:titlePg')
    expect(await part(docPath, 'word/settings.xml')).not.toMatch(
      /<w:(mirrorMargins|evenAndOddHeaders)\b/,
    )

    // and go back in with one Redo, after the save
    await clickMenu(app, ['Edit', 'Redo'])
    await expect.poll(() => pageBox(page)).toEqual(wide)
    await clickMenu(app, ['Format', 'Document…'])
    await expect(pages).toHaveValue('mirror')
    await dialog.locator('[role="tab"]').nth(1).click()
    await expect(oddEven).toBeChecked()
    await expect(firstPage).toBeChecked()
    await page.keyboard.press('Escape')
    await saveFromMenu(app)
    await expect
      .poll(async () => attr((await sectPrs(docPath))[0] ?? '', 'pgMar', 'left'), {
        timeout: 30_000,
      })
      .toBe('2880')
    ;[sectPr] = await sectPrs(docPath)
    expect(sectPr).toContain('<w:titlePg')
    const settings = await part(docPath, 'word/settings.xml')
    expect(settings).toContain('<w:mirrorMargins')
    expect(settings).toContain('<w:evenAndOddHeaders')
    expect(attr(sectPr, 'pgMar', 'gutter')).toBe('720')

    // undone once more, now that the file states them: the save takes them out again
    await clickMenu(app, ['Edit', 'Undo'])
    await expect.poll(() => pageBox(page)).toEqual(start)
    await saveFromMenu(app)
    await expect
      .poll(async () => attr((await sectPrs(docPath))[0] ?? '', 'pgMar', 'gutter'), {
        timeout: 30_000,
      })
      .toBe('0')
    ;[sectPr] = await sectPrs(docPath)
    expect(attr(sectPr, 'pgMar', 'left')).toBe('1440')
    expect(sectPr).not.toContain('<w:titlePg')
    expect(await part(docPath, 'word/settings.xml')).not.toMatch(
      /<w:(mirrorMargins|evenAndOddHeaders)\b/,
    )

    // and the file opens again as it was before the dialog
    await clickMenu(app, ['File', 'New Document'])
    await expect(paragraph('apple pie')).toHaveCount(0, { timeout: 30_000 })
    await page.evaluate(
      (path) =>
        (window as never as { __aidocs: { openPath(path: string): unknown } }).__aidocs.openPath(
          path,
        ),
      docPath,
    )
    await expect(paragraph('apple pie')).toBeVisible({ timeout: 30_000 })
    await expect.poll(() => pageBox(page)).toEqual(start)
    await caretIn(paragraph('apple pie'))
    await clickMenu(app, ['Format', 'Document…'])
    await expect(left).toHaveValue(opened.left)
    await expect(gutter).toHaveValue(opened.gutter)
    await expect(pages).toHaveValue('normal')
    await dialog.locator('[role="tab"]').nth(1).click()
    await expect(oddEven).not.toBeChecked()
    await expect(firstPage).not.toBeChecked()
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('page setup of one section and of the whole document undoes where it was made', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-page-setup-undo-')))
  const docPath = join(dir, 'sections.docx')
  const first = SECT_PR.replace('</w:sectPr>', '<w:type w:val="nextPage"/></w:sectPr>')
  writeFileSync(
    docPath,
    await docxWith(para('apple') + `<w:p><w:pPr>${first}</w:pPr></w:p>` + para('banana') + SECT_PR),
  )
  const { launched, app, page, paragraph } = await openDoc(docPath)
  try {
    const cols = async () => (await sectPrs(docPath)).map((s) => attr(s, 'cols', 'num') ?? '1')
    const lefts = async () => (await sectPrs(docPath)).map((s) => attr(s, 'pgMar', 'left'))
    const columns = page.locator('.columns-dialog')

    await caretIn(paragraph('apple'))
    await page.locator('[data-ribbon-body] button[aria-label^="Center"]').first().click()
    await expect.poll(() => menuItem(app, ['Edit', 'Repeat'])).toEqual({ enabled: true })
    // the first section alone, from the Layout tab
    await layoutPick(page, 'Margins', /^Narrow/)
    // page setup is the last thing done now: Repeat no longer offers Center
    await expect.poll(() => menuItem(app, ['Edit', "Can't Repeat"])).toEqual({ enabled: false })
    // every section, from the Columns dialog
    await clickMenu(app, ['Format', 'Columns…'])
    await columns.locator('input[name="columns"]').nth(1).check()
    await columns.locator('select[name="applyTo"]').selectOption('whole')
    await columns.locator('button.primary').click()
    await expect(columns).toBeHidden()
    await saveFromMenu(app)
    await expect.poll(cols, { timeout: 30_000 }).toEqual(['2', '2'])
    expect(await lefts()).toEqual(['720', '1440'])

    await clickMenu(app, ['Edit', 'Undo'])
    await saveFromMenu(app)
    await expect.poll(cols, { timeout: 30_000 }).toEqual(['1', '1'])
    expect(await lefts()).toEqual(['720', '1440'])

    await clickMenu(app, ['Edit', 'Undo'])
    await saveFromMenu(app)
    await expect.poll(lefts, { timeout: 30_000 }).toEqual(['1440', '1440'])
    await clickMenu(app, ['Edit', 'Undo'])
    await expect.poll(async () => (await menuItem(app, ['Edit', 'Undo']))?.enabled).toBe(false)

    await clickMenu(app, ['Edit', 'Redo'])
    await clickMenu(app, ['Edit', 'Redo'])
    await clickMenu(app, ['Edit', 'Redo'])
    await expect.poll(() => menuItem(app, ['Edit', 'Redo'])).toBeNull()
    await saveFromMenu(app)
    await expect.poll(cols, { timeout: 30_000 }).toEqual(['2', '2'])
    expect(await lefts()).toEqual(['720', '1440'])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
