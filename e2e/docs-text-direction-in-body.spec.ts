/**
 * Layout > Text Direction is a menu of the directions of what the caret is in:
 * the section in body text, the cells in a table, the text box when one is
 * selected. Its last item and Format > Text Direction… open the dialog.
 */
import { test, expect, type ElectronApplication, type Locator, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-text-direction-in-body'
const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>'
const cell = (t: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${para(t)}</w:tc>`
const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
  `<w:tr>${cell('cell one')}${cell('cell two')}</w:tr></w:tbl>`
const LONG = 'A long line of text that runs down the page when the section is vertical. '

type MenuItemState = { enabled: boolean } | null

const menuItem = (app: ElectronApplication, path: string[]): Promise<MenuItemState> =>
  app.evaluate(({ Menu }, path) => {
    let items = Menu.getApplicationMenu()?.items
    let item: Electron.MenuItem | undefined
    for (const label of path) {
      item = items?.find((candidate) => candidate.label === label)
      items = item?.submenu?.items
    }
    return item ? { enabled: item.enabled } : null
  }, path)

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

const documentXml = async (docPath: string) => {
  const zip = await JSZip.loadAsync(readFileSync(docPath))
  return (await zip.file('word/document.xml')?.async('string')) ?? ''
}
const sectionFlow = async (docPath: string) => {
  const sectPr = /<w:sectPr\b[\s\S]*?<\/w:sectPr>/.exec(await documentXml(docPath))?.[0] ?? ''
  return /<w:textDirection w:val="(\w+)"\/>/.exec(sectPr)?.[1] ?? null
}

const drawn = (paragraph: Locator) =>
  paragraph.evaluate((el) => {
    const box = el.getBoundingClientRect()
    return { mode: getComputedStyle(el).writingMode, tall: box.height > box.width }
  })

const button = (page: Page) =>
  page.locator('[data-ribbon-body] button:is(.rb-big, .rb-small)[data-tip="Text Direction"]')
const items = (page: Page) => page.locator('[data-ribbon-body] .layout-menu button')

async function openMenu(page: Page) {
  const tab = page.locator('.ribbon-tab', { hasText: /^Layout$/ }).first()
  // a click on the tab that is already showing folds the ribbon away
  if (!(await tab.evaluate((el) => el.classList.contains('active')))) await tab.click()
  await expect(button(page)).toBeEnabled()
  if ((await items(page).count()) === 0) await button(page).click()
  await expect(items(page).last()).toHaveText('Text Direction Options…')
}

async function pick(page: Page, label: string) {
  await openMenu(page)
  await items(page)
    .filter({ hasText: new RegExp(`^${label}$`) })
    .click()
  await expect(items(page)).toHaveCount(0)
}

const checked = async (page: Page) => {
  await openMenu(page)
  const label = await page
    .locator('[data-ribbon-body] .layout-menu button.active')
    .allTextContents()
  await button(page).click()
  await expect(items(page)).toHaveCount(0)
  return label
}

test('Text Direction turns the section in body text, the cell in a table and a selected text box', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-text-direction-body-')))
  const docPath = join(dir, 'body.docx')
  writeFileSync(docPath, await docxWith(para(LONG.repeat(3)) + TABLE + para('banana') + SECT_PR))
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const long = page.locator('.doc-page p', { hasText: 'A long line' }).first()
    await expect(long).toBeVisible({ timeout: 30_000 })
    await long.click()
    await expect
      .poll(() => menuItem(app, ['Format', 'Text Direction…']))
      .toMatchObject({ enabled: true })

    // body text: the section's two directions, then the dialog
    await openMenu(page)
    await expect(items(page)).toHaveText(['Horizontal', 'Vertical', 'Text Direction Options…'])
    await expect(page.locator('[data-ribbon-body] .layout-menu button.active')).toHaveText(
      'Horizontal',
    )
    await page.screenshot({ path: test.info().outputPath('menu-body.png') })
    await pick(page, 'Vertical')
    await expect.poll(() => drawn(long)).toEqual({ mode: 'vertical-rl', tall: true })
    await expect(page.locator('.doc-page .doc-vert-block').first()).toBeVisible()
    expect(await checked(page)).toEqual(['Vertical'])
    await page.screenshot({ path: test.info().outputPath('page-vertical.png') })

    // Format > Document… shows what the ribbon set
    const documentDialog = page.locator('.doc-dialog:not(.borders-shading-dialog)')
    await clickMenu(app, ['Format', 'Document…'])
    await documentDialog.locator('[role="tab"]').nth(2).click()
    await expect(
      documentDialog.locator('input[name="textDirection"][value="vertical"]'),
    ).toBeChecked()
    await page.keyboard.press('Escape')
    await expect(documentDialog).toBeHidden()

    await saveFromMenu(app)
    await expect.poll(() => sectionFlow(docPath), { timeout: 30_000 }).toBe('tbRl')

    // one undo step back to horizontal, one redo forward
    await clickMenu(app, ['Edit', 'Undo'])
    await expect.poll(() => drawn(long)).toEqual({ mode: 'horizontal-tb', tall: false })
    expect(await checked(page)).toEqual(['Horizontal'])
    await clickMenu(app, ['Edit', 'Redo'])
    await expect.poll(() => drawn(long)).toEqual({ mode: 'vertical-rl', tall: true })

    // the dialog opens on the section's direction and names what it applies to
    await openMenu(page)
    await items(page).last().click()
    const dialog = page.locator('.text-direction-dialog')
    const radios = dialog.locator('input[name="direction"]')
    await expect(radios).toHaveCount(2)
    await expect(radios.nth(1)).toBeChecked()
    await expect(dialog.locator('select[name="applyTo"] option')).toHaveText(['This section'])
    await page.screenshot({ path: test.info().outputPath('dialog-body.png') })
    await radios.nth(0).check()
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect.poll(() => drawn(long)).toEqual({ mode: 'horizontal-tb', tall: false })
    await expect(page.locator('.doc-page .doc-vert-block')).toHaveCount(0)
    await saveFromMenu(app)
    await expect.poll(() => sectionFlow(docPath), { timeout: 30_000 }).toBeNull()

    // in a table the same control turns the cell, not the section
    const first = page.locator('.doc-page .doc-table').first().locator('td, th').first()
    await first.click()
    await openMenu(page)
    await expect(items(page)).toHaveText([
      'Horizontal',
      'Rotate all text 90°',
      'Rotate all text 270°',
      'Text Direction Options…',
    ])
    await page.screenshot({ path: test.info().outputPath('menu-cell.png') })
    await pick(page, 'Rotate all text 90°')
    await expect.poll(() => first.evaluate((el) => el.outerHTML.includes('vertical-rl'))).toBe(true)
    expect((await drawn(long)).mode).toBe('horizontal-tb')
    expect(await checked(page)).toEqual(['Rotate all text 90°'])
    await clickMenu(app, ['Format', 'Text Direction…'])
    await expect(radios).toHaveCount(3)
    await expect(radios.nth(1)).toBeChecked()
    await expect(dialog.locator('select[name="applyTo"] option')).toHaveText(['Selected cells'])
    await page.keyboard.press('Escape')
    await expect(dialog).toBeHidden()
    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const xml = await documentXml(docPath)
          return [
            /<w:tcPr>(?:(?!<\/w:tcPr>)[\s\S])*<w:textDirection w:val="tbRl"\/>/.test(xml),
            await sectionFlow(docPath),
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([true, null])

    // a text box that is being typed in is the target, with its own directions
    const docPage = page.locator('.doc-page[contenteditable="true"]')
    await page.locator('.doc-page p', { hasText: 'banana' }).first().click()
    await clickMenu(app, ['Insert', 'Text Box', 'Text Box'])
    await expect(docPage).toHaveClass(/doc-shape-drawing/)
    const paper = (await docPage.boundingBox())!
    await page.mouse.click(paper.x + 300, paper.y + 420)
    const box = docPage.locator('.doc-textbox')
    await expect(box).toHaveCount(1)
    await expect
      .poll(() => page.evaluate(() => !!document.activeElement?.closest('.doc-textbox-editor')))
      .toBe(true)
    await page.keyboard.type('Turn me', { delay: 10 })
    await openMenu(page)
    await expect(items(page)).toHaveText([
      'Horizontal',
      'Vertical',
      'Rotate all text 90°',
      'Rotate all text 270°',
      'Stacked',
      'Text Direction Options…',
    ])
    await pick(page, 'Rotate all text 270°')
    await expect(box).toHaveClass(/doc-textbox-vert270/)
    await expect(box).toContainText('Turn me')
    expect((await drawn(long)).mode).toBe('horizontal-tb')
    await saveFromMenu(app)
    await expect
      .poll(async () => (await documentXml(docPath)).includes('<wps:bodyPr vert="vert270"'), {
        timeout: 30_000,
      })
      .toBe(true)
    expect(await sectionFlow(docPath)).toBeNull()
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
