/**
 * Word's Table menu: Convert Text to Table, Convert Table to Text, Sort and
 * Formula, and the Insert tab's Object control. Each is checked on the page
 * and in the saved file.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-table-convert-formula'
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

const select = (page: Page, from: number, to: number) =>
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

const documentXml = async (docPath: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file('word/document.xml')?.async('string') ?? ''

/** the first table on the page, as rows of cell texts */
const tableRows = (page: Page) =>
  page
    .locator('.doc-page table')
    .first()
    .locator('tr')
    .evaluateAll((rows) =>
      rows.map((row) =>
        Array.from(row.querySelectorAll('td, th'), (cell) => (cell.textContent ?? '').trim()),
      ),
    )

const cell = (page: Page, row: number, col: number) =>
  page.locator('.doc-page table').first().locator('tr').nth(row).locator('td, th').nth(col)

const tc = (inner: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`
const row = (...cells: string[]) => `<w:tr>${cells.map(tc).join('')}</w:tr>`
const formulaPara = (instr: string, cached: string) =>
  '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  `<w:r><w:t>${cached}</w:t></w:r>` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'

async function openDoc(docPath: string) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  return { launched, app: launched.app, page }
}

test('text becomes a table, a formula totals a column and follows its cells, and the table becomes text again', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-table-convert-')))
  const docPath = join(dir, 'convert.docx')
  writeFileSync(
    docPath,
    await docxWith(
      para('Item,Qty') + para('Apple,3') + para('Pear,4') + para('Total,') + para('tail') + SECT_PR,
    ),
  )
  let session = await openDoc(docPath)
  try {
    let { app, page } = session
    await expect(page.locator('.doc-page p', { hasText: 'Item,Qty' })).toBeVisible({
      timeout: 30_000,
    })

    // no embedded object can be made: Object… is not offered, Text from File… is
    expect(await menuItem(app, ['Insert', 'Object…'])).toBeNull()
    await page
      .locator('.ribbon-tab', { hasText: /^Insert$/ })
      .first()
      .click()
    const objectButton = page.locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', {
      hasText: 'Object',
    })
    await objectButton.click()
    const objectMenu = page.locator('[data-rb-panel]')
    await expect(objectMenu.getByRole('button', { name: 'Object…' })).toHaveCount(0)
    await expect(objectMenu.getByRole('button', { name: 'Text from File…' })).toBeEnabled()
    await page.screenshot({ path: test.info().outputPath('object-dropdown.png') })
    await page.keyboard.press('Escape')

    // the four comma-separated paragraphs: Word proposes commas and two columns
    await select(page, 1, 30)
    await clickMenu(app, ['Table', 'Convert', 'Convert Text to Table…'])
    const toTable = page.locator('.convert-text-dialog')
    await expect(toTable.locator('input[name="separator"][value="comma"]')).toBeChecked()
    await expect(toTable.locator('input[name="cols"]')).toHaveValue('2')
    await expect(toTable.locator('input[name="rows"]')).toHaveValue('4')
    await page.screenshot({ path: test.info().outputPath('convert-text-to-table.png') })
    await toTable.locator('button.primary').click()
    await expect
      .poll(() => tableRows(page))
      .toEqual([
        ['Item', 'Qty'],
        ['Apple', '3'],
        ['Pear', '4'],
        ['Total', ''],
      ])
    await expect(page.locator('.doc-page p', { hasText: 'tail' })).toBeVisible()
    // inside a table there is no text to convert
    await cell(page, 3, 1).click()
    await expect
      .poll(() => menuItem(app, ['Table', 'Convert', 'Convert Text to Table…']))
      .toEqual({
        enabled: false,
      })

    // Formula: Word proposes a sum of the column above
    await clickMenu(app, ['Table', 'Formula…'])
    const formula = page.locator('.formula-dialog')
    await expect(formula.locator('input[name="formula"]')).toHaveValue('=SUM(ABOVE)')
    await expect(formula.locator('select[name="pasteBookmark"]')).toBeEnabled()
    // a formula outside what is evaluated is refused, not written
    await formula.locator('input[name="formula"]').fill('=SUM(NoSuchBookmark)')
    await formula.locator('button.primary').click()
    await expect(formula.locator('[role="alert"]')).toBeVisible()
    await formula.locator('input[name="formula"]').fill('=SUM(ABOVE)')
    await formula.locator('input[name="format"]').fill('0.00')
    await page.screenshot({ path: test.info().outputPath('formula-dialog.png') })
    await formula.locator('button.primary').click()
    await expect(formula).toBeHidden()
    await expect.poll(async () => (await tableRows(page))[3]).toEqual(['Total', '7.00'])

    // the total follows the cells it reads
    await cell(page, 1, 1).click()
    await page.keyboard.press('End')
    await page.keyboard.type('5')
    await expect.poll(async () => (await tableRows(page))[3]).toEqual(['Total', '39.00'])
    await page.screenshot({ path: test.info().outputPath('formula-result.png') })

    // the Table menu's Sort… is the Home tab's Sort
    await clickMenu(app, ['Table', 'Sort…'])
    const sort = page.locator('.sort-dialog')
    await expect(sort.locator('select[name="key1"]')).toBeVisible()
    await page.keyboard.press('Escape')
    await expect(sort).toBeHidden()

    await saveFromMenu(app)
    await expect
      .poll(() => documentXml(docPath), { timeout: 30_000 })
      .toMatch(
        /<w:tbl>[\s\S]*<w:fldChar w:fldCharType="begin"\/>[\s\S]*?<w:instrText[^>]*> ?=SUM\(ABOVE\) \\# "0\.00" ?<\/w:instrText>[\s\S]*?<w:t[^>]*>39\.00<\/w:t>[\s\S]*<\/w:tbl>/,
      )
    // the result is the file's: no field is left for Word to update on open
    expect(await documentXml(docPath)).not.toContain('w:dirty')

    // reopened: the table, the result, and the dialog on the field shows its formula
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    session = await openDoc(docPath)
    ;({ app, page } = session)
    await expect
      .poll(() => tableRows(page), { timeout: 30_000 })
      .toEqual([
        ['Item', 'Qty'],
        ['Apple', '35'],
        ['Pear', '4'],
        ['Total', '39.00'],
      ])
    await cell(page, 3, 1).click()
    await clickMenu(app, ['Table', 'Formula…'])
    const reopened = page.locator('.formula-dialog')
    await expect(reopened.locator('input[name="formula"]')).toHaveValue('=SUM(ABOVE)')
    await expect(reopened.locator('input[name="format"]')).toHaveValue('0.00')
    await page.keyboard.press('Escape')
    await expect(reopened).toBeHidden()

    // Convert Table to Text opens on paragraph marks; tabs make one line of each row
    await cell(page, 0, 0).click()
    await clickMenu(app, ['Table', 'Convert', 'Convert Table to Text…'])
    const toText = page.locator('.convert-table-dialog')
    await expect(toText.locator('input[name="separator"][value="paragraph"]')).toBeChecked()
    await expect(toText.locator('input[name="nested"]')).toBeChecked()
    await toText.locator('input[name="separator"][value="tab"]').check()
    await page.screenshot({ path: test.info().outputPath('convert-table-to-text.png') })
    await toText.locator('button.primary').click()
    await expect(page.locator('.doc-page table')).toHaveCount(0)
    await expect
      .poll(() =>
        page
          .locator('.doc-page p')
          .evaluateAll((ps) => ps.map((p) => (p.textContent ?? '').replace(/\s+/g, ' ').trim())),
      )
      .toEqual(expect.arrayContaining(['Item Qty', 'Apple 35', 'Pear 4', 'Total 39.00', 'tail']))
    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const xml = await documentXml(docPath)
          return {
            table: xml.includes('<w:tbl>'),
            tabs: (xml.match(/<w:tab\/>/g) ?? []).length,
            field: /=SUM\(ABOVE\)/.test(xml),
          }
        },
        { timeout: 30_000 },
      )
      .toEqual({ table: false, tabs: 4, field: true })
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a formula a document arrives with keeps its result until Update Field, and the Insert tab converts text too', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-table-formula-')))
  const docPath = join(dir, 'stale.docx')
  writeFileSync(
    docPath,
    await docxWith(
      '<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr>' +
        '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
        row(para('North'), para('4')) +
        row(para('South'), para('6')) +
        row(para('All'), formulaPara('=SUM(ABOVE)', '99')) +
        row(para('Rule'), formulaPara('=SUM(NoSuchBookmark)', '1')) +
        '</w:tbl>' +
        '<w:p><w:r><w:t>one</w:t><w:tab/><w:t>two</w:t></w:r></w:p>' +
        SECT_PR,
    ),
  )
  const { launched, app, page } = await openDoc(docPath)
  try {
    // opening recomputes nothing: the cached results show as they were saved
    await expect
      .poll(() => tableRows(page), { timeout: 30_000 })
      .toEqual([
        ['North', '4'],
        ['South', '6'],
        ['All', '99'],
        ['Rule', '1'],
      ])
    await cell(page, 2, 1).click()
    await page.keyboard.press('F9')
    // the sum is recomputed; the formula outside what is evaluated keeps Word's result
    await expect
      .poll(async () => (await tableRows(page)).map((cells) => cells[1]))
      .toEqual(['4', '6', '10', '1'])

    // the Insert tab's Table button converts text too, through the same command
    await page.locator('.doc-page p', { hasText: 'one' }).first().click()
    await page
      .locator('.ribbon-tab', { hasText: /^Insert$/ })
      .first()
      .click()
    await page
      .locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: 'Table' })
      .first()
      .click()
    await page.locator('[data-rb-panel] button', { hasText: 'Convert Text to Table…' }).click()
    const toTable = page.locator('.convert-text-dialog')
    await expect(toTable.locator('input[name="separator"][value="tab"]')).toBeChecked()
    await expect(toTable.locator('input[name="cols"]')).toHaveValue('2')
    await toTable.locator('input[name="autoFit"][value="contents"]').check()
    await toTable.locator('button.primary').click()
    await expect(page.locator('.doc-page table')).toHaveCount(2)

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const xml = await documentXml(docPath)
          const results = [
            ...xml.matchAll(/w:fldCharType="separate"\/><\/w:r>[\s\S]*?<w:t[^>]*>([^<]*)<\/w:t>/g),
          ]
          return {
            tables: (xml.match(/<w:tbl>/g) ?? []).length,
            results: results.map((match) => match[1]),
            formulas: (xml.match(/=SUM\(ABOVE\)|=SUM\(NoSuchBookmark\)/g) ?? []).length,
          }
        },
        { timeout: 30_000 },
      )
      .toEqual({ tables: 2, results: ['10', '1'], formulas: 2 })
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
