/**
 * Table > Formula…: IF, AND, OR, NOT, DEFINED, comparisons, whole rows and
 * columns, and bookmark names, through the dialog and for formulas a document
 * arrives with. Checked on the page and in the saved file.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-formula-logic'
const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/></w:sectPr>'

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

const documentXml = async (docPath: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file('word/document.xml')?.async('string') ?? ''

/** each formula field of the saved file: its instruction as typed, and the result it caches */
const savedFormulas = async (docPath: string) => {
  const xml = await documentXml(docPath)
  const fields = xml.matchAll(
    /<w:instrText[^>]*>([^<]*)<\/w:instrText>[\s\S]*?w:fldCharType="separate"\/><\/w:r>[\s\S]*?<w:t[^>]*>([^<]*)<\/w:t>/g,
  )
  return Array.from(fields, (match) => [
    match[1]
      .trim()
      .replace(/&gt;/g, '>')
      .replace(/&lt;/g, '<')
      .replace(/&quot;/g, '"'),
    match[2],
  ])
}

const tableRows = (page: Page, index: number) =>
  page
    .locator('.doc-page table')
    .nth(index)
    .locator('tr')
    .evaluateAll((rows) =>
      rows.map((row) =>
        Array.from(row.querySelectorAll('td, th'), (cell) => (cell.textContent ?? '').trim()),
      ),
    )

const cell = (page: Page, table: number, row: number, col: number) =>
  page.locator('.doc-page table').nth(table).locator('tr').nth(row).locator('td, th').nth(col)

const tc = (inner: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`
const row = (...cells: string[]) => `<w:tr>${cells.map(tc).join('')}</w:tr>`
const table = (...rows: string[]) =>
  '<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
  rows.join('') +
  '</w:tbl>'
const bookmarked = (id: number, name: string, text: string) =>
  `<w:p><w:bookmarkStart w:id="${id}" w:name="${name}"/><w:r><w:t>${text}</w:t></w:r><w:bookmarkEnd w:id="${id}"/></w:p>`
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

async function enterFormula(app: ElectronApplication, page: Page, text: string, format = '') {
  await clickMenu(app, ['Table', 'Formula…'])
  const dialog = page.locator('.formula-dialog')
  await dialog.locator('input[name="formula"]').fill(text)
  await dialog.locator('input[name="format"]').fill(format)
  await dialog.locator('button.primary').click()
  await expect(dialog).toBeHidden()
}

test('the Formula dialog takes IF, logic, whole rows and columns and a pasted bookmark, and the results follow the cells', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-formula-logic-')))
  const docPath = join(dir, 'logic.docx')
  writeFileSync(
    docPath,
    await docxWith(
      bookmarked(1, 'Rate', 'Rate 10') +
        table(row(bookmarked(2, 'Prices', '5'), para('7'))) +
        para('between') +
        table(
          row(para('Item'), para('Qty')),
          row(para('Apple'), para('3')),
          row(para('Pear'), para('4')),
          row(para('Total'), para('')),
          row(para('Check'), para('')),
          row(para('Other'), para('')),
        ) +
        para('tail') +
        SECT_PR,
    ),
  )
  let session = await openDoc(docPath)
  let { app, page } = session
  try {
    await expect(page.locator('.doc-page table')).toHaveCount(2, { timeout: 30_000 })

    // Paste bookmark lists the document's bookmarks and puts the name into the formula
    await cell(page, 1, 3, 1).click()
    await clickMenu(app, ['Table', 'Formula…'])
    const dialog = page.locator('.formula-dialog')
    const input = dialog.locator('input[name="formula"]')
    await expect(input).toHaveValue('=SUM(ABOVE)')
    await expect(dialog.locator('select[name="pasteBookmark"] option')).toHaveText([
      '',
      'Prices',
      'Rate',
    ])
    await expect(dialog.locator('select[name="pasteFunction"] option')).toContainText([
      'AND',
      'DEFINED',
      'FALSE',
      'IF',
      'NOT',
      'OR',
      'TRUE',
    ])
    await input.fill('=IF(SUM(B:B)>40,2*')
    await dialog.locator('select[name="pasteBookmark"]').selectOption('Rate')
    await expect(input).toHaveValue('=IF(SUM(B:B)>40,2*Rate')
    await input.fill('=IF(SUM(B:B)>40,2*Rate,Rate)')
    await page.screenshot({ path: test.info().outputPath('formula-logic-dialog.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect.poll(async () => (await tableRows(page, 1))[3]).toEqual(['Total', '10'])

    await cell(page, 1, 4, 1).click()
    await enterFormula(app, page, '=AND(B2<B3,NOT(B4=0))+SUM(2:2)', '0.00')
    await expect.poll(async () => (await tableRows(page, 1))[4]).toEqual(['Check', '4.00'])

    // a bookmarked table's cells, and DEFINED on a name that is not there
    await cell(page, 1, 5, 1).click()
    await enterFormula(app, page, '=SUM(Prices A1:B1)+DEFINED(Nothing)+OR(B2>=3,FALSE)')
    await expect.poll(async () => (await tableRows(page, 1))[5]).toEqual(['Other', '13'])

    // the results follow the cells they read
    await cell(page, 1, 1, 1).click()
    await page.keyboard.press('End')
    await page.keyboard.type('5')
    await expect
      .poll(async () => (await tableRows(page, 1)).map((cells) => cells[1]))
      .toEqual(['Qty', '35', '4', '20', '35.00', '13'])
    await page.screenshot({ path: test.info().outputPath('formula-logic-result.png') })

    await saveFromMenu(app)
    await expect
      .poll(() => savedFormulas(docPath), { timeout: 30_000 })
      .toEqual([
        ['=IF(SUM(B:B)>40,2*Rate,Rate)', '20'],
        ['=AND(B2<B3,NOT(B4=0))+SUM(2:2) \\# "0.00"', '35.00'],
        ['=SUM(Prices A1:B1)+DEFINED(Nothing)+OR(B2>=3,FALSE)', '13'],
      ])
    expect(await documentXml(docPath)).toMatch(/<w:bookmarkStart[^>]*w:name="Rate"/)
    // each result is the evaluator's: no field is left for Word to update on open
    expect(await documentXml(docPath)).not.toContain('w:dirty')

    // reopened: the results, and the dialog on the field shows its formula
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    session = await openDoc(docPath)
    ;({ app, page } = session)
    await expect
      .poll(async () => (await tableRows(page, 1)).map((cells) => cells[1]), { timeout: 30_000 })
      .toEqual(['Qty', '35', '4', '20', '35.00', '13'])
    await cell(page, 1, 3, 1).click()
    await clickMenu(app, ['Table', 'Formula…'])
    const reopened = page.locator('.formula-dialog')
    await expect(reopened.locator('input[name="formula"]')).toHaveValue(
      '=IF(SUM(B:B)>40,2*Rate,Rate)',
    )
    await page.keyboard.press('Escape')
    await expect(reopened).toBeHidden()
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('logic formulas a document arrives with are recomputed on Update Field', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-formula-logic-')))
  const docPath = join(dir, 'stale.docx')
  writeFileSync(
    docPath,
    await docxWith(
      bookmarked(1, 'Bonus', 'Bonus: 100') +
        table(
          row(para('North'), para('4')),
          row(para('South'), para('6')),
          row(para('Gap'), formulaPara('=IF(B1&lt;B2,B2-B1,0)', '99')),
          row(para('Named'), formulaPara('=IF(DEFINED(Bonus),Bonus+SUM(1:2),0)', '98')),
          row(para('Column'), formulaPara('=MAX(B:B)', '97')),
        ) +
        para('tail') +
        SECT_PR,
    ),
  )
  const { launched, app, page } = await openDoc(docPath)
  try {
    // opening recomputes nothing: the cached results show as they were saved
    await expect
      .poll(async () => (await tableRows(page, 0)).map((cells) => cells[1]), { timeout: 30_000 })
      .toEqual(['4', '6', '99', '98', '97'])
    await cell(page, 0, 2, 1).click()
    await page.keyboard.press('F9')
    await expect
      .poll(async () => (await tableRows(page, 0)).map((cells) => cells[1]))
      .toEqual(['4', '6', '2', '110', '110'])

    await saveFromMenu(app)
    await expect
      .poll(() => savedFormulas(docPath), { timeout: 30_000 })
      .toEqual([
        ['=IF(B1<B2,B2-B1,0)', '2'],
        ['=IF(DEFINED(Bonus),Bonus+SUM(1:2),0)', '110'],
        ['=MAX(B:B)', '110'],
      ])
    expect(await documentXml(docPath)).not.toContain('w:dirty')
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
