/**
 * Dialog controls that were listed but could not be used: Page Number
 * Format's chapter number, Table Properties' Fit text and Caption's
 * AutoCaption. Each is driven through its dialog and checked on the page and
 * in the saved file.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-greyed-controls-misc'
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '</w:sectPr>'

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

const part = async (docPath: string, name: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file(name)?.async('string') ?? ''

async function openDoc(docPath: string, text: string) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await expect(page.locator('.doc-page', { hasText: text }).first()).toBeVisible({
    timeout: 30_000,
  })
  return { launched, page }
}

const chapterHeading = (text: string) =>
  `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`
const PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

/** a document whose Heading 1 paragraphs are numbered A, B, … by their style */
async function chapterDocx(): Promise<Buffer> {
  const zip = await JSZip.loadAsync(
    await docxWith(
      chapterHeading('Alpha') +
        para('alpha text') +
        PAGE_BREAK +
        chapterHeading('Beta') +
        para('beta text') +
        SECT_PR,
    ),
  )
  zip.file(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:styles ${W}>` +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>' +
      '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:outlineLvl w:val="0"/></w:pPr>' +
      '<w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style></w:styles>',
  )
  zip.file(
    'word/numbering.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:numbering ${W}>` +
      '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>' +
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperLetter"/><w:pStyle w:val="Heading1"/>' +
      '<w:lvlText w:val="Chapter %1"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr></w:lvl>' +
      '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>',
  )
  const types = await zip.file('[Content_Types].xml')!.async('string')
  zip.file(
    '[Content_Types].xml',
    types.replace(
      '</Types>',
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
        '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>',
    ),
  )
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
  zip.file(
    'word/_rels/document.xml.rels',
    rels.replace(
      '</Relationships>',
      '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        '<Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>',
    ),
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function openPageNumberFormat(app: ElectronApplication, page: Page) {
  await clickMenu(app, ['Insert', 'Page Numbers…'])
  await page
    .locator('.page-numbers-dialog')
    .getByRole('button', { name: 'Format Page Numbers…' })
    .click()
  return page.locator('.pgnum-dialog')
}

test('Page Number Format puts the chapter number in front of the page number', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-greyed-')))
  const docPath = join(dir, 'chapters.docx')
  writeFileSync(docPath, await chapterDocx())
  let { launched, page } = await openDoc(docPath, 'beta text')
  const footers = () => page.locator('.page-hf-footer')
  const pgNumType = async () =>
    /<w:pgNumType\b[^>]*>/.exec(await part(docPath, 'word/document.xml'))?.[0] ?? null
  try {
    const { app } = launched
    await page.locator('.doc-page p', { hasText: 'alpha text' }).first().click()
    await clickMenu(app, ['Insert', 'Page Numbers…'])
    const numbers = page.locator('.page-numbers-dialog')
    await numbers.getByLabel('Alignment').selectOption('center')
    await numbers.locator('button.primary').click()
    await expect(footers().nth(0)).toHaveText('1')
    await expect(footers().nth(1)).toHaveText('2')

    let format = await openPageNumberFormat(app, page)
    const chapter = format.getByLabel('Include chapter number')
    await expect(chapter).not.toBeChecked()
    await expect(format.locator('select[name="pgnum-chapter-style"]')).toBeDisabled()
    await expect(format.locator('select[name="pgnum-chapter-sep"]')).toBeDisabled()
    await chapter.check()
    await expect(format.locator('select[name="pgnum-chapter-style"]')).toHaveValue('1')
    await format.locator('select[name="pgnum-chapter-sep"]').selectOption('colon')
    await page.screenshot({ path: test.info().outputPath('chapter-number-dialog.png') })
    await format.locator('button.btn-primary').click()
    await expect(format).toBeHidden()
    // the heading's own number in its list's format, not the whole "Chapter A" label
    await expect(footers().nth(0)).toHaveText('A:1')
    await expect(footers().nth(1)).toHaveText('B:2')
    await page.screenshot({ path: test.info().outputPath('chapter-number-page.png') })

    await saveFromMenu(app)
    await expect
      .poll(pgNumType, { timeout: 30_000 })
      .toBe('<w:pgNumType w:chapStyle="1" w:chapSep="colon"/>')

    // the saved file opens showing it, and the dialog can take it off again
    await closeAndSaveVideo(launched, VIDEO_DIR)
    ;({ launched, page } = await openDoc(docPath, 'beta text'))
    await expect(footers().nth(0)).toHaveText('A:1')
    await expect(footers().nth(1)).toHaveText('B:2')
    await page.locator('.doc-page p', { hasText: 'alpha text' }).first().click()
    format = await openPageNumberFormat(launched.app, page)
    await expect(format.getByLabel('Include chapter number')).toBeChecked()
    await expect(format.locator('select[name="pgnum-chapter-sep"]')).toHaveValue('colon')
    await format.getByLabel('Include chapter number').uncheck()
    await format.locator('button.btn-primary').click()
    await expect(footers().nth(0)).toHaveText('1')
    await expect(footers().nth(1)).toHaveText('2')
    await saveFromMenu(launched.app)
    await expect.poll(pgNumType, { timeout: 30_000 }).toBeNull()
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

const LONG = 'This sentence is a good deal longer than the cell it sits in'
const cell = (text: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${para(text)}</w:tc>`
const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  `<w:tr>${cell(LONG)}${cell('two')}${cell(LONG)}</w:tr></w:tbl>`

/** per cell of the first row: its width, and the height and drawn text extent of its paragraph */
const cellLines = (page: Page) =>
  page.evaluate(() =>
    Array.from(document.querySelectorAll('.doc-page table tr:first-child td')).map((td) => {
      const p = td.querySelector('p')!
      const range = document.createRange()
      range.selectNodeContents(p)
      const text = range.getBoundingClientRect()
      const box = td.getBoundingClientRect()
      return {
        fitted: p.classList.contains('doc-fit-line'),
        height: Math.round(p.getBoundingClientRect().height),
        textWidth: Math.round(text.width),
        inside: text.left >= box.left - 1 && text.right <= box.right + 1,
        cellWidth: Math.round(box.width),
      }
    }),
  )

async function setFitText(app: ElectronApplication, page: Page, cellIndex: number, on: boolean) {
  await page.locator('.doc-page table tr:first-child td').nth(cellIndex).click()
  await clickMenu(app, ['Table', 'Table Properties…'])
  const dialog = page.locator('.table-properties-dialog:not(.table-cell-options)')
  await dialog.getByRole('tab', { name: 'Cell' }).click()
  await dialog.getByRole('button', { name: 'Options…' }).click()
  const options = page.locator('.table-cell-options')
  const fit = options.getByLabel('Fit text')
  await expect(fit).toBeEnabled()
  if (on) {
    await expect(fit).not.toBeChecked()
    await fit.check()
  } else {
    await expect(fit).toBeChecked()
    await fit.uncheck()
  }
  await options.locator('button.primary').click()
  await expect(options).toBeHidden()
  await dialog.locator('button.primary').click()
  await expect(dialog).toBeHidden()
}

test('Table Properties fits the text of a cell into one line as wide as the cell', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-greyed-')))
  const docPath = join(dir, 'fit.docx')
  writeFileSync(docPath, await docxWith(TABLE + para('after the table') + SECT_PR))
  let { launched, page } = await openDoc(docPath, 'after the table')
  const savedCells = async () =>
    ((await part(docPath, 'word/document.xml')).match(/<w:tc>[\s\S]*?<\/w:tc>/g) ?? []).map((c) =>
      c.includes('<w:tcFitText/>'),
    )
  try {
    const { app } = launched
    const before = await cellLines(page)
    // the long sentence wraps, the short word is one line
    expect(before[0].height).toBeGreaterThan(before[1].height * 1.5)
    expect(before.map((c) => c.fitted)).toEqual([false, false, false])

    await setFitText(app, page, 0, true)
    await setFitText(app, page, 1, true)
    await expect
      .poll(async () => {
        const [long, short, untouched] = await cellLines(page)
        return {
          fitted: [long.fitted, short.fitted, untouched.fitted],
          // squeezed onto one line, inside its cell
          longOneLine: long.height === short.height,
          longInside: long.inside && long.textWidth > long.cellWidth * 0.8,
          // spread across the cell instead of sitting at its left
          shortSpread: short.inside && short.textWidth > short.cellWidth * 0.8,
          untouchedWraps: untouched.height > short.height * 1.5,
        }
      })
      .toEqual({
        fitted: [true, true, false],
        longOneLine: true,
        longInside: true,
        shortSpread: true,
        untouchedWraps: true,
      })
    await page.screenshot({ path: test.info().outputPath('fit-text-page.png') })

    await saveFromMenu(app)
    await expect.poll(savedCells, { timeout: 30_000 }).toEqual([true, true, false])

    // the saved file opens fitted, and the option comes off again
    await closeAndSaveVideo(launched, VIDEO_DIR)
    ;({ launched, page } = await openDoc(docPath, 'after the table'))
    await expect
      .poll(async () => {
        const [long, short] = await cellLines(page)
        return [long.fitted, long.height === short.height, long.inside, short.fitted]
      })
      .toEqual([true, true, true, true])
    await setFitText(launched.app, page, 0, false)
    await expect
      .poll(async () => {
        const [long, short] = await cellLines(page)
        return [long.fitted, long.height > short.height * 1.5]
      })
      .toEqual([false, true])
    await saveFromMenu(launched.app)
    await expect.poll(savedCells, { timeout: 30_000 }).toEqual([false, true, false])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('AutoCaption captions a table as it is inserted', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-greyed-')))
  const docPath = join(dir, 'caption.docx')
  writeFileSync(docPath, await docxWith(para('Opening line') + para('') + SECT_PR))
  const { launched, page } = await openDoc(docPath, 'Opening line')
  const insertTable = async () => {
    await clickMenu(launched.app, ['Table', 'Insert', 'Table…'])
    await page.locator('.modal .btn-primary').click()
  }
  /** the caret goes between "Opening " and "line" */
  const caretMidLine = async () => {
    await page
      .locator('.doc-page p', { hasText: /Opening/ })
      .first()
      .click()
    await page.keyboard.press('End')
    for (let i = 0; i < 4; i++) await page.keyboard.press('ArrowLeft')
  }
  /** what the body holds, in order: paragraphs by their text, tables as "table" */
  const bodyOrder = async () => {
    const body = /<w:body>([\s\S]*)<\/w:body>/.exec(await part(docPath, 'word/document.xml'))?.[1]
    return [...(body ?? '').matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>|<w:p[ >][\s\S]*?<\/w:p>/g)]
      .map((m) =>
        m[0].startsWith('<w:tbl')
          ? 'table'
          : (/SEQ (\w+)/.exec(m[0])?.[0] ?? '') +
            [...m[0].matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((t) => t[1]).join(''),
      )
      .filter(Boolean)
  }
  try {
    const { app } = launched
    await page.locator('.doc-page p', { hasText: 'Opening line' }).first().click()
    await clickMenu(app, ['Insert', 'Caption…'])
    const caption = page.locator('.caption-dialog')
    await caption.locator('input[name="caption-text"]').fill('kept')
    await caption.locator('button[name="auto-caption"]').click()
    const auto = page.getByRole('dialog', { name: 'AutoCaption' })
    await expect(auto.locator('input[name="auto-caption-table"]')).not.toBeChecked()
    await auto.locator('input[name="auto-caption-table"]').check()
    await expect(auto.locator('select[name="auto-caption-label"]')).toHaveValue('Table')
    await expect(auto.locator('select[name="auto-caption-position"]')).toHaveValue('above')
    await page.screenshot({ path: test.info().outputPath('auto-caption-dialog.png') })
    await auto.locator('button[name="ok"]').click()
    await expect(auto).toBeHidden()
    // back in Insert Caption, with what was entered there still in place
    await expect(caption.locator('input[name="caption-text"]')).toHaveValue('kept')
    await caption.locator('.btn-ghost').click()
    await expect(caption).toBeHidden()

    // in the middle of a paragraph: the table splits it, and the caption still finds the table
    await caretMidLine()
    await insertTable()
    await expect(page.locator('.doc-page table')).toHaveCount(1)
    await expect(page.locator('.doc-page', { hasText: 'Table 1' })).toHaveCount(1)
    // one undo takes the table and its caption away together
    await clickMenu(app, ['Edit', 'Undo'])
    await expect(page.locator('.doc-page table')).toHaveCount(0)
    await expect(page.locator('.doc-page', { hasText: 'Table 1' })).toHaveCount(0)
    await expect(page.locator('.doc-page p', { hasText: 'Opening line' })).toHaveCount(1)

    await caretMidLine()
    await insertTable()
    await expect(page.locator('.doc-page', { hasText: 'Table 1' })).toHaveCount(1)
    await page.locator('.doc-page p').last().click()
    await insertTable()
    await expect(page.locator('.doc-page table')).toHaveCount(2)
    await expect(page.locator('.doc-page', { hasText: 'Table 2' })).toHaveCount(1)
    await page.screenshot({ path: test.info().outputPath('auto-caption-page.png') })

    await saveFromMenu(app)
    await expect
      .poll(bodyOrder, { timeout: 30_000 })
      .toEqual(['Opening ', 'SEQ TableTable 1', 'table', 'line', 'SEQ TableTable 2', 'table'])

    // switched off, the next table comes without a caption
    await clickMenu(app, ['Insert', 'Caption…'])
    await caption.locator('button[name="auto-caption"]').click()
    await expect(auto.locator('input[name="auto-caption-table"]')).toBeChecked()
    await auto.locator('input[name="auto-caption-table"]').uncheck()
    await auto.locator('button[name="ok"]').click()
    await caption.locator('.btn-ghost').click()
    await expect(caption).toBeHidden()
    await page.locator('.doc-page p').last().click()
    await insertTable()
    await expect(page.locator('.doc-page table')).toHaveCount(3)
    await expect(page.locator('.doc-page', { hasText: 'Table 3' })).toHaveCount(0)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
