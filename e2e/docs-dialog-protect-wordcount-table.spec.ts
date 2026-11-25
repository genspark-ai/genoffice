/**
 * Dialog controls Word has and these dialogs lacked: Table Properties'
 * Measure in and cell Options, Page Number Format's page numbering choice,
 * Word Count's footnotes and endnotes, Paste Special's source, list and
 * result line, and Protect Document's Sections button.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, rects, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-dialog-protect-wordcount-table'
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const sectPr = (pgNumType = '') =>
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  `${pgNumType}<w:cols w:space="720"/></w:sectPr>`

const cell = (text: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${para(text)}</w:tc>`
const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  `<w:tr>${cell('one')}${cell('two')}${cell('three')}</w:tr>` +
  `<w:tr>${cell('four')}${cell('five')}${cell('six')}</w:tr></w:tbl>`

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

const tableDialog = (page: Page) =>
  page.locator('.table-properties-dialog:not(.table-cell-options)')

test('Table Properties measures a column in percent and keeps cell options behind Options, Page Number Format restarts numbering', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-parity-m-')))
  const docPath = join(dir, 'table.docx')
  writeFileSync(
    docPath,
    await docxWith(
      TABLE + para('after the table') + sectPr('<w:pgNumType w:chapStyle="1" w:chapSep="emDash"/>'),
    ),
  )
  let { launched, page } = await openDoc(docPath, 'after the table')
  try {
    const { app } = launched
    await page.locator('.doc-page td', { hasText: 'one' }).first().click()
    await clickMenu(app, ['Table', 'Table Properties…'])
    const dialog = tableDialog(page)
    await dialog.getByRole('tab', { name: 'Column' }).click()
    const measureIn = dialog.locator('.table-properties-width select')
    await expect(measureIn).toHaveValue('len')
    // a column always has a width here: the checkbox is not offered
    await expect(dialog.locator('.table-properties-width input[type=checkbox]')).toHaveCount(0)
    await measureIn.selectOption('pct')
    const percent = dialog.locator('.table-properties-pct input')
    await expect(percent).toHaveValue('33.3')
    await percent.fill('50')
    await page.screenshot({ path: test.info().outputPath('column-percent.png') })

    await dialog.getByRole('tab', { name: 'Cell' }).click()
    await dialog.getByRole('button', { name: 'Options…' }).click()
    const options = page.locator('.table-cell-options')
    await expect(options.getByLabel('Fit text')).not.toBeChecked()
    await options.getByLabel('Wrap text').uncheck()
    await options.getByLabel('Same as the whole table').uncheck()
    await options.getByLabel('Top').fill('0.5 cm')
    await page.screenshot({ path: test.info().outputPath('cell-options.png') })
    await options.locator('button.primary').click()
    await expect(options).toBeHidden()
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()

    // half of the table for the first column, the rest shared by the other two
    const widths = async () =>
      (await rects(page, '.doc-page table tr:first-child td')).map((r) => Math.round(r.width))
    await expect
      .poll(async () => {
        const [a, b, c] = await widths()
        return [Math.abs(a - 2 * b) <= 3, Math.abs(b - c) <= 2]
      })
      .toEqual([true, true])

    // cancelling Options leaves the cell as it was
    await clickMenu(app, ['Table', 'Table Properties…'])
    await dialog.getByRole('tab', { name: 'Cell' }).click()
    await expect(dialog.locator('.table-properties-width select')).toHaveValue('pct')
    await dialog.getByRole('button', { name: 'Options…' }).click()
    await expect(options.getByLabel('Wrap text')).not.toBeChecked()
    await options.getByLabel('Wrap text').check()
    await options.getByRole('button', { name: 'Cancel' }).click()
    await dialog.locator('button.primary').click()

    // Page Number Format: the chapter number of the file is shown and stays through other changes
    await page.locator('.doc-page p', { hasText: 'after the table' }).first().click()
    await clickMenu(app, ['Insert', 'Page Numbers…'])
    await page
      .locator('.page-numbers-dialog')
      .getByRole('button', { name: 'Format Page Numbers…' })
      .click()
    const format = page.locator('.pgnum-dialog')
    const chapter = format.getByLabel('Include chapter number')
    await expect(chapter).toBeChecked()
    await expect(format.getByLabel('Use separator')).toHaveValue('emDash')
    await expect(format.getByLabel('Continue from previous section')).toBeChecked()
    await format.locator('input[type=number]').fill('5')
    await expect(format.getByRole('radio', { name: 'Start at' })).toBeChecked()
    await format.locator('.gs-dd-btn').click()
    await format.getByRole('option', { name: 'I, II, III, …', exact: true }).click()
    await page.screenshot({ path: test.info().outputPath('page-number-format.png') })
    await format.locator('button.btn-primary').click()
    await expect(format).toBeHidden()

    await saveFromMenu(app)
    const saved = async () => {
      const xml = await part(docPath, 'word/document.xml')
      const cells = xml.match(/<w:tc>[\s\S]*?<\/w:tc>/g) ?? []
      return [
        cells.map((c) => /<w:tcW [^>]*>/.exec(c)?.[0]).filter((w) => w?.includes('w:type="pct"'))
          .length,
        /<w:tcW w:w="2500" w:type="pct"\/>/.test(cells[0] ?? ''),
        /<w:tcW w:w="2500" w:type="pct"\/>/.test(cells[3] ?? ''),
        /<w:noWrap\/>/.test(cells[0] ?? ''),
        /<w:tcMar>.*<w:top w:w="28[34]" w:type="dxa"\/>/.test(cells[0] ?? ''),
        /<w:noWrap\/>/.test(cells[1] ?? ''),
        /<w:pgNumType\b[^>]*>/.exec(xml)?.[0],
      ]
    }
    await expect
      .poll(saved, { timeout: 30_000 })
      .toEqual([
        2,
        true,
        true,
        true,
        true,
        false,
        '<w:pgNumType w:fmt="upperRoman" w:start="5" w:chapStyle="1" w:chapSep="emDash"/>',
      ])

    // the saved file opens with the same choices
    await closeAndSaveVideo(launched, VIDEO_DIR)
    ;({ launched, page } = await openDoc(docPath, 'after the table'))
    await page.locator('.doc-page td', { hasText: 'one' }).first().click()
    await clickMenu(launched.app, ['Table', 'Table Properties…'])
    const reopened = tableDialog(page)
    await reopened.getByRole('tab', { name: 'Column' }).click()
    await expect(reopened.locator('.table-properties-width select')).toHaveValue('pct')
    await expect(reopened.locator('.table-properties-pct input')).toHaveValue('50')
    await reopened.getByRole('tab', { name: 'Cell' }).click()
    await reopened.getByRole('button', { name: 'Options…' }).click()
    await expect(page.locator('.table-cell-options').getByLabel('Wrap text')).not.toBeChecked()
    await page.keyboard.press('Escape')
    await expect(page.locator('.table-cell-options')).toBeHidden()
    // Escape closed the options only
    await expect(reopened).toBeVisible()
    await reopened.getByRole('button', { name: 'Cancel' }).click()
    await clickMenu(launched.app, ['Insert', 'Page Numbers…'])
    await page
      .locator('.page-numbers-dialog')
      .getByRole('button', { name: 'Format Page Numbers…' })
      .click()
    await expect(
      page.locator('.pgnum-dialog').getByRole('radio', { name: 'Start at' }),
    ).toBeChecked()
    await expect(page.locator('.pgnum-dialog input[type=number]')).toHaveValue('5')
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Word Count adds footnotes on request, Paste Special names its source and result, Protect Document lists Sections', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-parity-m-')))
  const docPath = join(dir, 'notes.docx')
  const zip = await JSZip.loadAsync(
    await docxWith(
      '<w:p><w:r><w:t>Alpha beta gamma</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p>' +
        para('tail') +
        sectPr(),
    ),
  )
  zip.file(
    'word/footnotes.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:footnotes ${W}>` +
      '<w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
      '<w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
      '<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Four more note words</w:t></w:r></w:p></w:footnote>' +
      '</w:footnotes>',
  )
  const rels = 'word/_rels/document.xml.rels'
  zip.file(
    rels,
    (await zip.file(rels)!.async('string')).replace(
      '</Relationships>',
      '<Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>',
    ),
  )
  const types = '[Content_Types].xml'
  zip.file(
    types,
    (await zip.file(types)!.async('string')).replace(
      '</Types>',
      '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>',
    ),
  )
  writeFileSync(docPath, await zip.generateAsync({ type: 'nodebuffer' }))
  const { launched, page } = await openDoc(docPath, 'Alpha beta gamma')
  try {
    const { app } = launched

    // Word Count: Word's row order, the note's words in or out
    await clickMenu(app, ['Tools', 'Word Count…'])
    const count = page.locator('.word-count-dialog')
    const figure = (field: string) => count.locator(`tr[data-stat="${field}"] td`).last()
    await expect(count.locator('tr')).toHaveCount(8)
    expect(
      await count.locator('tr').evaluateAll((rows) => rows.map((r) => r.getAttribute('data-stat'))),
    ).toEqual([
      'pages',
      'words',
      'charsNoSpace',
      'charsWithSpace',
      'nonAsianWords',
      'asianChars',
      'paragraphs',
      'lines',
    ])
    const notes = count.getByLabel('Include footnotes and endnotes')
    await expect(notes).toBeChecked()
    await expect(figure('words')).toHaveText('8')
    await expect(figure('paragraphs')).toHaveText('3')
    await page.screenshot({ path: test.info().outputPath('word-count.png') })
    await notes.uncheck()
    await expect(figure('words')).toHaveText('4')
    await expect(figure('paragraphs')).toHaveText('2')
    await count.getByRole('button', { name: 'Close' }).click()
    // the choice is remembered, as in Word
    await clickMenu(app, ['Tools', 'Word Count…'])
    await expect(notes).not.toBeChecked()
    await notes.check()
    await count.getByRole('button', { name: 'Close' }).click()

    // Paste Special
    await app.evaluate(({ clipboard }) =>
      clipboard.write({
        text: 'pasted words',
        html: '<html><head><meta name="Generator" content="Sample Writer 3"></head><body><b>pasted words</b></body></html>',
      }),
    )
    const tail = page.locator('.doc-page p', { hasText: 'tail' }).first()
    await tail.click()
    await page.keyboard.press('End')
    await clickMenu(app, ['Edit', 'Paste Special…'])
    const paste = page.locator('.paste-special-dialog')
    await expect(paste.locator('.paste-special-source span')).toHaveText('Sample Writer 3')
    // a clipboard that was not copied from a saved document cannot be linked: Paste link is grey
    await expect(paste.getByRole('radio')).toHaveCount(2)
    await expect(paste.getByRole('radio').nth(0)).toBeChecked()
    await expect(paste.getByRole('radio').nth(1)).toBeDisabled()
    await expect(paste.getByLabel('Display as icon')).toHaveCount(0)
    const as = paste.locator('.paste-special-as select')
    await expect(as).toHaveValue('html')
    await expect(paste.locator('.paste-special-result')).toHaveText(
      'Inserts the contents of the Clipboard as HTML format.',
    )
    await as.selectOption('text')
    await expect(paste.locator('.paste-special-result')).toHaveText(
      'Inserts the contents of the Clipboard as text without any formatting.',
    )
    await page.screenshot({ path: test.info().outputPath('paste-special.png') })
    await paste.locator('button.primary').click()
    await expect(paste).toBeHidden()
    await expect(tail).toHaveText('tailpasted words')
    await expect(tail.locator('strong, b')).toHaveCount(0)

    // Protect Document: Sections waits for the Forms mode, the password for the checkbox
    await clickMenu(app, ['Tools', 'Protect Document…'])
    const protect = page.locator('.protect-dialog')
    await expect(protect.getByRole('button', { name: 'Sections…' })).toBeDisabled()
    const password = protect
      .locator('label.fld', { hasText: 'Protection password' })
      .locator('input')
    await expect(password).toBeDisabled()
    // the native inputs sit under their drawn boxes: click the labels
    await protect.locator('label.protect-check', { hasText: 'Restrict editing' }).click()
    await expect(password).toBeEnabled()
    await protect.locator('label.protect-check', { hasText: 'Comments' }).click()
    await page.screenshot({ path: test.info().outputPath('protect.png') })
    await protect.locator('button.btn-primary').click()
    await expect(protect).toBeHidden()

    await saveFromMenu(app)
    await expect
      .poll(
        async () => [
          /<w:t[^>]*>tailpasted words<\/w:t>/.test(await part(docPath, 'word/document.xml')),
          /<w:b\/>/.test(await part(docPath, 'word/document.xml')),
          /<w:documentProtection\b[^>]*w:edit="comments"[^>]*w:enforcement="1"/.test(
            await part(docPath, 'word/settings.xml'),
          ),
        ],
        { timeout: 30_000 },
      )
      .toEqual([true, false, true])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
