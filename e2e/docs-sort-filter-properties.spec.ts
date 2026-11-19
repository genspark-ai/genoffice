/**
 * Sort by up to three keys (a table's rows by its columns, paragraphs by their
 * fields), Mailings > Filter Recipients with several conditions and Sort
 * Records, and the General, Statistics, Contents and Custom tabs of
 * File > Properties.
 */
import { test, expect, type ElectronApplication } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, statSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-sort-filter-properties'

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

const part = async (file: string, name: string) =>
  (await (await JSZip.loadAsync(readFileSync(file))).file(name)?.async('string')) ?? ''

const cell = (text: string, tcPr = '', rPr = '') =>
  `<w:tc>${tcPr}<w:p><w:r>${rPr}<w:t>${text}</w:t></w:r></w:p></w:tc>`

const table = (rows: string[]) =>
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/>' +
  `<w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>${rows.map((row) => `<w:tr>${row}</w:tr>`).join('')}</w:tbl>`

test('Sort orders a table by two columns under its header row and paragraphs by two fields, and refuses merged cells', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-sort-')))
  const docPath = join(dir, 'lists.docx')
  const people = [
    ['Name', 'City', 'Age'],
    ['Grace', 'Paris', '41'],
    ['Linus', 'Paris', '30'],
    ['Alan', 'London', '9'],
    ['Ada', 'London', '36'],
  ]
  writeFileSync(
    docPath,
    await docxWith(
      ['City,Name', 'Paris,Zoe', 'London,Bo', 'Paris,Al', 'London,Ada'].map(para).join('') +
        para('Between') +
        table(
          people.map((row) =>
            // Alan is bold: the formatting has to travel with the row
            row
              .map((text) => cell(text, '', text === 'Alan' ? '<w:rPr><w:b/></w:rPr>' : ''))
              .join(''),
          ),
        ) +
        para('Merged') +
        table([
          cell('wide', '<w:tcPr><w:gridSpan w:val="2"/></w:tcPr>') + cell('c'),
          cell('x') + cell('y') + cell('z'),
        ]) +
        para('End'),
    ),
  )
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const paragraph = (text: string) => page.locator('.doc-page p', { hasText: text }).first()
    await expect(paragraph('City,Name')).toBeVisible({ timeout: 30_000 })
    const sortButton = page.locator('[data-ribbon-body] button[aria-label="Sort"]')
    const sort = page.locator('.sort-dialog')
    const key = (n: number) => sort.locator(`select[name="key${n}"]`)

    // the caret in the table: the keys are its columns, named by the header row once it is one
    const rows = page.locator('.doc-page .doc-table').first().locator('tr')
    await rows.first().locator('td, th').first().click()
    await sortButton.click()
    await expect(key(1).locator('option')).toHaveText(['Column 1', 'Column 2', 'Column 3'])
    await expect(sort.locator('input[name="separator"]').first()).toBeDisabled()
    await sort.locator('input[name="header"][value="header"]').check()
    await expect(key(1).locator('option')).toHaveText(['Name', 'City', 'Age'])
    await expect(key(3)).toBeDisabled()
    await key(1).selectOption({ label: 'City' })
    await key(2).selectOption({ label: 'Age' })
    await sort.locator('select[name="type2"]').selectOption('number')
    await sort.locator('input[name="order2"][value="descending"]').check()
    await page.screenshot({ path: test.info().outputPath('sort-table.png') })
    await sort.locator('button.primary').click()
    await expect(sort).toBeHidden()
    const sorted = ['NameCityAge', 'AdaLondon36', 'AlanLondon9', 'GraceParis41', 'LinusParis30']
    await expect(rows).toHaveText(sorted)

    // the paragraphs selected: commas part them into fields, the first one names them
    await paragraph('City,Name').click()
    await paragraph('London,Ada').click({ modifiers: ['Shift'] })
    await sortButton.click()
    await expect(key(1).locator('option')).toHaveText(['Paragraphs'])
    await sort.locator('input[name="separator"][value="comma"]').check()
    await expect(key(1).locator('option')).toHaveText(['Paragraphs', 'Field 1', 'Field 2'])
    await sort.locator('input[name="header"][value="header"]').check()
    await key(1).selectOption({ label: 'City' })
    await key(2).selectOption({ label: 'Name' })
    await sort.locator('input[name="order2"][value="descending"]').check()
    await page.screenshot({ path: test.info().outputPath('sort-paragraphs.png') })
    await sort.locator('button.primary').click()
    await expect(sort).toBeHidden()
    const lines = page.locator('main.editor-scroll .ProseMirror').first().locator('> p')
    const sortedLines = ['City,Name', 'London,Bo', 'London,Ada', 'Paris,Zoe', 'Paris,Al']
    await expect.poll(async () => (await lines.allTextContents()).slice(0, 5)).toEqual(sortedLines)

    // a table with merged cells is refused before the dialog opens
    await page.locator('.doc-page .doc-table').nth(1).locator('td, th').last().click()
    await sortButton.click()
    await expect(page.locator('.status-msg')).toContainText('merged cells cannot be sorted')
    await expect(sort).toBeHidden()

    // both sorts reach the file, each row with its own formatting
    await saveFromMenu(app)
    const texts = (xml: string) =>
      Array.from(xml.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g), (match) => match[1]).join('')
    const savedRows = async () => {
      const xml = await part(docPath, 'word/document.xml')
      const first = /<w:tbl>[\s\S]*?<\/w:tbl>/.exec(xml)?.[0] ?? ''
      return Array.from(first.matchAll(/<w:tr[ >][\s\S]*?<\/w:tr>/g), (match) => match[0])
    }
    await expect
      .poll(async () => (await savedRows()).map(texts), { timeout: 30_000 })
      .toEqual(sorted)
    const saved = await savedRows()
    expect(saved[2]).toContain('<w:b/>')
    expect(saved[1]).not.toContain('<w:b/>')
    const body = await part(docPath, 'word/document.xml')
    const savedLines = Array.from(
      body.split('<w:tbl>')[0].matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g),
      (m) => texts(m[0]),
    )
    expect(savedLines.slice(0, 5)).toEqual(sortedLines)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Filter Recipients joins conditions with And and Or and Sort Records orders them, in the preview and in the merged document', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-filter-')))
  const docPath = join(dir, 'letter.docx')
  writeFileSync(docPath, await docxWith(para('Dear') + para('Regards')))
  const listPath = join(dir, 'people.csv')
  writeFileSync(
    listPath,
    'Name,City,Age\nAda,London,36\nGrace,New York,41\nAlan,London,9\nLinus,Helsinki,30\nBo,London,36\n',
  )
  const outPath = join(dir, 'merged.docx')
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const greeting = page.locator('.doc-page p', { hasText: 'Dear' }).first()
    await expect(greeting).toBeVisible({ timeout: 30_000 })
    await app.evaluate(
      ({ dialog }, [list, out]) => {
        dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [list] })
        dialog.showSaveDialog = async () => ({ canceled: false, filePath: out })
      },
      [listPath, outPath],
    )
    await page.locator('.ribbon-tab', { hasText: /^Mailings$/ }).click()
    const big = (name: string) =>
      page.locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: name })
    const record = page.locator('[data-ribbon-body] [data-testid="mailings-record-status"]')
    await big('Select Recipients').click()
    await expect(record).toHaveText('Record 1 of 5')

    // the letter names its recipient, and the preview shows who that is
    await greeting.click()
    await page.keyboard.press('End')
    await page.keyboard.type(' ')
    await big('Insert Merge Field').click()
    await page.locator('.rb-field-item', { hasText: /^Name$/ }).click()
    await big('Preview Results').click()
    const preview = page.locator('.doc-page .doc-merge-preview')
    await expect(preview).toHaveText('Ada')

    const filter = page.locator('.filter-recipients-dialog')
    const row = async (n: number, field: string, op: string, value: string, join?: string) => {
      const name = (base: string) => (n === 1 ? base : `${base}${n}`)
      if (join) await filter.locator(`select[name="join${n}"]`).selectOption(join)
      await filter.locator(`select[name="${name('field')}"]`).selectOption(field)
      await filter.locator(`select[name="${name('op')}"]`).selectOption(op)
      await filter.locator(`input[name="${name('value')}"]`).fill(value)
    }

    // London And an age of 36: Ada and Bo
    await big('Filter Recipients').click()
    await expect(filter.locator('select[name="field2"]')).toBeDisabled()
    await row(1, 'City', 'equals', 'london')
    await row(2, 'Age', 'equals', '36', 'and')
    await page.screenshot({ path: test.info().outputPath('filter-and.png') })
    await filter.locator('button.primary').click()
    await expect(record).toHaveText('Record 1 of 2')
    await expect(preview).toHaveText('Ada')
    await page.locator('[data-ribbon-body] button[aria-label="Last"]').click()
    await expect(preview).toHaveText('Bo')
    await page.locator('[data-ribbon-body] button[aria-label="First"]').click()

    // the same two joined by Or: everyone in London, and anyone aged 36; the dialog reopens on what it holds
    await big('Filter Recipients').click()
    await expect(filter.locator('input[name="value2"]')).toHaveValue('36')
    await filter.locator('select[name="join2"]').selectOption('or')
    await row(2, 'Name', 'contains', 'gr', 'or')
    // And binds first: London, or (Grace and 41)
    await row(3, 'Age', 'equals', '41', 'and')
    // a row left without a value is no condition
    await row(4, 'City', 'equals', '', 'and')
    await filter.locator('button.primary').click()
    await expect(record).toHaveText('Record 1 of 4')

    // Sort Records: the oldest first, the two aged 36 in the order of the list
    await big('Filter Recipients').click()
    await filter.getByRole('tab', { name: 'Sort Records' }).click()
    await expect(filter.locator('select[name="sort2"]')).toBeDisabled()
    await filter.locator('select[name="sort1"]').selectOption('Age')
    await filter.locator('input[name="sortOrder1"][value="descending"]').check()
    await page.screenshot({ path: test.info().outputPath('sort-records.png') })
    await filter.locator('button.primary').click()
    await expect(record).toHaveText('Record 1 of 4')
    await expect(preview).toHaveText('Grace')

    // Find Recipient counts in the same order
    const finder = page.locator('.find-recipient-dialog')
    await big('Find Recipient').click()
    await finder.locator('input[name="text"]').fill('alan')
    await finder.locator('button.primary').click()
    await expect(record).toHaveText('Record 4 of 4')
    await expect(preview).toHaveText('Alan')
    await finder.getByRole('button', { name: 'Close' }).click()

    // and so does the merged document
    await big('Finish & Merge').click()
    const panel = page.locator('.modal-mailmerge')
    await panel.locator('input[name="mailmerge-output"]').nth(1).check()
    await panel.locator('button.btn-primary').click()
    await expect(panel.locator('.mailmerge-done')).toBeVisible({ timeout: 30_000 })
    const merged = await part(outPath, 'word/document.xml')
    expect(merged.match(/Ada|Grace|Alan|Linus|Bo\b/g)).toEqual(['Grace', 'Ada', 'Bo', 'Alan'])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

const CORE_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
  'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
  'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Ada Example</dc:creator>' +
  '<cp:lastModifiedBy>Grace Example</cp:lastModifiedBy><cp:revision>3</cp:revision>' +
  '<dcterms:created xsi:type="dcterms:W3CDTF">2026-01-05T09:30:00Z</dcterms:created>' +
  '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-02-06T10:45:00Z</dcterms:modified></cp:coreProperties>'

const APP_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ' +
  'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
  '<TotalTime>12</TotalTime><Application>Microsoft Office Word</Application></Properties>'

async function docxWithProperties(body: string): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await docxWith(body))
  zip.file('docProps/core.xml', CORE_XML)
  zip.file('docProps/app.xml', APP_XML)
  const types = await zip.file('[Content_Types].xml')!.async('string')
  zip.file(
    '[Content_Types].xml',
    types.replace(
      '</Types>',
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>',
    ),
  )
  const rels = await zip.file('_rels/.rels')!.async('string')
  zip.file(
    '_rels/.rels',
    rels.replace(
      '</Relationships>',
      '<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
        '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>',
    ),
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

test('Properties shows the file, its statistics and its headings, and keeps custom properties in docProps/custom.xml', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-properties-')))
  const docPath = join(dir, 'report.docx')
  writeFileSync(
    docPath,
    await docxWithProperties(para('Quarterly plan') + para('alpha beta gamma')),
  )
  const properties = (page: import('@playwright/test').Page) => page.locator('.properties-dialog')
  const value = (dialog: ReturnType<typeof properties>, prop: string) =>
    dialog.locator(`tr[data-prop="${prop}"] td`).last()
  const customRows = (dialog: ReturnType<typeof properties>) =>
    dialog.locator('.properties-custom-list tbody tr')
  const add = async (
    dialog: ReturnType<typeof properties>,
    name: string,
    type: string,
    text: string,
  ) => {
    await dialog.locator('input[name="customName"]').fill(name)
    await dialog.locator('select[name="customType"]').selectOption(type)
    if (type === 'bool') await dialog.locator(`input[name="customValue"][value="${text}"]`).check()
    else await dialog.locator('input[name="customValue"]').fill(text)
    await dialog.locator('button[name="customAdd"]').click()
  }

  let launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const dialog = properties(page)
    const title = page.locator('.doc-page p', { hasText: 'Quarterly plan' }).first()
    await expect(title).toBeVisible({ timeout: 30_000 })
    await expect(page.locator('.status-page')).toContainText(/\d/, { timeout: 30_000 })

    // a heading for Contents
    await title.click()
    await page.locator('.ribbon-tab', { hasText: /^References$/ }).click()
    await page
      .locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: 'Add Text' })
      .click()
    await page.locator('.layout-menu button', { hasText: /^Level 1$/ }).click()
    await expect(page.locator('.doc-page h1')).toHaveText('Quarterly plan')

    // Statistics: what the file says of itself, and the counts of the document as it is now
    await clickMenu(app, ['File', 'Properties…'])
    await expect(dialog.getByRole('tab', { name: 'Summary' })).toHaveAttribute(
      'aria-selected',
      'true',
    )
    await dialog.getByRole('tab', { name: 'Statistics' }).click()
    await expect(value(dialog, 'lastModifiedBy')).toHaveText('Grace Example')
    await expect(value(dialog, 'revision')).toHaveText('3')
    await expect(value(dialog, 'totalTime')).toHaveText('12 minutes')
    await expect(value(dialog, 'created')).toContainText('2026')
    await expect(value(dialog, 'pages')).toHaveText('1')
    await expect(value(dialog, 'paragraphs')).toHaveText('2')
    await expect(value(dialog, 'words')).toHaveText('5')
    await expect(value(dialog, 'charsNoSpace')).toHaveText('27')
    await page.screenshot({ path: test.info().outputPath('properties-statistics.png') })

    // General: the file on disk
    await dialog.getByRole('tab', { name: 'General' }).click()
    await expect(value(dialog, 'name')).toHaveText('report.docx')
    await expect(value(dialog, 'location')).toHaveText(dir)
    await expect(value(dialog, 'size')).toContainText(
      `${statSync(docPath).size.toLocaleString('en-US')} bytes`,
    )
    await expect(value(dialog, 'modified')).toContainText(/\d/)

    await dialog.getByRole('tab', { name: 'Contents' }).click()
    await expect(dialog.locator('.properties-contents li')).toHaveText(['Quarterly plan'])
    await dialog.getByRole('button', { name: 'Cancel' }).click()

    // the counts follow the document
    await page.locator('.doc-page p', { hasText: 'alpha beta gamma' }).first().click()
    await page.keyboard.press('End')
    await page.keyboard.type(' delta epsilon')
    await clickMenu(app, ['File', 'Properties…'])
    await dialog.getByRole('tab', { name: 'Statistics' }).click()
    await expect(value(dialog, 'words')).toHaveText('7')

    // Custom: one property of each type; a number that is not one cannot be added
    await dialog.getByRole('tab', { name: 'Custom' }).click()
    await expect(customRows(dialog)).toHaveCount(0)
    await add(dialog, 'Client', 'text', 'Example Ltd')
    await add(dialog, 'Due', 'date', '2026-11-17')
    await dialog.locator('input[name="customName"]').fill('Copies')
    await dialog.locator('select[name="customType"]').selectOption('number')
    await dialog.locator('input[name="customValue"]').fill('three')
    await expect(dialog.locator('button[name="customAdd"]')).toBeDisabled()
    await add(dialog, 'Copies', 'number', '3')
    await add(dialog, 'Approved', 'bool', 'true')
    await expect(customRows(dialog)).toHaveCount(4)
    await expect(customRows(dialog).nth(3)).toHaveText('ApprovedYesYes or no')
    await page.screenshot({ path: test.info().outputPath('properties-custom.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()

    await saveFromMenu(app)
    await expect
      .poll(() => part(docPath, 'docProps/custom.xml'), { timeout: 30_000 })
      .toContain('name="Approved"')
    const custom = await part(docPath, 'docProps/custom.xml')
    expect(custom).toContain('pid="2" name="Client"><vt:lpwstr>Example Ltd</vt:lpwstr>')
    expect(custom).toContain('pid="3" name="Due"><vt:filetime>2026-11-17T00:00:00Z</vt:filetime>')
    expect(custom).toContain('pid="4" name="Copies"><vt:i4>3</vt:i4>')
    expect(custom).toContain('pid="5" name="Approved"><vt:bool>true</vt:bool>')
    expect(await part(docPath, '[Content_Types].xml')).toContain(
      '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>',
    )
    expect(await part(docPath, '_rels/.rels')).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"',
    )
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
  }

  // the file opened again lists them; one is changed, one deleted, the others are left byte for byte
  const before = await part(docPath, 'docProps/custom.xml')
  const property = (xml: string, name: string) =>
    new RegExp(`<property\\b[^>]*name="${name}"[^>]*>[\\s\\S]*?</property>`).exec(xml)?.[0] ?? ''
  launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const dialog = properties(page)
    await expect(page.locator('.doc-page h1')).toHaveText('Quarterly plan', { timeout: 30_000 })
    await page.locator('.doc-page h1').click()
    await clickMenu(app, ['File', 'Properties…'])
    await dialog.getByRole('tab', { name: 'Custom' }).click()
    await expect(customRows(dialog)).toHaveText([
      'ClientExample LtdText',
      'Due11/17/2026Date',
      'Copies3Number',
      'ApprovedYesYes or no',
    ])
    // a row picked from the list fills the form: Add reads Modify
    await customRows(dialog).nth(2).click()
    await expect(dialog.locator('input[name="customName"]')).toHaveValue('Copies')
    await expect(dialog.locator('button[name="customAdd"]')).toHaveText('Modify')
    await dialog.locator('input[name="customValue"]').fill('5')
    await dialog.locator('button[name="customAdd"]').click()
    await customRows(dialog).nth(0).click()
    await dialog.locator('button[name="customDelete"]').click()
    await expect(customRows(dialog)).toHaveText([
      'Due11/17/2026Date',
      'Copies5Number',
      'ApprovedYesYes or no',
    ])
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await saveFromMenu(app)
    await expect
      .poll(() => part(docPath, 'docProps/custom.xml'), { timeout: 30_000 })
      .toContain('<vt:i4>5</vt:i4>')
    const after = await part(docPath, 'docProps/custom.xml')
    expect(after).not.toContain('name="Client"')
    expect(after).toContain('pid="4" name="Copies"><vt:i4>5</vt:i4>')
    for (const name of ['Due', 'Approved']) {
      expect(property(after, name)).not.toBe('')
      expect(property(after, name)).toBe(property(before, name))
    }
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
