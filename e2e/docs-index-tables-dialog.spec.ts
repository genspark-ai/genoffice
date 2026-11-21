/**
 * Insert > Index and Tables: the options of each tab are switches of the
 * table's field. Each one changes the table on the page, is kept by Update
 * Table, is written to the saved file as Word writes it, and shows again
 * when the dialog opens on the table.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-index-tables-dialog'
const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '</w:sectPr>'

const para = (text: string, extra = '') => `<w:p><w:r><w:t>${text}</w:t></w:r>${extra}</w:p>`
const heading = (level: number, text: string) =>
  `<w:p><w:pPr><w:pStyle w:val="Heading${level}"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`
const field = (instr: string, result = '') =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  (result ? `<w:r><w:t>${result}</w:t></w:r>` : '') +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
const caption = (n: number, text: string) =>
  '<w:p><w:r><w:t xml:space="preserve">Figure </w:t></w:r>' +
  field('SEQ Figure \\* ARABIC', String(n)) +
  `<w:r><w:t xml:space="preserve"> ${text}</w:t></w:r></w:p>`
const xe = (term: string) => field(`XE "${term}"`)

async function docx(body: string): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  zip.file(
    'word/_rels/document.xml.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
  )
  const headingStyle = (n: number) =>
    `<w:style w:type="paragraph" w:styleId="Heading${n}"><w:name w:val="heading ${n}"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="${n - 1}"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>`
  zip.file(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:styles ${NS}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>${[1, 2, 3].map(headingStyle).join('')}</w:styles>`,
  )
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body>${body}${SECT_PR}</w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

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

interface Line {
  title: string
  page: string | null
  sep: string | null
  leader: string | null
  link: boolean
}

/** what each generated table line on the page shows */
const tableLines = (page: Page): Promise<Line[]> =>
  page.evaluate(() =>
    Array.from(document.querySelectorAll('main.editor-scroll .doc-page .doc-toc-line')).map(
      (el) => ({
        title: el.querySelector('.doc-toc-title')?.textContent ?? '',
        page: el.querySelector('.doc-toc-page')?.textContent ?? null,
        sep: el.querySelector('.doc-toc-sep')?.textContent ?? null,
        leader: el.querySelector('.doc-toc-dots')?.textContent?.slice(0, 1) ?? null,
        link: !el.hasAttribute('data-toc-nolink'),
      }),
    ),
  )

/** opens the dialog on that tab; `onTable`: the selection is in such a table, so the dialog opens there by itself */
const openDialog = async (app: ElectronApplication, page: Page, tab: string, onTable = false) => {
  await clickMenu(app, ['Insert', 'Index and Tables…'])
  const dialog = page.locator('.index-tables-dialog')
  await expect(dialog).toBeVisible()
  const button = dialog.getByRole('tab', { name: tab, exact: true })
  if (onTable) await expect(button).toHaveClass(/active/)
  else await button.click()
  return dialog
}

/**
 * Puts the selection on the table line with that text, as a click on it does.
 * A line read back from a saved file without a tab is a plain field line.
 */
const clickLine = (page: Page, text: string | RegExp) =>
  page
    .locator('main.editor-scroll .doc-page .doc-protected-field', { hasText: text })
    .locator('.doc-toc-title, .doc-field-text')
    .first()
    .click()

const updateTable = (page: Page, nth: number) =>
  page
    .locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: /^Update Table$/ })
    .nth(nth)
    .click()

test('Table of Contents and Table of Figures options shape the table, survive Update and are saved as field switches', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-index-tables-')))
  const docPath = join(dir, 'tables.docx')
  writeFileSync(
    docPath,
    await docx(
      para('Contents go here') +
        heading(1, 'Alpha') +
        para('about alpha') +
        heading(2, 'Beta') +
        caption(1, 'Overview') +
        heading(3, 'Gamma') +
        caption(2, 'Detail') +
        para('Figures go here'),
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
    const first = page.locator('.doc-page p', { hasText: 'Contents go here' })
    await expect(first).toBeVisible({ timeout: 30_000 })
    await first.click()
    await page.keyboard.press('End')

    // Word's defaults
    let dialog = await openDialog(app, page, 'Table of Contents')
    const levels = dialog.locator('[data-it="levels"]')
    const pageNumbers = dialog.locator('[data-it="page-numbers"]')
    const rightAlign = dialog.locator('[data-it="right-align"]')
    const leader = dialog.locator('[data-it="leader"]')
    const hyperlinks = dialog.locator('[data-it="hyperlinks"]')
    await expect(levels).toHaveValue('3')
    await expect(pageNumbers).toBeChecked()
    await expect(rightAlign).toBeChecked()
    await expect(leader).toHaveValue('dot')
    await expect(hyperlinks).toBeChecked()
    await levels.fill('2')
    await leader.selectOption('hyphen')
    await page.screenshot({ path: test.info().outputPath('toc-options.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    // two levels, a hyphen leader, pages behind the right tab
    await expect
      .poll(async () => (await tableLines(page)).map((l) => [l.title, l.leader, l.sep, l.link]))
      .toEqual([
        ['Alpha', '-', null, true],
        ['Beta', '-', null, true],
      ])
    expect((await tableLines(page)).every((l) => /^\d+$/.test(l.page ?? ''))).toBe(true)
    await page.screenshot({ path: test.info().outputPath('toc-two-levels-hyphen.png') })

    // opened on the table, the dialog shows what it carries; OK rebuilds it in place
    await clickLine(page, 'Alpha')
    dialog = await openDialog(app, page, 'Table of Contents', true)
    await expect(levels).toHaveValue('2')
    await expect(leader).toHaveValue('hyphen')
    await rightAlign.uncheck()
    await expect(leader).toBeDisabled()
    await hyperlinks.uncheck()
    await levels.fill('3')
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect
      .poll(async () => (await tableLines(page)).map((l) => [l.title, l.leader, l.sep, l.link]))
      .toEqual([
        ['Alpha', null, ' ', false],
        ['Beta', null, ' ', false],
        ['Gamma', null, ' ', false],
      ])
    await page.screenshot({ path: test.info().outputPath('toc-page-after-title.png') })

    // Update Table keeps the switches
    await page.locator('.ribbon-tab', { hasText: /^References$/ }).click()
    await updateTable(page, 0)
    await expect
      .poll(async () => (await tableLines(page)).map((l) => [l.title, l.sep, l.link]))
      .toEqual([
        ['Alpha', ' ', false],
        ['Beta', ' ', false],
        ['Gamma', ' ', false],
      ])
    await saveFromMenu(app)
    await expect
      .poll(() => documentXml(docPath), { timeout: 30_000 })
      .toContain(' TOC \\o "1-3" \\z \\u \\p " " ')
    let xml = await documentXml(docPath)
    expect(xml).toMatch(/<w:pStyle w:val="TOC1"\/><w:rPr>/)
    expect(xml).toMatch(
      />Alpha<\/w:t><\/w:r><w:r><w:rPr><w:noProof\/><\/w:rPr><w:t xml:space="preserve"> \d+<\/w:t>/,
    )
    // the entries and their pages are the field's result: Word is not asked to rebuild the table
    expect(xml).not.toContain('w:dirty')

    // no page numbers at all
    await clickLine(page, 'Alpha')
    dialog = await openDialog(app, page, 'Table of Contents', true)
    await expect(rightAlign).not.toBeChecked()
    await expect(hyperlinks).not.toBeChecked()
    await pageNumbers.uncheck()
    await expect(rightAlign).toBeDisabled()
    await dialog.locator('button.primary').click()
    await expect
      .poll(async () => (await tableLines(page)).map((l) => [l.title, l.page, l.leader]))
      .toEqual([
        ['Alpha', null, null],
        ['Beta', null, null],
        ['Gamma', null, null],
      ])

    // Table of Figures: captions without label and number, an underscore leader
    const last = page.locator('.doc-page p', { hasText: 'Figures go here' })
    await last.click()
    await page.keyboard.press('End')
    dialog = await openDialog(app, page, 'Table of Figures')
    await expect(dialog.locator('[data-it="label"]')).toHaveValue('Figure')
    const includeLabel = dialog.locator('[data-it="include-label"]')
    await expect(includeLabel).toBeChecked()
    await includeLabel.uncheck()
    await leader.selectOption('underscore')
    await page.screenshot({ path: test.info().outputPath('figures-options.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    const figures = async () => (await tableLines(page)).slice(3)
    await expect
      .poll(async () => (await figures()).map((l) => [l.title, l.leader, l.link]))
      .toEqual([
        ['Overview', '_', true],
        ['Detail', '_', true],
      ])
    await page.screenshot({ path: test.info().outputPath('figures-without-label.png') })
    await updateTable(page, 1)
    await expect
      .poll(async () => (await figures()).map((l) => [l.title, l.leader]))
      .toEqual([
        ['Overview', '_'],
        ['Detail', '_'],
      ])
    // opened on it again: the label and the unchecked box come back
    await clickLine(page, /^Overview/)
    dialog = await openDialog(app, page, 'Table of Figures', true)
    await expect(includeLabel).not.toBeChecked()
    await expect(leader).toHaveValue('underscore')
    await includeLabel.check()
    await dialog.locator('button.primary').click()
    await expect
      .poll(async () => (await figures()).map((l) => l.title))
      .toEqual(['Figure 1 Overview', 'Figure 2 Detail'])

    await saveFromMenu(app)
    await expect
      .poll(() => documentXml(docPath), { timeout: 30_000 })
      .toContain(' TOC \\h \\z \\c "Figure" ')
    xml = await documentXml(docPath)
    expect(xml).toContain(' TOC \\o "1-3" \\z \\u \\p " " \\n ')
    expect(xml).toContain('<w:tab w:val="right" w:leader="underscore" w:pos="9350"/>')
    // a table without page numbers and a table of figures with them: both complete, neither marked
    expect(xml).toMatch(/>Figure 2 Detail<\/w:t><\/w:r>[\s\S]*?<w:t[^>]*>\d+<\/w:t>/)
    expect(xml).not.toContain('w:dirty')
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a table that arrives with switches keeps them through Update, and the dialog reads them', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-index-tables-')))
  const docPath = join(dir, 'arrives.docx')
  const tocLine = (text: string, parts: { begin?: boolean; end?: boolean }) =>
    '<w:p><w:pPr><w:pStyle w:val="TOC1"/><w:tabs><w:tab w:val="right" w:leader="underscore" w:pos="9350"/></w:tabs></w:pPr>' +
    (parts.begin
      ? '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-1" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>'
      : '') +
    `<w:r><w:t>${text}</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>1</w:t></w:r>` +
    (parts.end ? '<w:r><w:fldChar w:fldCharType="end"/></w:r>' : '') +
    '</w:p>'
  writeFileSync(
    docPath,
    await docx(
      tocLine('Alpha', { begin: true }) +
        tocLine('Delta', { end: true }) +
        heading(1, 'Alpha') +
        heading(2, 'Beta') +
        heading(1, 'Delta') +
        para('The end'),
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
    await expect(page.locator('.doc-page .doc-toc-line')).toHaveCount(2, { timeout: 30_000 })

    await clickLine(page, 'Alpha')
    const dialog = await openDialog(app, page, 'Table of Contents', true)
    await expect(dialog.locator('[data-it="levels"]')).toHaveValue('1')
    await expect(dialog.locator('[data-it="leader"]')).toHaveValue('underscore')
    await expect(dialog.locator('[data-it="hyperlinks"]')).toBeChecked()
    await dialog.getByRole('button', { name: 'Cancel' }).click()

    // a new top-level heading reaches the table; the second-level one stays out, the leader stays
    const end = page.locator('.doc-page p', { hasText: 'The end' })
    await end.click()
    await page.keyboard.press('End')
    await page.locator('.ribbon-tab', { hasText: /^References$/ }).click()
    await updateTable(page, 0)
    await expect
      .poll(async () => (await tableLines(page)).map((l) => [l.title, l.leader]))
      .toEqual([
        ['Alpha', '_'],
        ['Delta', '_'],
      ])
    await saveFromMenu(app)
    await expect
      .poll(() => documentXml(docPath), { timeout: 30_000 })
      .toContain('<w:tab w:val="right" w:leader="underscore" w:pos="9350"/></w:tabs><w:rPr>')
    const xml = await documentXml(docPath)
    expect(xml).toContain(' TOC \\o "1-1" \\h \\z \\u ')
    expect(xml).not.toContain('>Beta</w:t></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:tab/>')
    expect(xml).not.toContain('w:dirty')
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Index and Table of Authorities options: run-in, right aligned pages, every category, Mark Citation', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-index-tables-')))
  const docPath = join(dir, 'index.docx')
  writeFileSync(
    docPath,
    await docx(
      para('Baldwin') +
        para('about apples', xe('Fruit:Apple')) +
        para('about zebras', xe('Zebra')) +
        para('about pears', xe('Fruit:Pear')) +
        para('Index goes here'),
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
    const last = page.locator('.doc-page p', { hasText: 'Index goes here' })
    await expect(last).toBeVisible({ timeout: 30_000 })
    await last.click()
    await page.keyboard.press('End')

    // Index: indented, pages after the entries
    let dialog = await openDialog(app, page, 'Index')
    const runIn = dialog.locator('[data-it="run-in"]')
    const rightAlign = dialog.locator('[data-it="right-align"]')
    const leader = dialog.locator('[data-it="leader"]')
    await expect(dialog.locator('[data-it="indented"]')).toBeChecked()
    await expect(rightAlign).not.toBeChecked()
    await expect(leader).toBeDisabled()
    await expect(dialog.locator('[data-it="columns"]')).toBeEnabled()
    await expect(dialog.locator('[data-it="columns"]')).toHaveValue('1')
    await expect(dialog.locator('[data-it="mark-entry"]')).toBeEnabled()
    await page.screenshot({ path: test.info().outputPath('index-options.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect
      .poll(async () => (await tableLines(page)).map((l) => [l.title, l.sep, l.page]))
      .toEqual([
        ['Fruit', null, null],
        ['Apple', ', ', '1'],
        ['Pear', ', ', '1'],
        ['Zebra', ', ', '1'],
      ])
    await page.screenshot({ path: test.info().outputPath('index-indented.png') })

    // run-in, with the pages right aligned behind a hyphen leader
    await clickLine(page, 'Zebra')
    dialog = await openDialog(app, page, 'Index', true)
    await runIn.check()
    await rightAlign.check()
    await leader.selectOption('hyphen')
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    const runInLines = [
      ['Fruit: Apple, 1; Pear, 1', null, null],
      ['Zebra', '-', '1'],
    ]
    await expect
      .poll(async () => (await tableLines(page)).map((l) => [l.title, l.leader, l.page]))
      .toEqual(runInLines)
    await page.screenshot({ path: test.info().outputPath('index-run-in.png') })
    await page.locator('.ribbon-tab', { hasText: /^References$/ }).click()
    await page
      .locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: /^Update Index$/ })
      .first()
      .click()
    await expect
      .poll(async () => (await tableLines(page)).map((l) => [l.title, l.leader, l.page]))
      .toEqual(runInLines)
    await clickLine(page, 'Zebra')
    dialog = await openDialog(app, page, 'Index', true)
    await expect(runIn).toBeChecked()
    await expect(rightAlign).toBeChecked()
    await expect(leader).toHaveValue('hyphen')
    await dialog.getByRole('button', { name: 'Cancel' }).click()

    // Mark Entry… opens the marking prompt on the selection; the entry reaches the index through Update
    await page.locator('.doc-page p', { hasText: 'Baldwin' }).first().dblclick()
    dialog = await openDialog(app, page, 'Index')
    await dialog.locator('[data-it="mark-entry"]').click()
    await expect(dialog).toBeHidden()
    const prompt = page.locator('.modal', { hasText: 'Mark Index Entry' })
    await expect(prompt.locator('input[name="main"]')).toHaveValue('Baldwin')
    await prompt.locator('input[name="main"]').press('Enter')
    await expect(prompt).toBeHidden()
    await page
      .locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: /^Update Index$/ })
      .first()
      .click()
    await expect
      .poll(async () => (await tableLines(page)).map((l) => [l.title, l.leader, l.page]))
      .toEqual([['Baldwin', '-', '1'], ...runInLines])

    // Table of Authorities: Mark Citation… opens the marking dialog on the selection
    await page.locator('.doc-page p', { hasText: 'Baldwin' }).first().dblclick()
    dialog = await openDialog(app, page, 'Table of Authorities')
    await dialog.locator('[data-it="mark-citation"]').click()
    const citation = page.locator('.mark-citation-dialog')
    await expect(citation.locator('input[name="long"]')).toHaveValue('Baldwin')
    await citation.locator('input[name="long"]').fill('Baldwin v. Alberti')
    await citation.locator('button.primary').click()
    await expect(citation).toBeHidden()

    await last.click()
    await page.keyboard.press('End')
    dialog = await openDialog(app, page, 'Table of Authorities')
    await expect(dialog.locator('[data-it="category"]')).toHaveValue('0')
    await expect(dialog.locator('[data-it="passim"]')).toBeChecked()
    await expect(dialog.locator('[data-it="keep-formatting"]')).toBeChecked()
    await dialog.locator('[data-it="leader"]').selectOption('none')
    await page.screenshot({ path: test.info().outputPath('authorities-options.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    // at the caret, above the index
    const authorities = async () => (await tableLines(page)).slice(0, 2)
    await expect
      .poll(async () => (await authorities()).map((l) => [l.title, l.leader, l.page]))
      .toEqual([
        ['Cases', '', ''],
        ['Baldwin v. Alberti', '', '1'],
      ])

    await saveFromMenu(app)
    await expect.poll(() => documentXml(docPath), { timeout: 30_000 }).toContain(' TOA \\h \\p ')
    const xml = await documentXml(docPath)
    expect(xml).toContain(' INDEX \\r \\e "\t" ')
    expect(xml).toContain('<w:tab w:val="right" w:leader="hyphen" w:pos="4300"/>')
    expect(xml).toContain('<w:tab w:val="right" w:leader="none" w:pos="9350"/>')
    expect(xml).toContain('>Fruit: Apple, 1; Pear, 1</w:t>')
    // the index and the table of authorities carry their pages as results
    expect(xml).toMatch(/>Baldwin v\. Alberti<\/w:t><\/w:r>[\s\S]*?<w:t[^>]*>1<\/w:t>/)
    expect(xml).not.toContain('w:dirty')
    // read back from the saved file, the table of authorities is still no link
    await expect
      .poll(async () => (await authorities()).map((l) => [l.title, l.link]))
      .toEqual([
        ['Cases', false],
        ['Baldwin v. Alberti', false],
      ])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
