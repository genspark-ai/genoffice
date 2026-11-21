/**
 * Insert > Index and Tables: the controls that used to be greyed. Formats and
 * Modify redefine the entry styles, Options names what a table is built from
 * (TOC \t, \u, \f), an index gets letter headings (\h, \a) and AutoMark, a
 * table of authorities keeps a citation's formatting. Each one is driven
 * through the real dialog and checked on the page, in the saved file and
 * after the file is opened again.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-index-tables-greyed'
const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '</w:sectPr>'

const para = (text: string, extra = '') => `<w:p><w:r><w:t>${text}</w:t></w:r>${extra}</w:p>`
const styled = (style: string, text: string) =>
  `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`
const field = (instr: string) =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

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
  const custom = (id: string, name: string) =>
    `<w:style w:type="paragraph" w:customStyle="1" w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/></w:style>`
  zip.file(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:styles ${NS}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>${[1, 2, 3].map(headingStyle).join('')}${custom('Appendix', 'Appendix Title')}${custom('Plate', 'Plate')}</w:styles>`,
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

const part = async (docPath: string, name: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file(name)?.async('string') ?? ''
const documentXml = (docPath: string) => part(docPath, 'word/document.xml')
/** the w:style element of that id in the saved styles part */
const styleXml = async (docPath: string, id: string) =>
  new RegExp(`<w:style\\b[^>]*w:styleId="${id}"[\\s\\S]*?</w:style>`).exec(
    await part(docPath, 'word/styles.xml'),
  )?.[0] ?? ''

interface Line {
  title: string
  page: string | null
  style: string | null
  weight: string
  slant: string
  transform: string
  align: string
}

/** what each generated table line on the page shows, and how its title is drawn */
const tableLines = (page: Page): Promise<Line[]> =>
  page.evaluate(() =>
    Array.from(document.querySelectorAll('main.editor-scroll .doc-page .doc-toc-line')).map(
      (el) => {
        const title = el.querySelector('.doc-toc-title')!
        const css = getComputedStyle(title)
        return {
          title: title.textContent ?? '',
          page: el.querySelector('.doc-toc-page')?.textContent ?? null,
          style: el.closest('[data-style]')?.getAttribute('data-style') ?? null,
          weight: Number(css.fontWeight) >= 600 ? 'bold' : 'normal',
          slant: css.fontStyle,
          transform: css.textTransform,
          align: getComputedStyle(el.closest('.doc-protected') ?? el).textAlign,
        }
      },
    ),
  )
/** every paragraph of the page in one of those styles: its text, style, weight and alignment */
const styledLines = (
  page: Page,
  styles: string[],
): Promise<Array<[string, string | null, string, string]>> =>
  page.evaluate(
    (styles) =>
      Array.from(document.querySelectorAll('main.editor-scroll .doc-page [data-style]'))
        .filter((el) => styles.includes(el.getAttribute('data-style') ?? ''))
        .map((el) => {
          const text = el.querySelector('.doc-toc-title, .doc-field-text') ?? el
          return [
            (text.textContent ?? '').replace(/\u00a0/g, '').trim(),
            el.getAttribute('data-style'),
            Number(getComputedStyle(text).fontWeight) >= 600 ? 'bold' : 'normal',
            getComputedStyle(el).textAlign,
          ]
        }),
    styles,
  )
const titles = async (page: Page) =>
  (await tableLines(page)).map((l) => l.title.replace(/\u00a0/g, '').trim())

const openDialog = async (app: ElectronApplication, page: Page, tab: string, onTable = false) => {
  await clickMenu(app, ['Insert', 'Index and Tables…'])
  const dialog = page.locator('.index-tables-dialog')
  await expect(dialog).toBeVisible()
  const button = dialog.getByRole('tab', { name: tab, exact: true })
  if (onTable) await expect(button).toHaveClass(/active/)
  else await button.click()
  return dialog
}

const clickLine = (page: Page, text: string | RegExp) =>
  page
    .locator('main.editor-scroll .doc-page .doc-protected-field', { hasText: text })
    .locator('.doc-toc-title, .doc-field-text')
    .first()
    .click()

const open = async (docPath: string) => {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  return { launched, app: launched.app, page }
}

test('Table of Contents: Options builds from styles and TC fields, a format and Modify redefine the TOC styles', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-it-greyed-')))
  const docPath = join(dir, 'contents.docx')
  writeFileSync(
    docPath,
    await docx(
      para('Contents go here') +
        styled('Heading1', 'Alpha') +
        styled('Appendix', 'Appendix one') +
        para('about alpha', field('TC "Marked entry" \\l 2')) +
        styled('Heading2', 'Beta') +
        para('The end'),
    ),
  )
  let session = await open(docPath)
  try {
    const { app, page } = session
    const first = page.locator('.doc-page p', { hasText: 'Contents go here' })
    await expect(first).toBeVisible({ timeout: 30_000 })
    await first.click()
    await page.keyboard.press('End')

    const dialog = await openDialog(app, page, 'Table of Contents')
    for (const live of ['formats', 'modify', 'options'])
      await expect(dialog.locator(`[data-it="${live}"]`)).toBeEnabled()
    // Options: the headings the field lists at their own level, nothing else yet
    await dialog.locator('[data-it="options"]').click()
    const options = page.locator('[data-it-dialog="toc-options"]')
    await expect(options).toBeVisible()
    await expect(options.locator('[data-it="from-styles"]')).toBeChecked()
    await expect(options.locator('[data-it="outline-levels"]')).toBeChecked()
    await expect(options.locator('[data-it="entry-fields"]')).not.toBeChecked()
    for (const n of [1, 2, 3])
      await expect(options.locator(`[data-it-style="heading ${n}"] input`)).toHaveValue(String(n))
    await expect(options.locator('[data-it-style="appendix title"] input')).toHaveValue('')
    await options.locator('[data-it-style="appendix title"] input').fill('2')
    await options.locator('[data-it="entry-fields"]').check()
    await page.screenshot({ path: test.info().outputPath('toc-options.png') })
    await options.locator('button.primary').click()
    await expect(options).toBeHidden()
    // a format brings its own styles: Modify is for the template's
    await dialog.locator('[data-it="formats"]').selectOption('classic')
    await expect(dialog.locator('[data-it="modify"]')).toBeDisabled()
    await page.screenshot({ path: test.info().outputPath('toc-format.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()

    await expect.poll(() => titles(page)).toEqual(['Alpha', 'Appendix one', 'Marked entry', 'Beta'])
    // Classic: level 1 bold capitals, level 2 small capitals
    await expect
      .poll(async () =>
        (await tableLines(page)).map((l) => [
          l.style,
          l.weight,
          l.transform,
          /^\d+$/.test(l.page!),
        ]),
      )
      .toEqual([
        ['TOC1', 'bold', 'uppercase', true],
        ['TOC2', 'normal', 'none', true],
        ['TOC2', 'normal', 'none', true],
        ['TOC2', 'normal', 'none', true],
      ])
    await page.screenshot({ path: test.info().outputPath('toc-classic.png') })

    await saveFromMenu(app)
    await expect
      .poll(() => documentXml(docPath), { timeout: 30_000 })
      .toContain(' TOC \\o "1-3" \\h \\z \\u \\t "Appendix Title,2" \\f ')
    expect(await styleXml(docPath, 'TOC1')).toMatch(/<w:name w:val="toc 1"\/>[\s\S]*<w:b\/>/)
    expect(await styleXml(docPath, 'TOC1')).toContain('<w:caps/>')
    expect(await styleXml(docPath, 'TOC1')).not.toContain('w:customStyle')
    expect(await styleXml(docPath, 'TOC2')).toContain('<w:smallCaps/>')
    expect(await documentXml(docPath)).toContain(' TC "Marked entry" \\l 2 ')
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
  }

  // the saved file opens showing the same table, and the dialog reads its sources back
  session = await open(docPath)
  try {
    const { app, page } = session
    await expect
      .poll(() => titles(page), { timeout: 30_000 })
      .toEqual(['Alpha', 'Appendix one', 'Marked entry', 'Beta'])
    await expect
      .poll(async () => (await tableLines(page)).map((l) => [l.style, l.weight, l.transform]))
      .toEqual([
        ['TOC1', 'bold', 'uppercase'],
        ['TOC2', 'normal', 'none'],
        ['TOC2', 'normal', 'none'],
        ['TOC2', 'normal', 'none'],
      ])
    await clickLine(page, 'Appendix one')
    const dialog = await openDialog(app, page, 'Table of Contents', true)
    await dialog.locator('[data-it="options"]').click()
    const options = page.locator('[data-it-dialog="toc-options"]')
    await expect(options.locator('[data-it-style="appendix title"] input')).toHaveValue('2')
    await expect(options.locator('[data-it="entry-fields"]')).toBeChecked()
    // the appendix style and the TC fields are dropped again; outline levels too
    await options.locator('[data-it-style="appendix title"] input').fill('')
    await options.locator('[data-it="entry-fields"]').uncheck()
    await options.locator('[data-it="outline-levels"]').uncheck()
    await options.locator('button.primary').click()

    // Modify: the entry styles by name, each opening the Modify Style dialog
    await dialog.locator('[data-it="modify"]').click()
    const list = page.locator('[data-it-dialog="entry-styles"]')
    await expect(list.locator('[data-it="entry-style"] option')).toHaveCount(9)
    await list.locator('[data-it="entry-style"]').selectOption('TOC2')
    await list.locator('[data-it="modify-style"]').click()
    const modify = page.locator('.modal.style-dialog')
    await expect(modify).toBeVisible()
    await modify.locator('label', { hasText: 'Italic' }).locator('input').check()
    await page.screenshot({ path: test.info().outputPath('toc-modify-style.png') })
    await modify.locator('button.btn-primary').click()
    await expect(modify).toBeHidden()
    await list.getByRole('button', { name: 'Close' }).click()
    await expect(dialog).toBeVisible()
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()

    await expect.poll(() => titles(page)).toEqual(['Alpha', 'Beta'])
    await expect
      .poll(async () => (await tableLines(page)).map((l) => [l.style, l.slant]))
      .toEqual([
        ['TOC1', 'normal'],
        ['TOC2', 'italic'],
      ])
    await saveFromMenu(app)
    await expect
      .poll(() => documentXml(docPath), { timeout: 30_000 })
      .toContain(' TOC \\o "1-3" \\h \\z </w:instrText>')
    expect(await documentXml(docPath)).not.toContain('Appendix Title,2')
    await expect.poll(() => styleXml(docPath, 'TOC2'), { timeout: 30_000 }).toContain('<w:i/>')
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Index: AutoMark from a concordance file, a format with letter headings, headings for accented letters; Table of Figures Options', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-it-greyed-')))
  const docPath = join(dir, 'index.docx')
  const concordance = join(dir, 'concordance.docx')
  const row = (find: string, entry: string) =>
    `<w:tr><w:tc><w:p><w:r><w:t>${find}</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>${entry}</w:t></w:r></w:p></w:tc></w:tr>`
  writeFileSync(
    concordance,
    await docx(
      `<w:tbl><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>${row('apple', 'apple')}${row('\u00e9clair', '\u00e9clair')}${row('egg', 'egg')}${row('Zebra', 'Zebra')}</w:tbl>` +
        para(''),
    ),
  )
  writeFileSync(
    docPath,
    await docx(
      para('An apple a day, and one more apple') +
        para('The \u00e9clair recipe') +
        para('An egg salad') +
        para('A Zebra crossing') +
        styled('Plate', 'Plate one') +
        para('Index goes here'),
    ),
  )
  let session = await open(docPath)
  try {
    const { app, page } = session
    const last = page.locator('.doc-page p', { hasText: 'Index goes here' })
    await expect(last).toBeVisible({ timeout: 30_000 })
    await last.click()
    await page.keyboard.press('End')

    let dialog = await openDialog(app, page, 'Index')
    for (const live of ['formats', 'accented', 'modify', 'automark'])
      await expect(dialog.locator(`[data-it="${live}"]`)).toBeEnabled()
    await page.screenshot({ path: test.info().outputPath('index-options.png') })
    // the native picker is stubbed in the main process
    await app.evaluate(({ dialog }, file) => {
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [file] })
    }, concordance)
    await dialog.locator('[data-it="automark"]').click()
    await expect(dialog).toBeHidden()
    // one mark per paragraph that has the text: the second "apple" of the first one is not marked
    await expect(page.locator('.doc-page .doc-xe-mark')).toHaveCount(4)

    await last.click()
    await page.keyboard.press('End')
    dialog = await openDialog(app, page, 'Index')
    await dialog.locator('[data-it="formats"]').selectOption('classic')
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    // the accented letter is filed under its base letter
    await expect
      .poll(() => titles(page))
      .toEqual(['A', 'apple', 'E', '\u00e9clair', 'egg', 'Z', 'Zebra'])
    await expect
      .poll(async () =>
        (await tableLines(page))
          .filter((l) => l.style === 'IndexHeading')
          .map((l) => [l.title, l.weight, l.align]),
      )
      .toEqual([
        ['A', 'bold', 'center'],
        ['E', 'bold', 'center'],
        ['Z', 'bold', 'center'],
      ])
    await page.screenshot({ path: test.info().outputPath('index-classic.png') })

    // Headings for accented letters: a heading of its own, behind the base letter
    await clickLine(page, 'egg')
    dialog = await openDialog(app, page, 'Index', true)
    await expect(dialog.locator('[data-it="accented"]')).not.toBeChecked()
    await dialog.locator('[data-it="accented"]').check()
    await dialog.locator('button.primary').click()
    await expect
      .poll(() => titles(page))
      .toEqual(['A', 'apple', 'E', 'egg', '\u00c9', '\u00e9clair', 'Z', 'Zebra'])
    await page.screenshot({ path: test.info().outputPath('index-accented.png') })

    // Table of Figures: Options lists the paragraphs of a style in place of captions
    const plate = page.locator('.doc-page p', { hasText: 'Plate one' })
    await plate.click()
    await page.keyboard.press('End')
    dialog = await openDialog(app, page, 'Table of Figures')
    await dialog.locator('[data-it="options"]').click()
    const options = page.locator('[data-it-dialog="figures-options"]')
    await expect(options).toBeVisible()
    await expect(options.locator('[data-it="by-style"]')).not.toBeChecked()
    await options.locator('[data-it="by-style"]').check()
    await options.locator('[data-it="style"]').selectOption('Plate')
    await page.screenshot({ path: test.info().outputPath('figures-options.png') })
    await options.locator('button.primary').click()
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect
      .poll(async () =>
        (await tableLines(page)).filter((l) => l.style === 'TableofFigures').map((l) => l.title),
      )
      .toEqual(['Plate one'])

    await saveFromMenu(app)
    await expect
      .poll(() => documentXml(docPath), { timeout: 30_000 })
      .toContain(' INDEX \\h "A" \\a ')
    const xml = await documentXml(docPath)
    for (const term of ['apple', '\u00e9clair', 'egg', 'Zebra'])
      expect(xml).toContain(` XE "${term}" `)
    expect(xml.match(/ XE "apple" /g)).toHaveLength(1)
    expect(xml).toContain('<w:pStyle w:val="IndexHeading"/>')
    expect(xml).toMatch(/ TOC \\h \\z \\c "[^"]*" \\t "Plate,1" /)
    expect(await styleXml(docPath, 'IndexHeading')).toMatch(
      /<w:name w:val="index heading"\/>[\s\S]*<w:jc w:val="center"\/>[\s\S]*<w:b\/>/,
    )
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
  }

  session = await open(docPath)
  try {
    const { app, page } = session
    // an index line without a tab reads back as a paragraph in its Index style: entry, separator, page
    const INDEX_STYLES = ['TableofFigures', 'IndexHeading', 'Index1']
    await expect
      .poll(
        async () => (await styledLines(page, INDEX_STYLES)).map(([text, style]) => [text, style]),
        {
          timeout: 30_000,
        },
      )
      .toEqual([
        ['Plate one', 'TableofFigures'],
        ['A', 'IndexHeading'],
        ['apple, 1', 'Index1'],
        ['E', 'IndexHeading'],
        ['egg, 1', 'Index1'],
        ['\u00c9', 'IndexHeading'],
        ['\u00e9clair, 1', 'Index1'],
        ['Z', 'IndexHeading'],
        ['Zebra, 1', 'Index1'],
      ])
    await expect
      .poll(async () =>
        (await styledLines(page, ['IndexHeading'])).map(([, , weight, align]) => [weight, align]),
      )
      .toEqual([
        ['bold', 'center'],
        ['bold', 'center'],
        ['bold', 'center'],
        ['bold', 'center'],
      ])
    await page.locator('.doc-page [data-style="Index1"]', { hasText: 'egg' }).click()
    const dialog = await openDialog(app, page, 'Index', true)
    await expect(dialog.locator('[data-it="accented"]')).toBeChecked()
    await dialog.getByRole('button', { name: 'Cancel' }).click()
    await clickLine(page, 'Plate one')
    const figures = await openDialog(app, page, 'Table of Figures', true)
    await figures.locator('[data-it="options"]').click()
    const options = page.locator('[data-it-dialog="figures-options"]')
    await expect(options.locator('[data-it="by-style"]')).toBeChecked()
    await expect(options.locator('[data-it="style"]')).toHaveValue('Plate')
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Table of Authorities: Keep original formatting lists a citation as it was marked', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-it-greyed-')))
  const docPath = join(dir, 'authorities.docx')
  writeFileSync(
    docPath,
    await docx(
      '<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Baldwin v. Alberti</w:t></w:r><w:r><w:t xml:space="preserve"> (1961)</w:t></w:r></w:p>' +
        para('Authorities go here'),
    ),
  )
  /** the pieces of the first citation line's title: text and whether it is italic */
  const pieces = (page: Page) =>
    page.evaluate(() => {
      const title = Array.from(
        document.querySelectorAll('main.editor-scroll .doc-page .doc-toc-line .doc-toc-title'),
      ).find((el) => el.textContent?.includes('Baldwin'))
      if (!title) return null
      const spans = Array.from(title.querySelectorAll('span'))
      return (spans.length > 0 ? spans : [title]).map((el) => [
        el.textContent,
        getComputedStyle(el).fontStyle,
      ])
    })
  let session = await open(docPath)
  try {
    const { app, page } = session
    const cited = page.locator('.doc-page p', { hasText: 'Baldwin' }).first()
    await expect(cited).toBeVisible({ timeout: 30_000 })
    await cited.click()
    await page.evaluate((length) => {
      const { editor } = (
        window as never as {
          __aidocs: {
            editor: {
              chain(): {
                focus(): { setTextSelection(r: { from: number; to: number }): { run(): boolean } }
              }
            }
          }
        }
      ).__aidocs
      editor
        .chain()
        .focus()
        .setTextSelection({ from: 1, to: 1 + length })
        .run()
    }, 'Baldwin v. Alberti (1961)'.length)
    let dialog = await openDialog(app, page, 'Table of Authorities')
    await dialog.locator('[data-it="mark-citation"]').click()
    const citation = page.locator('.mark-citation-dialog')
    await expect(citation.locator('input[name="long"]')).toHaveValue('Baldwin v. Alberti (1961)')
    await citation.locator('button.primary').click()
    await expect(citation).toBeHidden()

    const last = page.locator('.doc-page p', { hasText: 'Authorities go here' })
    await last.click()
    await page.keyboard.press('End')
    dialog = await openDialog(app, page, 'Table of Authorities')
    const keep = dialog.locator('[data-it="keep-formatting"]')
    await expect(keep).toBeEnabled()
    await expect(keep).toBeChecked()
    await page.screenshot({ path: test.info().outputPath('authorities-options.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect
      .poll(() => pieces(page))
      .toEqual([
        ['Baldwin v. Alberti', 'italic'],
        [' (1961)', 'normal'],
      ])
    await page.screenshot({ path: test.info().outputPath('authorities-formatted.png') })

    await saveFromMenu(app)
    await expect.poll(() => documentXml(docPath), { timeout: 30_000 }).toContain(' TOA \\h \\p ')
    const xml = await documentXml(docPath)
    // the citation's field code and the table line both carry the italic piece
    expect(xml).toContain(
      '<w:r><w:rPr><w:i/></w:rPr><w:instrText xml:space="preserve">Baldwin v. Alberti</w:instrText></w:r>',
    )
    expect(xml).toContain(
      '<w:r><w:rPr><w:i/><w:noProof/></w:rPr><w:t xml:space="preserve">Baldwin v. Alberti</w:t></w:r>',
    )
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
  }

  session = await open(docPath)
  try {
    const { app, page } = session
    await expect
      .poll(() => pieces(page), { timeout: 30_000 })
      .toEqual([
        ['Baldwin v. Alberti', 'italic'],
        [' (1961)', 'normal'],
      ])
    // without Keep original formatting the table is rebuilt plain, from the reopened marker
    await clickLine(page, 'Baldwin v. Alberti (1961)')
    const dialog = await openDialog(app, page, 'Table of Authorities', true)
    await dialog.locator('[data-it="keep-formatting"]').uncheck()
    await dialog.locator('button.primary').click()
    await expect.poll(() => pieces(page)).toEqual([['Baldwin v. Alberti (1961)', 'normal']])
    await saveFromMenu(app)
    await expect
      .poll(() => documentXml(docPath), { timeout: 30_000 })
      .toContain(' TOA \\h \\p \\f ')
    // and with it again, formatted: the marker kept its pieces through the save
    await clickLine(page, 'Baldwin v. Alberti (1961)')
    const again = await openDialog(app, page, 'Table of Authorities', true)
    await expect(again.locator('[data-it="keep-formatting"]')).not.toBeChecked()
    await again.locator('[data-it="keep-formatting"]').check()
    await again.locator('button.primary').click()
    await expect
      .poll(() => pieces(page))
      .toEqual([
        ['Baldwin v. Alberti', 'italic'],
        [' (1961)', 'normal'],
      ])
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
