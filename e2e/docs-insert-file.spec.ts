/**
 * Insert > File through the menu: the body of another document goes in at
 * the caret with its formatting, table, picture and shape, and the saved file
 * holds what was inserted. A file with bookmarks asks for a range, and the
 * footnotes of what is inserted come along.
 */
import { test, expect } from '@playwright/test'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, pictureRun, saveFromMenu, shapeParagraph } from './docs-objects-fixture'

const cell = (t: string) => `<w:tc><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`

/** the document that is inserted: bold text, a table, an in-line picture, a shape, a closing line */
const otherBody = (): string =>
  '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Imported heading</w:t></w:r></w:p>' +
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
  `<w:tr>${cell('North')}${cell('South')}</w:tr></w:tbl>` +
  `<w:p>${pictureRun(1, 20, 10, 60, 40)}</w:p>` +
  shapeParagraph(2, 40, 10, 80, 40) +
  para('Imported closing')

test('the body of another document is inserted at the caret and saved', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-insert-file-')))
  const docPath = join(dir, 'host.docx')
  const otherPath = join(dir, 'other.docx')
  writeFileSync(docPath, await docxWith(para('Host first') + para('Host last')))
  writeFileSync(otherPath, await docxWith(otherBody()))
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-insert-file',
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const first = page.locator('.doc-page p', { hasText: 'Host first' })
    await expect(first).toBeVisible({ timeout: 30_000 })
    // the caret at the end of the first paragraph
    await page.evaluate(() => {
      const { editor } = (
        window as never as {
          __aidocs: {
            editor: { chain(): { focus(): { setTextSelection(p: number): { run(): boolean } } } }
          }
        }
      ).__aidocs
      editor.chain().focus().setTextSelection(11).run()
    })
    // the native picker is stubbed in the main process
    await app.evaluate(({ dialog }, file) => {
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [file] })
    }, otherPath)
    await app.evaluate(({ Menu }) => {
      const insert = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Insert')
      insert?.submenu?.items.find((item) => item.label === 'File…')?.click()
    })
    // a file without bookmarks has one range to offer; Link to File is left off
    const options = page.locator('.insert-file-range-dialog')
    await expect(options).toBeVisible({ timeout: 15_000 })
    await expect(options.locator('select option')).toHaveText(['Whole document'])
    await expect(options.locator('input[name="link"]')).not.toBeChecked()
    await options.locator('button.primary').click()

    await expect(page.locator('.doc-page', { hasText: 'Imported closing' })).toBeVisible()
    // the formatting, the table and the picture came along
    await expect(
      page.locator('.doc-page strong, .doc-page b', { hasText: 'Imported heading' }),
    ).toHaveCount(1)
    await expect(page.locator('.doc-page td', { hasText: 'South' })).toHaveCount(1)
    await expect(page.locator('.doc-page img')).toHaveCount(1)
    // and the shape
    await expect(page.locator('.doc-page .doc-textbox')).toHaveCount(1)
    await expect(options).toHaveCount(0)
    // in order: the host's first line, what was inserted, the host's last line
    const text = await page.locator('.doc-page').innerText()
    expect(text.indexOf('Host first')).toBeLessThan(text.indexOf('Imported heading'))
    expect(text.indexOf('Imported closing')).toBeLessThan(text.indexOf('Host last'))
    await page.screenshot({ path: test.info().outputPath('inserted.png') })
    // the file was read, not opened: it is not among the recent documents
    const recents = await page.evaluate(() =>
      (
        window as never as { desktop: { getRecentFiles(): Promise<string[]> } }
      ).desktop.getRecentFiles(),
    )
    expect(recents).not.toContain(otherPath)

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const zip = await JSZip.loadAsync(readFileSync(docPath))
          const xml = (await zip.file('word/document.xml')?.async('string')) ?? ''
          return [
            xml.includes('Imported heading'),
            xml.includes('<w:tbl>'),
            xml.includes('South'),
            /<w:drawing>/.test(xml),
            Object.keys(zip.files).some((name) => name.startsWith('word/media/')),
            xml.includes('Imported closing'),
            (xml.match(/<wps:wsp>/g) ?? []).length,
            xml.includes('<a:srgbClr val="4472C4"/>'),
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([true, true, true, true, true, true, 1, true])
  } finally {
    await closeAndSaveVideo(launched, 'docs-insert-file')
    rmSync(dir, { recursive: true, force: true })
  }
})

const FOOTNOTES_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml'
const FOOTNOTES_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes'
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const noteRef = (id: number) =>
  `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`
const text = (t: string) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`
const footnote = (id: number, body: string) =>
  `<w:footnote w:id="${id}"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>` +
  `${text(' ')}<w:r><w:rPr><w:i/></w:rPr><w:t>${body}</w:t></w:r></w:p></w:footnote>`

/** `body` as a document with footnotes 1 and 2 */
async function docxWithFootnotes(body: string, notes: [string, string]): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await docxWith(body))
  zip.file(
    'word/footnotes.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:footnotes ${W_NS}>${footnote(1, notes[0])}${footnote(2, notes[1])}</w:footnotes>`,
  )
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
  zip.file(
    'word/_rels/document.xml.rels',
    rels.replace(
      '</Relationships>',
      `<Relationship Id="rIdNotes" Type="${FOOTNOTES_REL}" Target="footnotes.xml"/></Relationships>`,
    ),
  )
  const types = await zip.file('[Content_Types].xml')!.async('string')
  zip.file(
    '[Content_Types].xml',
    types.replace(
      '</Types>',
      `<Override PartName="/word/footnotes.xml" ContentType="${FOOTNOTES_TYPE}"/></Types>`,
    ),
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

test('a bookmark of the other document is inserted as a range, with the footnote inside it', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-insert-range-')))
  const docPath = join(dir, 'host.docx')
  const otherPath = join(dir, 'other.docx')
  writeFileSync(
    docPath,
    await docxWithFootnotes(`<w:p>${text('Host first')}${noteRef(1)}</w:p>` + para('Host last'), [
      'Host note',
      'Host spare note',
    ]),
  )
  // the bookmark opens inside the second paragraph and closes at the start of the fourth
  writeFileSync(
    otherPath,
    await docxWithFootnotes(
      `<w:p>${text('Outside before')}${noteRef(1)}</w:p>` +
        `<w:p>${text('Skipped lead ')}<w:bookmarkStart w:id="5" w:name="Clause"/>${text('Clause opens')}${noteRef(2)}</w:p>` +
        para('Clause closes') +
        `<w:p><w:bookmarkEnd w:id="5"/>${text('Outside after')}</w:p>` +
        `<w:p><w:bookmarkStart w:id="6" w:name="Tail"/>${text('Tail text')}<w:bookmarkEnd w:id="6"/></w:p>`,
      ['Note outside the range', 'Note inside the range'],
    ),
  )
  let launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-insert-file',
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'Host first' })).toBeVisible({
      timeout: 30_000,
    })
    await page.locator('.doc-page p', { hasText: 'Host last' }).click()
    await page.keyboard.press('Home')
    await app.evaluate(({ dialog }, file) => {
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [file] })
    }, otherPath)
    const insertFile = () =>
      app.evaluate(({ Menu }) => {
        const insert = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Insert')
        insert?.submenu?.items.find((item) => item.label === 'File…')?.click()
      })

    // giving up on the range inserts nothing
    await insertFile()
    const dialog = page.locator('.insert-file-range-dialog')
    await expect(dialog).toBeVisible({ timeout: 15_000 })
    await expect(dialog.locator('select option')).toHaveText(['Whole document', 'Clause', 'Tail'])
    await page.screenshot({ path: test.info().outputPath('range-dialog.png') })
    await dialog.getByRole('button', { name: 'Cancel' }).click()
    await expect(dialog).toHaveCount(0)
    await expect(page.locator('.doc-page')).not.toContainText('Clause')

    await insertFile()
    await expect(dialog).toBeVisible({ timeout: 15_000 })
    await dialog.locator('select').selectOption('Clause')
    await dialog.locator('button.primary').click()
    await expect(page.locator('.doc-page')).toContainText('Clause closes')
    const body = page.locator('.doc-page')
    await expect(body).not.toContainText('Skipped lead')
    await expect(body).not.toContainText('Outside')
    await expect(body).not.toContainText('Tail text')
    // two reference marks now, numbered in the order they stand in
    await expect(page.locator('.doc-page sup.doc-note-ref')).toHaveText(['1', '2'])
    await page.screenshot({ path: test.info().outputPath('range-inserted.png') })

    await saveFromMenu(app)
    const saved = async () => {
      const zip = await JSZip.loadAsync(readFileSync(docPath))
      return {
        body: (await zip.file('word/document.xml')?.async('string')) ?? '',
        notes: (await zip.file('word/footnotes.xml')?.async('string')) ?? '',
      }
    }
    await expect
      .poll(async () => (await saved()).body, { timeout: 30_000 })
      .toContain('Clause closes')
    const { body: xml, notes } = await saved()
    expect(xml).toContain('Clause opens')
    expect(xml).not.toContain('Skipped lead')
    expect(xml).not.toContain('Outside')
    const refs = [...xml.matchAll(/<w:footnoteReference w:id="(\d+)"/g)].map((m) => m[1])
    expect(refs).toHaveLength(2)
    expect(new Set(refs).size).toBe(2)
    // the note of the range is stored under the id its mark names; the one outside did not come
    const carried = new RegExp(
      `<w:footnote w:id="${refs[1]}"[^>]*>(?:(?!</w:footnote>).)*Note inside the range`,
    )
    expect(notes).toMatch(carried)
    expect(notes).toContain('Host note')
    expect(notes).toContain('Host spare note')
    expect(notes).not.toContain('Note outside the range')
    expect(notes).toMatch(/<w:i\/>(?:(?!<\/w:r>).)*Note inside the range/)

    // reopened: the same text, the same two marks, and the note where notes are shown
    await closeAndSaveVideo(launched, 'docs-insert-file')
    launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-insert-file',
      openFile: docPath,
    })
    const reopened = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(reopened.locator('.doc-page')).toContainText('Clause closes', { timeout: 30_000 })
    await expect(reopened.locator('.doc-page sup.doc-note-ref')).toHaveText(['1', '2'])
    await expect(reopened.locator('body')).toContainText('Note inside the range')
    await reopened.screenshot({ path: test.info().outputPath('range-reopened.png') })
  } finally {
    await closeAndSaveVideo(launched, 'docs-insert-file')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a section break of the other document comes along with its page setup', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-insert-file-section-')))
  const docPath = join(dir, 'host.docx')
  const otherPath = join(dir, 'other.docx')
  const margins =
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
  const portrait = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>${margins}</w:sectPr>`
  const landscape = `<w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>${margins}</w:sectPr>`
  writeFileSync(docPath, await docxWith(para('Host first') + para('Host last') + portrait))
  writeFileSync(
    otherPath,
    await docxWith(
      para('Wide part') +
        `<w:p><w:pPr>${landscape}</w:pPr></w:p>` +
        para('Imported tail') +
        portrait,
    ),
  )
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-insert-file',
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'Host first' })).toBeVisible({
      timeout: 30_000,
    })
    await page.evaluate(() => {
      const { editor } = (
        window as never as {
          __aidocs: {
            editor: { chain(): { focus(): { setTextSelection(p: number): { run(): boolean } } } }
          }
        }
      ).__aidocs
      editor.chain().focus().setTextSelection(11).run()
    })
    await app.evaluate(({ dialog }, file) => {
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [file] })
    }, otherPath)
    await app.evaluate(({ Menu }) => {
      const insert = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Insert')
      insert?.submenu?.items.find((item) => item.label === 'File…')?.click()
    })
    const options = page.locator('.insert-file-range-dialog')
    await expect(options).toBeVisible({ timeout: 15_000 })
    await options.locator('button.primary').click()
    await expect(page.locator('.doc-page', { hasText: 'Imported tail' })).toBeVisible()

    // as after Insert > Break: the file is written at once and read again
    const documentXml = async () =>
      (await (
        await JSZip.loadAsync(readFileSync(docPath))
      )
        .file('word/document.xml')
        ?.async('string')) ?? ''
    await expect
      .poll(async () => (await documentXml()).match(/<w:sectPr\b/g)?.length ?? 0, {
        timeout: 30_000,
      })
      .toBe(2)
    const xml = await documentXml()
    const wide = xml.indexOf('w:orient="landscape"')
    expect(wide).toBeGreaterThan(xml.indexOf('Wide part'))
    expect(wide).toBeLessThan(xml.indexOf('Imported tail'))
    // the section the break closes is laid out on its wider paper, what follows on the host's:
    // the text column of the first is the wider one (12960 against 9360 twips)
    const columnOf = async (text: string) =>
      (await page.locator('.doc-page p', { hasText: text }).first().boundingBox())!.width
    await expect
      .poll(async () => (await columnOf('Wide part')) / (await columnOf('Host last')), {
        timeout: 30_000,
      })
      .toBeGreaterThan(1.2)
    await page.screenshot({ path: test.info().outputPath('section-came-along.png') })
  } finally {
    await closeAndSaveVideo(launched, 'docs-insert-file')
    rmSync(dir, { recursive: true, force: true })
  }
})
