/**
 * Insert > Footnote…: Word's Footnote and Endnote dialog. Number format, start
 * number and per-section restart renumber the marks on the page; a custom mark
 * replaces the number; Convert moves notes to the other kind; Below text moves
 * the note area under the last line; This section numbers one section on its
 * own. Each is checked on the page, in the saved file and after reopening it.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-footnote-dialog'
const PAGE =
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
const SECT_PR = `<w:sectPr>${PAGE}</w:sectPr>`
const sectionEnd = (text: string) =>
  `<w:p><w:pPr><w:sectPr>${PAGE}</w:sectPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`

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

const openDialog = async (app: ElectronApplication, page: Page) => {
  await clickMenu(app, ['Insert', 'Footnote…'])
  const dialog = page.locator('.footnote-dialog')
  await expect(dialog).toBeVisible()
  return dialog
}

/** the dialog's Insert, then the note text */
async function insertNote(
  app: ElectronApplication,
  page: Page,
  paragraph: string,
  text: string,
  setup?: (dialog: ReturnType<Page['locator']>) => Promise<void>,
) {
  const target = page.locator('.doc-page p', { hasText: paragraph }).first()
  // the note goes where the caret is; a click can lose to the focus a closing dialog hands back
  await expect(async () => {
    await target.click()
    await page.keyboard.press('End')
    const at = await page.evaluate(() => {
      const { editor } = (
        window as never as {
          __aidocs: {
            editor: { state: { selection: { $from: { parent: { textContent: string } } } } }
          }
        }
      ).__aidocs
      return editor.state.selection.$from.parent.textContent
    })
    expect(at).toContain(paragraph)
  }).toPass({ timeout: 15_000 })
  const dialog = await openDialog(app, page)
  await setup?.(dialog)
  await dialog.locator('button.primary').click()
  await expect(dialog).toBeHidden()
  const prompt = page.locator('.modal textarea.prompt-textarea')
  await prompt.fill(text)
  await page.keyboard.press('ControlOrMeta+Enter')
  await expect(prompt).toBeHidden()
}

const marks = (page: Page, kind = 'footnote') =>
  page.locator(`.doc-page sup.doc-note-ref[data-note-kind="${kind}"]`).allTextContents()

/** the numbers the note areas show next to the note texts */
const areaMarks = (page: Page) =>
  page.evaluate(() =>
    Array.from(document.querySelectorAll('.page-gap-note, .page-note'))
      .map((row) => [
        row.querySelector('sup')?.textContent ?? '',
        (row.querySelector('.page-note-text') ?? row.querySelector('span'))?.textContent ?? '',
      ])
      .sort((a, b) => a[1].localeCompare(b[1])),
  )

async function openDoc(docPath: string) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await expect(page.locator('.doc-page p').first()).toBeVisible({ timeout: 30_000 })
  return { launched, app: launched.app, page }
}

test('number format, start number and per-section restart renumber the notes, save and reopen', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-footnote-dialog-')))
  const docPath = join(dir, 'notes.docx')
  writeFileSync(
    docPath,
    await docxWith(
      para('Alpha first') + sectionEnd('Bravo second') + para('Charlie third') + SECT_PR,
    ),
  )
  let session = await openDoc(docPath)
  try {
    let { app, page } = session
    // what Word offers and the page cannot do is not listed
    const first = await openDialog(app, page)
    await expect(first.locator('h2')).toHaveText('Footnote and Endnote')
    // a package without a compatibility mode is laid out as Word 2010 or older: one column of notes
    await expect(first.locator('select[name="note-columns"]')).toBeDisabled()
    await expect(first.locator('option[value="sectEnd"]')).toBeEnabled()
    await expect(first.locator('option[value="eachPage"]')).toBeEnabled()
    await expect(first.locator('option[value="section"]')).not.toHaveAttribute('disabled', '')
    await expect(first.locator('button[name="convert"]')).toBeDisabled()
    await expect(first.locator('button[name="apply"]')).toBeDisabled()
    await expect(first.locator('select[name="note-format"]')).toHaveValue('decimal')
    await page.screenshot({ path: test.info().outputPath('dialog.png') })
    await page.keyboard.press('Escape')
    await expect(first).toBeHidden()

    await insertNote(app, page, 'Alpha first', 'note alpha')
    await insertNote(app, page, 'Bravo second', 'note bravo')
    await insertNote(app, page, 'Charlie third', 'note charlie')
    await expect.poll(() => marks(page)).toEqual(['1', '2', '3'])

    const dialog = await openDialog(app, page)
    await dialog.locator('select[name="note-format"]').selectOption('lowerLetter')
    await dialog.locator('input[name="note-start"]').fill('3')
    await dialog.locator('select[name="note-restart"]').selectOption('eachSect')
    await dialog.locator('button[name="apply"]').click()
    await expect(dialog).toBeHidden()
    // c, d in the first section; the second starts over at c
    await expect.poll(() => marks(page)).toEqual(['c', 'd', 'c'])
    await expect
      .poll(() => areaMarks(page))
      .toEqual([
        ['c', 'note alpha'],
        ['d', 'note bravo'],
        ['c', 'note charlie'],
      ])
    await page.screenshot({ path: test.info().outputPath('renumbered.png') })

    await saveFromMenu(app)
    const NOTE_PR =
      '<w:footnotePr><w:numFmt w:val="lowerLetter"/><w:numStart w:val="3"/><w:numRestart w:val="eachSect"/>'
    await expect
      .poll(() => part(docPath, 'word/settings.xml'), { timeout: 30_000 })
      .toContain(NOTE_PR)
    const documentXml = await part(docPath, 'word/document.xml')
    // both sections carry it: a section's own footnotePr overrides the document's
    expect(documentXml.split(`<w:sectPr>${NOTE_PR}</w:footnotePr>`)).toHaveLength(3)
    expect(documentXml.match(/<w:footnoteReference /g)).toHaveLength(3)

    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    session = await openDoc(docPath)
    ;({ app, page } = session)
    await expect.poll(() => marks(page), { timeout: 30_000 }).toEqual(['c', 'd', 'c'])
    const reopened = await openDialog(app, page)
    await expect(reopened.locator('select[name="note-format"]')).toHaveValue('lowerLetter')
    await expect(reopened.locator('input[name="note-start"]')).toHaveValue('3')
    await expect(reopened.locator('select[name="note-restart"]')).toHaveValue('eachSect')
    // a note added under these settings takes the next number of its section
    await page.keyboard.press('Escape')
    await expect(reopened).toBeHidden()
    await insertNote(app, page, 'Charlie third', 'note delta')
    await expect.poll(() => marks(page)).toEqual(['c', 'd', 'c', 'd'])
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

/** the caret in a paragraph, then the dialog: it opens on that paragraph's section */
async function openDialogIn(app: ElectronApplication, page: Page, paragraph: string) {
  await expect(async () => {
    await page.locator('.doc-page p', { hasText: paragraph }).first().click()
    const at = await page.evaluate(() => {
      const { editor } = (
        window as never as {
          __aidocs: {
            editor: { state: { selection: { $from: { parent: { textContent: string } } } } }
          }
        }
      ).__aidocs
      return editor.state.selection.$from.parent.textContent
    })
    expect(at).toContain(paragraph)
  }).toPass({ timeout: 15_000 })
  return openDialog(app, page)
}

const numbering = async (dialog: ReturnType<Page['locator']>) => [
  await dialog.locator('select[name="note-format"]').inputValue(),
  await dialog.locator('input[name="note-start"]').inputValue(),
  await dialog.locator('select[name="note-restart"]').inputValue(),
]

test('Apply changes to This section numbers that section on its own, saves and reopens', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-footnote-dialog-')))
  const docPath = join(dir, 'sections.docx')
  writeFileSync(
    docPath,
    await docxWith(
      para('Alpha first') + sectionEnd('Bravo second') + para('Charlie third') + SECT_PR,
    ),
  )
  let session = await openDoc(docPath)
  try {
    let { app, page } = session
    await insertNote(app, page, 'Alpha first', 'note alpha')
    await insertNote(app, page, 'Bravo second', 'note bravo')
    await insertNote(app, page, 'Charlie third', 'note charlie')
    await expect.poll(() => marks(page)).toEqual(['1', '2', '3'])

    // the last section: capitals from 5, counted from its own start
    let dialog = await openDialogIn(app, page, 'Charlie third')
    await dialog.locator('select[name="note-apply-to"]').selectOption('section')
    // the page places every section's notes the same way
    await expect(dialog.locator('select[name="footnote-pos"]')).toBeDisabled()
    await dialog.locator('select[name="note-format"]').selectOption('upperRoman')
    await dialog.locator('input[name="note-start"]').fill('5')
    await dialog.locator('select[name="note-restart"]').selectOption('eachSect')
    await page.screenshot({ path: test.info().outputPath('this-section-dialog.png') })
    await dialog.locator('button[name="apply"]').click()
    await expect(dialog).toBeHidden()
    await expect.poll(() => marks(page)).toEqual(['1', '2', 'V'])
    await expect
      .poll(() => areaMarks(page))
      .toEqual([
        ['1', 'note alpha'],
        ['2', 'note bravo'],
        ['V', 'note charlie'],
      ])
    await page.screenshot({ path: test.info().outputPath('this-section-page.png') })

    // the dialog opens on the numbering of the section the caret is in
    dialog = await openDialogIn(app, page, 'Charlie third')
    expect(await numbering(dialog)).toEqual(['upperRoman', '5', 'eachSect'])
    await page.keyboard.press('Escape')
    await expect(dialog).toBeHidden()
    dialog = await openDialogIn(app, page, 'Alpha first')
    expect(await numbering(dialog)).toEqual(['decimal', '1', 'continuous'])
    // the first section, behind its break paragraph: letters
    await dialog.locator('select[name="note-apply-to"]').selectOption('section')
    await dialog.locator('select[name="note-format"]').selectOption('lowerLetter')
    await dialog.locator('button[name="apply"]').click()
    await expect(dialog).toBeHidden()
    await expect.poll(() => marks(page)).toEqual(['a', 'b', 'V'])

    await saveFromMenu(app)
    const LAST =
      '<w:footnotePr><w:numFmt w:val="upperRoman"/><w:numStart w:val="5"/><w:numRestart w:val="eachSect"/></w:footnotePr>'
    await expect
      .poll(() => part(docPath, 'word/document.xml'), { timeout: 30_000 })
      .toContain(`<w:sectPr>${LAST}<w:pgSz `)
    const documentXml = await part(docPath, 'word/document.xml')
    expect(documentXml).toContain(
      '<w:pPr><w:sectPr><w:footnotePr><w:numFmt w:val="lowerLetter"/></w:footnotePr><w:pgSz ',
    )
    expect(documentXml.match(/<w:footnotePr>/g)).toHaveLength(2)
    // nothing was set for the whole document
    expect(await part(docPath, 'word/settings.xml')).not.toContain('w:numFmt')

    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    session = await openDoc(docPath)
    ;({ app, page } = session)
    await expect.poll(() => marks(page), { timeout: 30_000 }).toEqual(['a', 'b', 'V'])
    await expect
      .poll(() => areaMarks(page))
      .toEqual([
        ['a', 'note alpha'],
        ['b', 'note bravo'],
        ['V', 'note charlie'],
      ])
    dialog = await openDialogIn(app, page, 'Bravo second')
    expect(await numbering(dialog)).toEqual(['lowerLetter', '1', 'continuous'])
    await page.keyboard.press('Escape')
    await expect(dialog).toBeHidden()
    // a note added to a section takes that section's next number in its format
    await insertNote(app, page, 'Charlie third', 'note delta')
    await expect.poll(() => marks(page)).toEqual(['a', 'b', 'V', 'VI'])

    // Whole document afterwards reaches every section
    dialog = await openDialogIn(app, page, 'Alpha first')
    await dialog.locator('select[name="note-format"]').selectOption('chicago')
    await dialog.locator('button[name="apply"]').click()
    await expect(dialog).toBeHidden()
    // the last section still starts over at its own 5
    await expect.poll(() => marks(page)).toEqual(['*', '†', '**', '††'])
    await saveFromMenu(app)
    await expect
      .poll(() => part(docPath, 'word/settings.xml'), { timeout: 30_000 })
      .toContain('<w:footnotePr><w:numFmt w:val="chicago"/>')
    const whole = await part(docPath, 'word/document.xml')
    expect(whole.match(/<w:footnotePr><w:numFmt w:val="chicago"\/>/g)).toHaveLength(2)
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a custom mark replaces the number, and Convert turns footnotes into endnotes', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-footnote-dialog-')))
  const docPath = join(dir, 'marks.docx')
  writeFileSync(docPath, await docxWith(para('Alpha first') + para('Bravo second') + SECT_PR))
  let session = await openDoc(docPath)
  try {
    let { app, page } = session
    await insertNote(app, page, 'Alpha first', 'starred note', async (dialog) => {
      // one section: there is no other to tell it from
      await expect(dialog.locator('option[value="section"]')).toHaveAttribute('disabled', '')
      // Symbol… fills the mark; Escape in it leaves the Footnote dialog open
      await dialog.locator('button[name="symbol"]').click()
      const symbols = page.locator('.symbol-dialog')
      await expect(symbols).toBeVisible()
      await page.keyboard.press('Escape')
      await expect(symbols).toBeHidden()
      await expect(dialog).toBeVisible()
      await dialog.locator('button[name="symbol"]').click()
      await symbols.locator('.symbol-cell', { hasText: '§' }).first().dblclick()
      await expect(symbols).toBeHidden()
      await expect(dialog.locator('input[name="custom-mark"]')).toHaveValue('§')
      await dialog.locator('input[name="custom-mark"]').fill('†')
    })
    await insertNote(app, page, 'Bravo second', 'numbered note')
    // the custom mark takes no number: the next note is 1
    await expect.poll(() => marks(page)).toEqual(['†', '1'])
    await expect(page.locator('.doc-page p', { hasText: 'Bravo second' })).toHaveText(
      'Bravo second1',
    )
    await expect
      .poll(() => areaMarks(page))
      .toEqual([
        ['', '† starred note'],
        ['1', 'numbered note'],
      ])
    await page.screenshot({ path: test.info().outputPath('custom-mark.png') })

    await saveFromMenu(app)
    await expect
      .poll(() => part(docPath, 'word/document.xml'), { timeout: 30_000 })
      .toMatch(
        /<w:footnoteReference w:customMarkFollows="1" w:id="(\d+)"\/><w:t xml:space="preserve">†<\/w:t><\/w:r>/,
      )
    const footnotesXml = await part(docPath, 'word/footnotes.xml')
    expect(footnotesXml).toContain(
      '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t xml:space="preserve">†</w:t></w:r>' +
        '<w:r><w:t xml:space="preserve"> starred note</w:t></w:r>',
    )
    expect(footnotesXml.match(/<w:footnoteRef\/>/g)).toHaveLength(1)

    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    session = await openDoc(docPath)
    ;({ app, page } = session)
    await expect.poll(() => marks(page), { timeout: 30_000 }).toEqual(['†', '1'])

    // Convert…: every footnote becomes an endnote, numbered i, ii in Word's endnote format
    const dialog = await openDialog(app, page)
    await dialog.locator('button[name="convert"]').click()
    const convert = page.locator('.convert-notes-dialog')
    await expect(convert.locator('input[value="toEndnotes"]')).toBeChecked()
    await expect(convert.locator('input[value="toFootnotes"]')).toBeDisabled()
    await expect(convert.locator('input[value="swap"]')).toBeDisabled()
    await page.screenshot({ path: test.info().outputPath('convert.png') })
    await convert.locator('button.primary').click()
    await expect(convert).toBeHidden()
    await page.keyboard.press('Escape')
    await expect(dialog).toBeHidden()
    await expect.poll(() => marks(page)).toEqual([])
    await expect.poll(() => marks(page, 'endnote')).toEqual(['†', 'i'])
    await expect(page.locator('.page-endnotes .page-note')).toHaveCount(2)
    await page.screenshot({ path: test.info().outputPath('converted.png') })

    await saveFromMenu(app)
    await expect
      .poll(() => part(docPath, 'word/document.xml'), { timeout: 30_000 })
      .toMatch(/<w:endnoteReference w:customMarkFollows="1"/)
    const converted = await part(docPath, 'word/document.xml')
    expect(converted).not.toContain('<w:footnoteReference')
    expect(converted.match(/<w:endnoteReference /g)).toHaveLength(2)
    const endnotesXml = await part(docPath, 'word/endnotes.xml')
    expect(endnotesXml).toContain('numbered note')
    expect(endnotesXml).toContain(' starred note')
    expect(endnotesXml.match(/<w:endnoteRef\/>/g)).toHaveLength(1)
    expect(endnotesXml).not.toContain('<w:footnote')
    expect(await part(docPath, 'word/footnotes.xml')).not.toContain('note</w:t>')

    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    session = await openDoc(docPath)
    ;({ app, page } = session)
    await expect.poll(() => marks(page, 'endnote'), { timeout: 30_000 }).toEqual(['†', 'i'])
    await expect(page.locator('.page-endnotes .page-note')).toHaveCount(2)
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Below text moves a page footnote area up under the last line', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-footnote-dialog-')))
  const docPath = join(dir, 'position.docx')
  const pageBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
  writeFileSync(
    docPath,
    await docxWith(para('Alpha first') + pageBreak + para('Bravo second') + SECT_PR),
  )
  let session = await openDoc(docPath)
  try {
    let { app, page } = session
    await insertNote(app, page, 'Alpha first', 'page one note')
    // a page's own note area is built from the file: save and reopen puts the new note in it
    await saveFromMenu(app)
    await expect
      .poll(() => part(docPath, 'word/footnotes.xml'), { timeout: 30_000 })
      .toContain('page one note')
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    session = await openDoc(docPath)
    ;({ app, page } = session)
    const area = page.locator('.page-gap-notes').first()
    await expect(area).toBeVisible({ timeout: 30_000 })
    /** how far below the first page's text its note area starts */
    const drop = () =>
      page.evaluate(() => {
        const text = Array.from(document.querySelectorAll('.doc-page p')).find((p) =>
          p.textContent?.includes('Alpha first'),
        )
        const notes = document.querySelector('.page-gap-notes')
        if (!text || !notes) return -1
        return Math.round(notes.getBoundingClientRect().top - text.getBoundingClientRect().bottom)
      })
    await expect.poll(drop).toBeGreaterThan(400)
    await page.screenshot({ path: test.info().outputPath('bottom-of-page.png') })

    const dialog = await openDialog(app, page)
    await dialog.locator('select[name="footnote-pos"]').selectOption('beneathText')
    await dialog.locator('button[name="apply"]').click()
    await expect(dialog).toBeHidden()
    await expect.poll(drop).toBeLessThan(120)
    await expect.poll(drop).toBeGreaterThan(-1)
    await page.screenshot({ path: test.info().outputPath('below-text.png') })

    await saveFromMenu(app)
    await expect
      .poll(() => part(docPath, 'word/settings.xml'), { timeout: 30_000 })
      .toContain('<w:footnotePr><w:pos w:val="beneathText"/></w:footnotePr>')
    expect(await part(docPath, 'word/document.xml')).toContain(
      '<w:sectPr><w:footnotePr><w:pos w:val="beneathText"/></w:footnotePr>',
    )

    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    session = await openDoc(docPath)
    ;({ app, page } = session)
    await expect(page.locator('.page-gap-notes').first()).toBeVisible({ timeout: 30_000 })
    await expect.poll(drop).toBeLessThan(120)
    const reopened = await openDialog(app, page)
    await expect(reopened.locator('select[name="footnote-pos"]')).toHaveValue('beneathText')
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
