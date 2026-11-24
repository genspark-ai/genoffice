/**
 * Insert > Footnote… > Numbering > Restart each page: the first footnote of
 * every page takes the section's start number. The page a mark lands on is
 * only known once the pages are laid out, so the marks are numbered from the
 * layout and laid out again until they hold. Checked through the dialog on
 * the page, in the sheets that print, in the saved file, after reopening and
 * after edits that move a mark to another page.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import {
  copyFileSync,
  mkdtempSync,
  readFileSync,
  realpathSync,
  rmSync,
  writeFileSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-footnote-restart-each-page'
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const PAGE =
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
const SECT_PR = `<w:sectPr>${PAGE}</w:sectPr>`

const noteMark = (id: number) =>
  `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`
/** a paragraph tall enough that four of them, with their notes, fill a Letter page */
const tall = (n: number, extra = '', pPr = '') =>
  `<w:p><w:pPr><w:spacing w:after="2400"/>${pPr}</w:pPr><w:r><w:t xml:space="preserve">Para ${n} body</w:t></w:r>${extra}</w:p>`
const noted = (n: number, pPr = '') => tall(n, noteMark(n + 1), pPr)
const noteRef = (name: string, result: string) =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> NOTEREF ${name} \\h </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  `<w:r><w:t>${result}</w:t></w:r>` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

/** the document with `count` footnotes "note 1" … in word/footnotes.xml */
async function docxWithNotes(body: string, count: number): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await docxWith(body))
  const notes = Array.from(
    { length: count },
    (_, i) =>
      `<w:footnote w:id="${i + 2}"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> note ${i + 1}</w:t></w:r></w:p></w:footnote>`,
  ).join('')
  zip.file(
    'word/footnotes.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:footnotes ${W}>` +
      '<w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
      '<w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
      `${notes}</w:footnotes>`,
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
  return zip.generateAsync({ type: 'nodebuffer' })
}

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

/** the caret at the end of a paragraph */
async function caretAtEnd(page: Page, paragraph: string) {
  const target = page.locator('.doc-page p', { hasText: paragraph }).first()
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
}

async function openDialog(app: ElectronApplication, page: Page, paragraph: string) {
  await caretAtEnd(page, paragraph)
  await clickMenu(app, ['Insert', 'Footnote…'])
  const dialog = page.locator('.footnote-dialog')
  await expect(dialog).toBeVisible()
  return dialog
}

const marks = (page: Page) =>
  page.locator('.doc-page sup.doc-note-ref[data-note-kind="footnote"]').allTextContents()

/** the numbers in front of the notes at the bottom of each page that has notes (the last page's area is not in a page gap) */
const areas = (page: Page) =>
  page.evaluate(() =>
    Array.from(document.querySelectorAll('.page-gap-notes, .page-notes'))
      .map((area) =>
        Array.from(area.querySelectorAll('.page-gap-note, .page-note')).map(
          (row) =>
            `${row.querySelector('sup')?.textContent ?? ''}:${(
              row.querySelector('.page-note-text')?.textContent ??
              (row.textContent ?? '').slice((row.querySelector('sup')?.textContent ?? '').length)
            ).trim()}`,
        ),
      )
      .filter((rows) => rows.length > 0),
  )

/** the same in the sheets that print */
async function printedAreas(page: Page, sheets: number) {
  await page.evaluate(() =>
    (window as unknown as { __openPagePreview?: () => void }).__openPagePreview?.(),
  )
  await expect(page.locator('.pv-page')).toHaveCount(sheets, { timeout: 30_000 })
  const out = await page.evaluate(() =>
    Array.from(document.querySelectorAll('.pv-page')).map((sheet) =>
      Array.from(sheet.querySelectorAll('.pv-footnotes .pv-footnote > sup')).map(
        (sup) => sup.textContent ?? '',
      ),
    ),
  )
  await page.keyboard.press('Escape')
  await expect(page.locator('.pv-page')).toHaveCount(0)
  return out
}

const noteRefs = (page: Page) =>
  page.locator('.doc-page [data-instr-field^="NOTEREF"]').allTextContents()

const edited = (page: Page) =>
  page.evaluate(() =>
    (window as never as { __aidocs: { edited: () => boolean } }).__aidocs.edited(),
  )

const rows = (numbers: number[], first: number) => numbers.map((n, i) => `${n}:note ${first + i}`)
const seq = (...counts: number[]) =>
  counts.flatMap((count) => Array.from({ length: count }, (_, i) => String(i + 1)))

/** a copy for a check in Word, when the run asks for one */
const keep = (docPath: string, name: string) => {
  const dir = process.env.KEEP_FOOTNOTE_DOCX_DIR
  if (dir) copyFileSync(docPath, join(dir, name))
}

test('Restart each page numbers a dozen footnotes page by page: screen, print, file, reopen, edits', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-footnote-each-page-')))
  const docPath = join(dir, 'each-page.docx')
  // a short line, then twelve paragraphs with a note each, four to a page; the last one refers to the sixth note
  const body =
    '<w:p><w:r><w:t>Intro line</w:t></w:r></w:p>' +
    Array.from({ length: 12 }, (_, i) =>
      i === 5
        ? tall(
            6,
            `<w:bookmarkStart w:id="7" w:name="_RefSixth"/>${noteMark(7)}<w:bookmarkEnd w:id="7"/>`,
          )
        : i === 11
          ? tall(
              12,
              `${noteMark(13)}<w:r><w:t xml:space="preserve"> see </w:t></w:r>${noteRef('_RefSixth', '6')}`,
            )
          : noted(i + 1),
    ).join('') +
    SECT_PR
  writeFileSync(docPath, await docxWithNotes(body, 12))
  let session = await openDoc(docPath)
  try {
    let { app, page } = session
    await expect.poll(() => marks(page), { timeout: 30_000 }).toEqual(seq(12))
    await expect.poll(() => noteRefs(page)).toEqual(['6'])
    expect(await edited(page)).toBe(false)

    const dialog = await openDialog(app, page, 'Para 1 body')
    const restart = dialog.locator('select[name="note-restart"]')
    await expect(restart.locator('option')).toHaveText([
      'Continuous',
      'Restart Each Section',
      'Restart Each Page',
    ])
    await restart.selectOption('eachPage')
    await page.screenshot({ path: test.info().outputPath('dialog-each-page.png') })
    await dialog.locator('button[name="apply"]').click()
    await expect(dialog).toBeHidden()

    // three pages of four: every page opens at 1, in the text and under it
    await expect.poll(() => marks(page), { timeout: 30_000 }).toEqual(seq(4, 4, 4))
    await expect
      .poll(() => areas(page))
      .toEqual([rows([1, 2, 3, 4], 1), rows([1, 2, 3, 4], 5), rows([1, 2, 3, 4], 9)])
    // the cross-reference shows the number its note has now
    await expect.poll(() => noteRefs(page)).toEqual(['2'])
    await page.screenshot({ path: test.info().outputPath('each-page.png'), fullPage: true })
    expect(await printedAreas(page, 3)).toEqual([seq(4), seq(4), seq(4)])

    // endnotes have no such choice
    const again = await openDialog(app, page, 'Para 1 body')
    await expect(again.locator('select[name="note-restart"]')).toHaveValue('eachPage')
    await again.locator('input[type="radio"][value="endnote"]').check()
    await expect(again.locator('select[name="note-restart"] option')).toHaveText([
      'Continuous',
      'Restart Each Section',
    ])
    await page.keyboard.press('Escape')
    await expect(again).toBeHidden()

    await saveFromMenu(app)
    await expect
      .poll(() => part(docPath, 'word/settings.xml'), { timeout: 30_000 })
      .toContain('<w:footnotePr><w:numRestart w:val="eachPage"/></w:footnotePr>')
    const saved = await part(docPath, 'word/document.xml')
    expect(saved).toContain(
      '<w:sectPr><w:footnotePr><w:numRestart w:val="eachPage"/></w:footnotePr>',
    )
    expect(saved.match(/<w:footnoteReference /g)).toHaveLength(12)
    // the cached result of the cross-reference is the page's number
    expect(saved).toMatch(
      /NOTEREF _RefSixth \\h <\/w:instrText><\/w:r><w:r><w:fldChar w:fldCharType="separate"\/><\/w:r><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:t[^>]*>2<\/w:t>/,
    )
    keep(docPath, 'each-page-document.docx')

    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    session = await openDoc(docPath)
    ;({ app, page } = session)
    await expect.poll(() => marks(page), { timeout: 30_000 }).toEqual(seq(4, 4, 4))
    await expect
      .poll(() => areas(page))
      .toEqual([rows([1, 2, 3, 4], 1), rows([1, 2, 3, 4], 5), rows([1, 2, 3, 4], 9)])
    await expect.poll(() => noteRefs(page)).toEqual(['2'])
    // numbering the marks by page is not an edit
    expect(await edited(page)).toBe(false)
    expect(await printedAreas(page, 3)).toEqual([seq(4), seq(4), seq(4)])

    // four more lines at the top of the first page push its fourth paragraph, and its mark, to the second
    await caretAtEnd(page, 'Intro line')
    for (let i = 0; i < 4; i++) {
      await page.keyboard.press('Enter')
      await page.keyboard.type('pushed')
    }
    await expect.poll(() => marks(page), { timeout: 30_000 }).toEqual(seq(3, 4, 4, 1))
    await expect
      .poll(() => areas(page))
      .toEqual([rows([1, 2, 3], 1), rows([1, 2, 3, 4], 4), rows([1, 2, 3, 4], 8), rows([1], 12)])
    await expect.poll(() => noteRefs(page)).toEqual(['3'])
    await page.screenshot({ path: test.info().outputPath('pushed.png'), fullPage: true })

    // undoing the edit takes the marks back with it
    await expect(async () => {
      await page.keyboard.press('ControlOrMeta+z')
      expect(await page.locator('.doc-page p').count()).toBe(13)
      expect(await page.locator('.doc-page p', { hasText: 'pushed' }).count()).toBe(0)
    }).toPass({ timeout: 20_000 })
    await expect.poll(() => marks(page), { timeout: 30_000 }).toEqual(seq(4, 4, 4))
    await expect.poll(() => noteRefs(page)).toEqual(['2'])

    // a note added on the second page takes that page's next number, the ones after it move up
    await caretAtEnd(page, 'Para 5 body')
    await page.keyboard.press('Home')
    await clickMenu(app, ['Insert', 'Footnote…'])
    const insert = page.locator('.footnote-dialog')
    await expect(insert).toBeVisible()
    await insert.locator('button.primary').click()
    await expect(insert).toBeHidden()
    const prompt = page.locator('.modal textarea.prompt-textarea')
    await prompt.fill('note added')
    await page.keyboard.press('ControlOrMeta+Enter')
    await expect(prompt).toBeHidden()
    await expect.poll(() => marks(page), { timeout: 30_000 }).toEqual(seq(4, 5, 4))
    await expect.poll(() => noteRefs(page)).toEqual(['3'])

    await saveFromMenu(app)
    await expect
      .poll(async () => (await part(docPath, 'word/footnotes.xml')).includes('note added'), {
        timeout: 30_000,
      })
      .toBe(true)
    await expect.poll(() => marks(page), { timeout: 30_000 }).toEqual(seq(4, 5, 4))
    await expect
      .poll(() => areas(page), { timeout: 30_000 })
      .toEqual([
        rows([1, 2, 3, 4], 1),
        ['1:note added', ...rows([2, 3, 4, 5], 5)],
        rows([1, 2, 3, 4], 9),
      ])

    // back to Continuous: one run of numbers again
    const back = await openDialog(app, page, 'Para 1 body')
    await back.locator('select[name="note-restart"]').selectOption('continuous')
    await back.locator('button[name="apply"]').click()
    await expect(back).toBeHidden()
    await expect.poll(() => marks(page), { timeout: 30_000 }).toEqual(seq(13))
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('This section restarts on each page while the section before it counts on; start value and format follow', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-footnote-each-page-')))
  const docPath = join(dir, 'each-page-section.docx')
  // six notes over two pages, a next-page section break, then six more over two pages
  const body =
    Array.from({ length: 5 }, (_, i) => noted(i + 1)).join('') +
    noted(6, `<w:sectPr>${PAGE}</w:sectPr>`) +
    Array.from({ length: 6 }, (_, i) => noted(i + 7)).join('') +
    SECT_PR
  writeFileSync(docPath, await docxWithNotes(body, 12))
  let session = await openDoc(docPath)
  try {
    let { app, page } = session
    await expect.poll(() => marks(page), { timeout: 30_000 }).toEqual(seq(12))

    const dialog = await openDialog(app, page, 'Para 8 body')
    await dialog.locator('select[name="note-apply-to"]').selectOption('section')
    await dialog.locator('select[name="note-format"]').selectOption('lowerLetter')
    await dialog.locator('input[name="note-start"]').fill('3')
    await dialog.locator('select[name="note-restart"]').selectOption('eachPage')
    await dialog.locator('button[name="apply"]').click()
    await expect(dialog).toBeHidden()

    // pages 1–2 count 1…6; pages 3 and 4 each open at c
    const expected = [...seq(6), 'c', 'd', 'e', 'f', 'c', 'd']
    await expect.poll(() => marks(page), { timeout: 30_000 }).toEqual(expected)
    await expect
      .poll(async () => (await areas(page)).map((rows) => rows.map((row) => row.split(':')[0])))
      .toEqual([seq(4), ['5', '6'], ['c', 'd', 'e', 'f'], ['c', 'd']])
    await page.screenshot({ path: test.info().outputPath('section.png'), fullPage: true })
    expect(await printedAreas(page, 4)).toEqual([
      seq(4),
      ['5', '6'],
      ['c', 'd', 'e', 'f'],
      ['c', 'd'],
    ])

    await saveFromMenu(app)
    const NOTE_PR =
      '<w:footnotePr><w:numFmt w:val="lowerLetter"/><w:numStart w:val="3"/><w:numRestart w:val="eachPage"/></w:footnotePr>'
    await expect
      .poll(() => part(docPath, 'word/document.xml'), { timeout: 30_000 })
      .toContain(NOTE_PR)
    // the last section only: neither the first one nor the document takes it
    expect((await part(docPath, 'word/document.xml')).split('eachPage')).toHaveLength(2)
    expect(await part(docPath, 'word/settings.xml')).not.toContain('eachPage')
    keep(docPath, 'each-page-section.docx')

    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    session = await openDoc(docPath)
    ;({ app, page } = session)
    await expect.poll(() => marks(page), { timeout: 30_000 }).toEqual(expected)
    expect(await edited(page)).toBe(false)
    const first = await openDialog(app, page, 'Para 2 body')
    await expect(first.locator('select[name="note-restart"]')).toHaveValue('continuous')
    await page.keyboard.press('Escape')
    await expect(first).toBeHidden()
    const second = await openDialog(app, page, 'Para 8 body')
    await expect(second.locator('select[name="note-restart"]')).toHaveValue('eachPage')
    await expect(second.locator('input[name="note-start"]')).toHaveValue('3')
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
