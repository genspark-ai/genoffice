/**
 * Insert > Footnote… > Footnote layout > Columns: the notes at the bottom of a page stand in
 * columns (w15:footnoteColumns of the section, else as many as it has text columns) and the
 * page reserves the height of the tallest. Checked on the page, in the sheets that print, in
 * the exported PDF, in the saved file and after reopening.
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
import { docxWith, saveFromMenu, withWordPart } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-footnote-columns'
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const PAGE =
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
const SECT_PR = `<w:sectPr>${PAGE}</w:sectPr>`
const TWO_COLUMNS = `<w:sectPr>${PAGE}<w:cols w:num="2" w:space="720"/></w:sectPr>`
/** Word 2013 and later: the mode in which the footnote area has columns */
const SETTINGS =
  `<?xml version="1.0" encoding="UTF-8"?><w:settings ${W}><w:compat>` +
  '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>' +
  '</w:compat></w:settings>'

const noteMark = (id: number) =>
  `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`
const BREAK_BEFORE = '<w:pageBreakBefore/>'
/** a paragraph with a note */
const noted = (n: number, pPr = '') =>
  `<w:p><w:pPr>${pPr}</w:pPr><w:r><w:t xml:space="preserve">Para ${n} body</w:t></w:r>${noteMark(n + 1)}</w:p>`
const noteName = (n: number) => `NOTE${String(n).padStart(2, '0')}`
const names = (from: number, to: number) =>
  Array.from({ length: to - from + 1 }, (_, i) => noteName(from + i))
/** takes two lines across a Letter page, and more in a column */
const LONG =
  'about a matter that takes some words to set out in full and then a few more, said twice over so that it runs on to a second line'

/** the document with `count` footnotes NOTE01 …, `long` of them of several lines, and Word 2013's layout */
async function docxWithNotes(
  body: string,
  count: number,
  long: (n: number) => boolean = () => false,
): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await withWordPart(await docxWith(body), 'settings', SETTINGS))
  const notes = Array.from(
    { length: count },
    (_, i) =>
      `<w:footnote w:id="${i + 2}"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>` +
      `<w:r><w:t xml:space="preserve"> ${noteName(i + 1)}${long(i + 1) ? ` ${LONG}` : ' short'}</w:t></w:r></w:p></w:footnote>`,
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

async function openDialog(app: ElectronApplication, page: Page, paragraph: string) {
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
  await clickMenu(app, ['Insert', 'Footnote…'])
  const dialog = page.locator('.footnote-dialog')
  await expect(dialog).toBeVisible()
  return dialog
}

const edited = (page: Page) =>
  page.evaluate(() =>
    (window as never as { __aidocs: { edited: () => boolean } }).__aidocs.edited(),
  )

interface Area {
  /** the rows sit in the box that lays them out in columns */
  boxed: boolean
  /** height of the notes over the width of the area */
  height: number
  /** where every note starts, as shares of the area's width */
  notes: Array<{ name: string; left: number; top: number }>
  /** how far the notes reach past the area's right edge and the bottom of their box (px) */
  pastRight: number
  pastBottom: number
}

/** the footnote areas matched by `areas`, with their rows `rows` */
const shapes = (page: Page, areas: string, rows: string): Promise<Area[]> =>
  page.evaluate(
    ([areas, rows]) =>
      Array.from(document.querySelectorAll(areas)).map((area) => {
        const a = area.getBoundingClientRect()
        const box = area.querySelector('.note-columns')
        const list = Array.from(area.querySelectorAll(rows))
        const rects = list.flatMap((row) => Array.from(row.getClientRects()))
        const top = Math.min(...rects.map((r) => r.top))
        const bottom = Math.max(...rects.map((r) => r.bottom))
        return {
          boxed: !!box,
          height: (bottom - top) / a.width,
          notes: list.map((row) => {
            const start = (row.querySelector('sup') ?? row).getBoundingClientRect()
            return {
              name: /NOTE\d+/.exec(row.textContent ?? '')?.[0] ?? '',
              left: Math.round(((start.left - a.left) / a.width) * 100) / 100,
              top: Math.round(((start.top - top) / a.width) * 100) / 100,
            }
          }),
          pastRight: Math.max(...rects.map((r) => r.right)) - a.right,
          pastBottom: box ? bottom - box.getBoundingClientRect().bottom : 0,
        }
      }),
    [areas, rows],
  )

const onScreen = (page: Page) => shapes(page, '.page-gap-notes', '.page-gap-note')

/** the same in the sheets that print */
async function printed(page: Page): Promise<Area[]> {
  await page.evaluate(() =>
    (window as unknown as { __openPagePreview?: () => void }).__openPagePreview?.(),
  )
  await expect(page.locator('.pv-page').first()).toBeVisible({ timeout: 30_000 })
  await expect
    .poll(() => page.locator('.pv-footnotes').count(), { timeout: 30_000 })
    .toBeGreaterThan(0)
  const out = await shapes(page, '.pv-footnotes', '.pv-footnote')
  await page.keyboard.press('Escape')
  await expect(page.locator('.pv-page')).toHaveCount(0)
  return out
}

/** the columns the notes of an area start in, left to right, by the notes' names */
const columnsOf = (area: Area): string[][] => {
  const lefts = [...new Set(area.notes.map((n) => n.left))].sort((a, b) => a - b)
  return lefts.map((left) => area.notes.filter((n) => n.left === left).map((n) => n.name))
}
const leftsOf = (area: Area) => [...new Set(area.notes.map((n) => n.left))].sort((a, b) => a - b)

/** the notes stay inside their area: no column past the last, nothing under the box */
const contained = (area: Area) => {
  expect(area.pastRight).toBeLessThanOrEqual(1)
  expect(area.pastBottom).toBeLessThanOrEqual(1.5)
}

interface PdfWord {
  page: number
  text: string
  x: number
  y: number
}

/** the text runs of a PDF with the left edge and the distance of the baseline from the page top (pt) */
async function pdfWords(path: string): Promise<PdfWord[]> {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs')
  const doc = await pdfjs.getDocument({ data: new Uint8Array(readFileSync(path)) }).promise
  const out: PdfWord[] = []
  for (let p = 1; p <= doc.numPages; p++) {
    const pdfPage = await doc.getPage(p)
    const height = pdfPage.getViewport({ scale: 1 }).height
    const content = await pdfPage.getTextContent()
    for (const it of content.items) {
      if (!('transform' in it) || !it.str.trim()) continue
      out.push({ page: p, text: it.str, x: it.transform[4], y: height - it.transform[5] })
    }
  }
  return out
}

async function exportPdf(app: ElectronApplication, page: Page, out: string) {
  await app.evaluate(({ dialog: d, shell }, out) => {
    d.showSaveDialog = async () => ({ canceled: false, filePath: out })
    shell.showItemInFolder = () => {}
  }, out)
  expect(
    await page.evaluate(() =>
      (window as unknown as { __exportPdf: () => Promise<boolean> }).__exportPdf(),
    ),
  ).toBe(true)
  return pdfWords(out)
}

/** a copy for a check in Word, when the run asks for one */
const keep = (docPath: string, name: string) => {
  const dir = process.env.KEEP_FOOTNOTE_DOCX_DIR
  if (dir) copyFileSync(docPath, join(dir, name))
}

// columns of 216pt with half an inch between them: the second starts at 252 / 468 of the area
const SECOND_OF_TWO = 0.54
const OF_THREE = [0.36, 0.72]

test('Columns lays a dozen footnotes out in two and three columns: screen, print, PDF, file, reopen, back to one', async () => {
  test.setTimeout(420_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-footnote-columns-')))
  const docPath = join(dir, 'columns.docx')
  // six notes on each of two pages, then a page without
  const body =
    Array.from({ length: 12 }, (_, i) => noted(i + 1, i === 6 ? BREAK_BEFORE : '')).join('') +
    `<w:p><w:pPr>${BREAK_BEFORE}</w:pPr><w:r><w:t>Last page</w:t></w:r></w:p>` +
    SECT_PR
  writeFileSync(docPath, await docxWithNotes(body, 12))
  let session = await openDoc(docPath)
  try {
    let { app, page } = session
    // one under the other: every note starts at the left edge
    await expect.poll(() => onScreen(page).then((a) => a.length), { timeout: 30_000 }).toBe(2)
    const stacked = await onScreen(page)
    expect(stacked.every((area) => !area.boxed)).toBe(true)
    expect(stacked.map(columnsOf)).toEqual([[names(1, 6)], [names(7, 12)]])
    const stackedSheets = await printed(page)
    expect(stackedSheets.map(columnsOf)).toEqual([[names(1, 6)], [names(7, 12)]])
    await page.screenshot({ path: test.info().outputPath('one-column.png'), fullPage: true })

    const dialog = await openDialog(app, page, 'Para 1 body')
    const columns = dialog.locator('select[name="note-columns"]')
    await expect(columns).toBeEnabled()
    await expect(columns).toHaveValue('match')
    await expect(columns.locator('option')).toHaveText(['Match section layout', '1', '2', '3', '4'])
    await expect(dialog.locator('button[name="apply"]')).toBeDisabled()
    await columns.selectOption('2')
    await page.screenshot({ path: test.info().outputPath('dialog-columns.png') })
    // endnotes have no columns
    await dialog.locator('input[type="radio"][value="endnote"]').check()
    await expect(columns).toBeDisabled()
    await dialog.locator('input[type="radio"][value="footnote"]').check()
    await dialog.locator('button[name="apply"]').click()
    await expect(dialog).toBeHidden()

    // three and three, the second three beside the first and level with them
    const halves = [
      [names(1, 3), names(4, 6)],
      [names(7, 9), names(10, 12)],
    ]
    await expect
      .poll(() => onScreen(page).then((a) => a.map(columnsOf)), { timeout: 30_000 })
      .toEqual(halves)
    for (const [i, area] of (await onScreen(page)).entries()) {
      expect(area.boxed).toBe(true)
      expect(leftsOf(area)).toEqual([0, SECOND_OF_TWO])
      expect(area.notes[3].top).toBe(area.notes[0].top)
      contained(area)
      // as tall as three notes, not six
      expect(area.height).toBeLessThan(stacked[i].height * 0.6)
    }
    await page.screenshot({ path: test.info().outputPath('two-columns.png'), fullPage: true })

    const sheets = await printed(page)
    expect(sheets.map(columnsOf)).toEqual(halves)
    for (const [i, area] of sheets.entries()) {
      expect(area.boxed).toBe(true)
      expect(leftsOf(area)).toEqual([0, SECOND_OF_TWO])
      contained(area)
      expect(area.height).toBeLessThan(stackedSheets[i].height * 0.6)
    }

    await saveFromMenu(app)
    await expect
      .poll(() => part(docPath, 'word/document.xml'), { timeout: 30_000 })
      .toContain('<w15:footnoteColumns w:val="2"/></w:sectPr>')
    const saved = await part(docPath, 'word/document.xml')
    const root = /<w:document\b[^>]*>/.exec(saved)![0]
    expect(root).toContain('xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"')
    expect(/mc:Ignorable="([^"]*)"/.exec(root)?.[1].split(' ')).toContain('w15')
    expect(saved.match(/<w:footnoteReference /g)).toHaveLength(12)
    keep(docPath, 'columns-two.docx')

    // the PDF: the notes start in two columns 252pt apart (216pt and half an inch), under the last body line
    const words = await exportPdf(app, page, join(dir, 'columns.pdf'))
    expect([...new Set(words.map((w) => w.page))]).toEqual([1, 2, 3])
    for (const p of [1, 2]) {
      const starts = words.filter((w) => w.page === p && /NOTE\d\d/.test(w.text))
      expect(starts).toHaveLength(6)
      const xs = [...new Set(starts.map((w) => Math.round(w.x)))].sort((a, b) => a - b)
      expect(xs).toHaveLength(2)
      // the text starts after the note's number, which is wider from 10 on
      expect(Math.abs(xs[1] - xs[0] - 252)).toBeLessThanOrEqual(8)
      const ys = [...new Set(starts.map((w) => Math.round(w.y)))]
      expect(ys).toHaveLength(3)
      const body = words.filter((w) => w.page === p && /Para \d+ body/.test(w.text))
      expect(body).toHaveLength(6)
      expect(Math.max(...body.map((w) => w.y))).toBeLessThan(Math.min(...ys))
      // above the bottom margin of a Letter page
      expect(Math.max(...ys)).toBeLessThanOrEqual(792 - 72 + 2)
    }

    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    session = await openDoc(docPath)
    ;({ app, page } = session)
    await expect
      .poll(() => onScreen(page).then((a) => a.map(columnsOf)), { timeout: 30_000 })
      .toEqual(halves)
    expect(await edited(page)).toBe(false)
    expect((await printed(page)).map(columnsOf)).toEqual(halves)

    // three columns: two notes in each
    const three = await openDialog(app, page, 'Para 1 body')
    await expect(three.locator('select[name="note-columns"]')).toHaveValue('2')
    await three.locator('select[name="note-columns"]').selectOption('3')
    await three.locator('button[name="apply"]').click()
    await expect(three).toBeHidden()
    const thirds = [
      [names(1, 2), names(3, 4), names(5, 6)],
      [names(7, 8), names(9, 10), names(11, 12)],
    ]
    await expect
      .poll(() => onScreen(page).then((a) => a.map(columnsOf)), { timeout: 30_000 })
      .toEqual(thirds)
    for (const area of await onScreen(page)) {
      expect(leftsOf(area)).toEqual([0, ...OF_THREE])
      contained(area)
    }
    expect((await printed(page)).map(columnsOf)).toEqual(thirds)
    await page.screenshot({ path: test.info().outputPath('three-columns.png'), fullPage: true })

    // Match section layout: one column of text, one column of notes; the file states nothing
    const back = await openDialog(app, page, 'Para 1 body')
    await back.locator('select[name="note-columns"]').selectOption('match')
    await back.locator('button[name="apply"]').click()
    await expect(back).toBeHidden()
    await expect
      .poll(() => onScreen(page).then((a) => a.map(columnsOf)), { timeout: 30_000 })
      .toEqual([[names(1, 6)], [names(7, 12)]])
    expect((await onScreen(page)).every((area) => !area.boxed)).toBe(true)
    await saveFromMenu(app)
    await expect
      .poll(() => part(docPath, 'word/document.xml'), { timeout: 30_000 })
      .not.toContain('footnoteColumns')
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a two-column section has its notes in two columns unless it states one; This section takes its own count', async () => {
  test.setTimeout(420_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-footnote-columns-')))
  const docPath = join(dir, 'sections.docx')
  // a page of one column with five notes, a next-page break, then two columns of text with six
  const body =
    Array.from({ length: 4 }, (_, i) => noted(i + 1)).join('') +
    noted(5, `<w:sectPr>${PAGE}</w:sectPr>`) +
    Array.from({ length: 6 }, (_, i) => noted(i + 6)).join('') +
    TWO_COLUMNS
  writeFileSync(docPath, await docxWithNotes(body, 11))
  let session = await openDoc(docPath)
  try {
    let { app, page } = session
    // page 1 (one column of text): notes one under the other; page 2 matches its section
    await expect
      .poll(() => printed(page).then((a) => a.map(columnsOf)), { timeout: 60_000 })
      .toEqual([[names(1, 5)], [names(6, 8), names(9, 11)]])
    const opened = await printed(page)
    expect(opened.map((area) => area.boxed)).toEqual([false, true])
    expect(leftsOf(opened[1])).toEqual([0, SECOND_OF_TWO])
    contained(opened[1])

    // the first section alone takes three columns: 2 + 2 + 1, the least height that holds five
    const first = await openDialog(app, page, 'Para 2 body')
    await expect(first.locator('select[name="note-columns"]')).toHaveValue('match')
    await first.locator('select[name="note-apply-to"]').selectOption('section')
    await first.locator('select[name="note-columns"]').selectOption('3')
    await first.locator('button[name="apply"]').click()
    await expect(first).toBeHidden()
    const firstInThree = [names(1, 2), names(3, 4), [noteName(5)]]
    await expect
      .poll(() => printed(page).then((a) => a.map(columnsOf)), { timeout: 60_000 })
      .toEqual([firstInThree, [names(6, 8), names(9, 11)]])
    const sectioned = await printed(page)
    expect(leftsOf(sectioned[0])).toEqual([0, ...OF_THREE])
    contained(sectioned[0])
    await page.screenshot({ path: test.info().outputPath('sections.png'), fullPage: true })

    // the second section states one column: its notes go one under the other again
    const second = await openDialog(app, page, 'Para 7 body')
    await expect(second.locator('select[name="note-columns"]')).toHaveValue('match')
    await second.locator('select[name="note-apply-to"]').selectOption('section')
    await second.locator('select[name="note-columns"]').selectOption('1')
    await second.locator('button[name="apply"]').click()
    await expect(second).toBeHidden()
    await expect
      .poll(() => printed(page).then((a) => a.map(columnsOf)), { timeout: 60_000 })
      .toEqual([firstInThree, [names(6, 11)]])

    await saveFromMenu(app)
    await expect
      .poll(() => part(docPath, 'word/document.xml'), { timeout: 30_000 })
      .toContain('<w15:footnoteColumns w:val="1"/></w:sectPr></w:body>')
    const saved = await part(docPath, 'word/document.xml')
    expect(saved).toContain('<w15:footnoteColumns w:val="3"/></w:sectPr></w:pPr>')
    expect(saved.match(/<w15:footnoteColumns /g)).toHaveLength(2)
    keep(docPath, 'columns-sections.docx')

    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    session = await openDoc(docPath)
    ;({ app, page } = session)
    await expect
      .poll(() => printed(page).then((a) => a.map(columnsOf)), { timeout: 60_000 })
      .toEqual([firstInThree, [names(6, 11)]])
    expect(await edited(page)).toBe(false)
    const reopened = await openDialog(app, page, 'Para 2 body')
    await expect(reopened.locator('select[name="note-columns"]')).toHaveValue('3')
    await page.keyboard.press('Escape')
    await expect(reopened).toBeHidden()
    const other = await openDialog(app, page, 'Para 7 body')
    await expect(other.locator('select[name="note-columns"]')).toHaveValue('1')
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a page takes more lines once its notes stand in columns, and the area stays clear of the text', async () => {
  test.setTimeout(420_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-footnote-columns-')))
  const docPath = join(dir, 'flow.docx')
  // sixty lines with a note each; every fifth note runs over several lines
  const body = Array.from({ length: 60 }, (_, i) => noted(i + 1)).join('') + SECT_PR
  writeFileSync(docPath, await docxWithNotes(body, 60, (n) => n % 5 === 0))
  const session = await openDoc(docPath)
  try {
    const { app, page } = session
    await expect
      .poll(() => onScreen(page).then((a) => a.length), { timeout: 30_000 })
      .toBeGreaterThan(0)
    const before = await printed(page)
    expect(before.flatMap((area) => area.notes.map((n) => n.name))).toEqual(names(1, 60))

    const dialog = await openDialog(app, page, 'Para 1 body')
    await dialog.locator('select[name="note-columns"]').selectOption('2')
    await dialog.locator('button[name="apply"]').click()
    await expect(dialog).toBeHidden()
    await expect
      .poll(() => onScreen(page).then((a) => a.length > 0 && a.every((area) => area.boxed)), {
        timeout: 30_000,
      })
      .toBe(true)
    const after = await printed(page)
    // every note once, in order, and the first page holds more of them than it did
    expect(after.flatMap((area) => area.notes.map((n) => n.name))).toEqual(names(1, 60))
    expect(after[0].notes.length).toBeGreaterThan(before[0].notes.length)
    expect(after.length).toBeLessThanOrEqual(before.length)
    for (const area of after) {
      expect(leftsOf(area).every((left) => left === 0 || left === SECOND_OF_TWO)).toBe(true)
      contained(area)
    }
    await page.screenshot({ path: test.info().outputPath('flow.png'), fullPage: true })

    // in the PDF no body line reaches into the notes and no note runs under the bottom margin
    const words = await exportPdf(app, page, join(dir, 'flow.pdf'))
    const pages = [...new Set(words.map((w) => w.page))]
    expect(pages).toHaveLength(after.length)
    for (const p of pages) {
      const on = words.filter((w) => w.page === p)
      const body = on.filter((w) => /Para \d+ body/.test(w.text))
      const firstNote = Math.min(...on.filter((w) => /NOTE\d\d/.test(w.text)).map((w) => w.y))
      expect(Math.max(...body.map((w) => w.y))).toBeLessThan(firstNote)
      expect(Math.max(...on.map((w) => w.y))).toBeLessThanOrEqual(792 - 72 + 4)
      expect(Math.max(...on.map((w) => w.x))).toBeLessThan(540)
    }
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
