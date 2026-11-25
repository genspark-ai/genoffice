/**
 * PAGE, NUMPAGES and page cross-references put in through the Field and
 * Cross-reference dialogs are saved with the number the page shows where they
 * stand (a roman section, a section that restarts at 1, a table that runs over
 * two pages) and without `w:dirty`. A save from Draft leaves a new field
 * marked; untouched fields go back byte for byte; a footer written anew stores
 * the number of the first page it shows on. A section in columns beside one in a single column, a
 * bookmark that opens inside a paragraph running over two pages, and digits
 * asked for in full width are saved the same way; so are the fields of a
 * section whose text runs sideways, those of text boxes, and a page reference
 * to a bookmark in a footnote.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-page-fields-saved'
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const PAGE_SETUP =
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
const FOOTER_REF = '<w:footerReference w:type="default" r:id="rId7"/>'

const para = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
const PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
const field = (instr: string, result: string, dirty = false) =>
  `<w:r><w:fldChar w:fldCharType="begin"${dirty ? ' w:dirty="true"' : ''}/></w:r>` +
  `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  `<w:r><w:t>${result}</w:t></w:r>` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
const FOOTER = `<?xml version="1.0" encoding="UTF-8"?><w:ftr ${W} ${R}><w:p><w:r><w:t xml:space="preserve">Sheet </w:t></w:r>${field('PAGE', '9')}</w:p></w:ftr>`

async function docx(body: string): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  zip.file(
    'word/_rels/document.xml.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>',
  )
  zip.file('word/footer1.xml', FOOTER)
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${W} ${R}><w:body>${body}</w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

const ROWS = 40
const DEEP = `Row ${ROWS - 2} `
const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>' +
  Array.from(
    { length: ROWS },
    (_, i) =>
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>${para(`Row ${i + 1} `)}</w:tc></w:tr>`,
  ).join('') +
  '</w:tbl>'

/** a roman front section of two pages, then a section that restarts at 1 and holds a table over two pages */
const BODY =
  para('Front matter') +
  PAGE_BREAK +
  `<w:p><w:bookmarkStart w:id="1" w:name="RomanMark"/><w:r><w:t>Roman target</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>` +
  `<w:p><w:r><w:t xml:space="preserve">Stored page: </w:t></w:r>${field('PAGE', '7')}</w:p>` +
  `<w:p><w:pPr><w:sectPr>${FOOTER_REF}${PAGE_SETUP}<w:pgNumType w:fmt="lowerRoman"/></w:sectPr></w:pPr></w:p>` +
  para('Chapter one') +
  para('Page field here: ') +
  para('Count here: ') +
  para('Reference here: ') +
  para('Roman reference here: ') +
  para('Relative here: ') +
  PAGE_BREAK +
  TABLE +
  para('Tail paragraph ') +
  `<w:sectPr>${FOOTER_REF}${PAGE_SETUP}<w:pgNumType w:start="1"/></w:sectPr>`

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

type Hook = {
  __aidocs: { editor: import('@tiptap/core').Editor; pagesSettled: () => boolean }
}

/** the caret at the end of the paragraph that starts with `text` */
const caretIn = (page: Page, text: string) =>
  page.evaluate((text) => {
    const { editor } = (window as unknown as Hook).__aidocs
    let to = -1
    editor.state.doc.descendants((node, pos) => {
      if (to < 0 && node.isTextblock && node.textContent.startsWith(text))
        to = pos + node.nodeSize - 1
      return to < 0
    })
    if (to < 0) throw new Error(`no paragraph starts with ${text}`)
    editor.chain().focus().setTextSelection(to).run()
  }, text)

const settled = (page: Page) =>
  page.waitForFunction(() => (window as unknown as Hook).__aidocs.pagesSettled(), undefined, {
    timeout: 30_000,
  })

async function insertField(app: ElectronApplication, page: Page, at: string, code: string) {
  await caretIn(page, at)
  await clickMenu(app, ['Insert', 'Field…'])
  const dialog = page.locator('.field-dialog').first()
  await dialog.getByLabel('Field code').first().fill(code)
  await dialog.locator('button.btn-primary').click()
  await expect(dialog).toBeHidden()
}

/**
 * What the page shows, read off the DOM and not off the layout the save uses: the number in
 * the footer of each sheet, and for each field and paragraph the sheet it is drawn on (the
 * page-turn bands above it).
 */
const shown = (page: Page) =>
  page.evaluate((deep) => {
    const top = (el: Element) => el.getBoundingClientRect().top
    const footers = [
      ...document.querySelectorAll('.page-gap-hf.page-hf-footer'),
      ...document.querySelectorAll('.page-wrap > .page-hf-footer'),
    ]
    const turns = footers.slice(0, -1).map(top)
    const sheetOf = (el: Element) => turns.filter((y) => y < top(el)).length
    const numbers = footers.map((f) => (f.textContent ?? '').replace('Sheet', '').trim())
    const fields = [...document.querySelectorAll('.doc-page .doc-ref-field')].map((el) => ({
      instr: el.getAttribute('data-instr-field') ?? '',
      text: el.textContent ?? '',
      sheet: sheetOf(el),
    }))
    const sheetOfText = (text: string) => {
      const el = [...document.querySelectorAll('.doc-page p')].find((p) =>
        (p.textContent ?? '').startsWith(text),
      )
      return el ? sheetOf(el) : -1
    }
    return {
      numbers,
      fields,
      deep: sheetOfText(deep),
      roman: sheetOfText('Roman target'),
      firstRow: sheetOfText('Row 1 '),
    }
  }, DEEP)

const part = async (docPath: string, name: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file(name)?.async('string') ?? ''

/** each page field of the saved body: its code, its result, whether it is marked */
const savedFields = (xml: string) =>
  [
    ...xml.matchAll(
      /<w:fldChar w:fldCharType="begin"( w:dirty="true")?\/><\/w:r><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:instrText[^>]*>\s*((?:PAGEREF|PAGE|NUMPAGES)\b[^<]*?)\s*<\/w:instrText>.*?<w:t[^>]*>([^<]*)<\/w:t>/g,
    ),
  ].map((m) => ({ instr: m[2], result: m[3], dirty: m[1] !== undefined }))

const fieldParagraphs = (xml: string) =>
  [...xml.matchAll(/<w:p>(?:(?!<\/w:p>).)*<\/w:p>/g)]
    .map((m) => m[0])
    .filter((p) => p.includes('w:fldChar') && !p.includes('<w:tbl>'))

test('page fields put in through the dialogs are saved with the number their page shows, unmarked', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-page-fields-')))
  const docPath = join(dir, 'pages.docx')
  writeFileSync(docPath, await docx(BODY))
  let launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
  try {
    let { app } = launched
    let page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'Tail paragraph' }).first()).toBeVisible({
      timeout: 30_000,
    })

    await insertField(app, page, 'Page field here: ', 'PAGE')
    await insertField(app, page, 'Count here: ', 'NUMPAGES')
    await insertField(app, page, 'Row 3 ', 'PAGE')
    await insertField(app, page, DEEP, 'PAGE')

    const bookmark = page.locator('.bookmark-dialog')
    await caretIn(page, DEEP)
    await clickMenu(app, ['Insert', 'Bookmark…'])
    await bookmark.locator('input[name="bookmark-name"]').fill('Deep')
    await bookmark.locator('button[name="add"]').click()
    await expect(bookmark.locator('.bookmark-row', { hasText: 'Deep' })).toHaveCount(1)
    await page.keyboard.press('Escape')
    await expect(bookmark).toBeHidden()

    const crossRef = page.locator('.crossref-dialog')
    const pageRef = async (at: string, name: string, aboveBelow = false) => {
      await caretIn(page, at)
      await clickMenu(app, ['Insert', 'Cross-reference…'])
      await expect(crossRef).toBeVisible()
      await crossRef.locator('select[name="ref-type"]').selectOption('bookmark')
      await crossRef.locator('select[name="ref-content"]').selectOption('page')
      if (aboveBelow) await crossRef.locator('input[name="ref-above-below"]').check()
      await crossRef.locator('.bookmark-row', { hasText: name }).click()
      await crossRef.locator('button[name="insert"]').click()
      if (await crossRef.isVisible()) await page.keyboard.press('Escape')
      await expect(crossRef).toBeHidden()
    }
    await pageRef('Reference here: ', 'Deep')
    await pageRef('Roman reference here: ', 'RomanMark')
    await pageRef('Relative here: ', 'Deep', true)

    await settled(page)
    const before = await shown(page)
    // two roman sheets, the chapter's opening sheet, and the table running over the sheets after it
    expect(before.numbers.slice(0, 4)).toEqual(['i', 'ii', '1', '2'])
    expect(before.firstRow).toBe(3)
    expect(before.deep).toBeGreaterThan(3)
    await page.screenshot({ path: test.info().outputPath('page-fields.png') })

    await saveFromMenu(app)
    await expect
      .poll(async () => savedFields(await part(docPath, 'word/document.xml')).length, {
        timeout: 30_000,
      })
      .toBe(8)
    const xml = await part(docPath, 'word/document.xml')
    expect(xml).not.toContain('w:dirty')
    const numberOf = (sheet: number) => before.numbers[sheet]
    expect(savedFields(xml).map((f) => `${f.instr} = ${f.result}`)).toEqual([
      // the field the file brought said 7; the page it stands on shows ii
      `PAGE = ${numberOf(1)}`,
      `PAGE \\* MERGEFORMAT = ${numberOf(2)}`,
      `NUMPAGES \\* MERGEFORMAT = ${before.numbers.length}`,
      `PAGEREF Deep \\h = ${numberOf(before.deep)}`,
      `PAGEREF RomanMark \\h = ${numberOf(before.roman)}`,
      `PAGEREF Deep \\p \\h = on page ${numberOf(before.deep)}`,
      // in the table: the row's own sheet, not the one the table opens on
      `PAGE \\* MERGEFORMAT = ${numberOf(3)}`,
      `PAGE \\* MERGEFORMAT = ${numberOf(before.deep)}`,
    ])
    expect(numberOf(before.roman)).toBe('ii')
    expect(numberOf(3)).toBe('2')
    expect(Number(numberOf(before.deep))).toBe(before.deep - 1)
    // what the page shows after the save is what the file holds
    const after = await shown(page)
    expect(after.fields.map((f) => f.text)).toEqual(savedFields(xml).map((f) => f.result))
    for (const f of after.fields.filter((f) => /^PAGE\b/.test(f.instr)))
      expect(f.text, `PAGE on sheet ${f.sheet}`).toBe(after.numbers[f.sheet])
    // the footer part was not touched: it goes back as it came, stale result and all
    expect(await part(docPath, 'word/footer1.xml')).toBe(FOOTER)

    // ---- Draft has no pages: a field put in there waits, and Update Field writes no page ----
    await clickMenu(app, ['View', 'Draft'])
    await insertField(app, page, 'Tail paragraph ', 'PAGE')
    await page.keyboard.press('F9')
    await saveFromMenu(app)
    await expect
      .poll(async () => savedFields(await part(docPath, 'word/document.xml')).length, {
        timeout: 30_000,
      })
      .toBe(9)
    const draft = savedFields(await part(docPath, 'word/document.xml'))
    expect(draft.filter((f) => f.dirty).map((f) => f.instr)).toEqual(['PAGE \\* MERGEFORMAT'])
    expect(draft.at(-1)?.dirty).toBe(true)
    expect(draft.slice(0, 8)).toEqual(savedFields(xml))

    await clickMenu(app, ['View', 'Print Layout'])
    await settled(page)
    // a save does not run when nothing changed since the last one
    await caretIn(page, 'Chapter one')
    await page.keyboard.type('!')
    await settled(page)
    await saveFromMenu(app)
    await expect
      .poll(async () => (await part(docPath, 'word/document.xml')).includes('w:dirty'), {
        timeout: 30_000,
      })
      .toBe(false)
    const paged = await part(docPath, 'word/document.xml')
    const end = await shown(page)
    expect(savedFields(paged).at(-1)?.result).toBe(end.numbers.at(-1))
    expect(savedFields(paged).slice(0, 8)).toEqual(savedFields(xml))

    // ---- reopened: the results show as saved, and an edit elsewhere leaves every field's bytes alone ----
    await closeAndSaveVideo(launched, VIDEO_DIR)
    launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
    app = launched.app
    page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'Tail paragraph' }).first()).toBeVisible({
      timeout: 30_000,
    })
    await settled(page)
    const reopened = await shown(page)
    expect(reopened.numbers).toEqual(end.numbers)
    expect(reopened.fields.map((f) => f.text)).toEqual(savedFields(paged).map((f) => f.result))
    await caretIn(page, 'Front matter')
    await page.keyboard.type('!')
    await settled(page)
    await saveFromMenu(app)
    await expect
      .poll(() => part(docPath, 'word/document.xml'), { timeout: 30_000 })
      .toContain('Front matter!')
    const again = await part(docPath, 'word/document.xml')
    expect(again).not.toContain('w:dirty')
    expect(fieldParagraphs(again)).toEqual(fieldParagraphs(paged))
    expect(savedFields(again)).toEqual(savedFields(paged))
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a footer typed into stores the number of the first page it shows on', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-page-fields-hf-')))
  const docPath = join(dir, 'footer.docx')
  writeFileSync(
    docPath,
    await docx(
      para('First sheet') +
        PAGE_BREAK +
        para('Second sheet') +
        `<w:sectPr>${FOOTER_REF}${PAGE_SETUP}<w:pgNumType w:fmt="lowerRoman" w:start="3"/></w:sectPr>`,
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
    await expect(page.locator('.doc-page p', { hasText: 'Second sheet' }).first()).toBeVisible({
      timeout: 30_000,
    })
    await settled(page)
    expect((await shown(page)).numbers).toEqual(['iii', 'iv'])
    // the open-time respell kick focuses the body and shields keystrokes for a moment
    await page.waitForTimeout(1500)
    const footer = page.locator('.page-wrap > .page-hf-footer')
    await footer.scrollIntoViewIfNeeded()
    await footer.dblclick()
    const editor = footer.locator('.page-hf-editor')
    await expect(editor).toBeFocused()
    await page.keyboard.press('End')
    await page.keyboard.type(' of the report')
    await page.getByRole('button', { name: 'Close Header and Footer' }).click()
    await expect(editor).toHaveCount(0)
    await settled(page)
    await saveFromMenu(app)
    await expect
      .poll(() => part(docPath, 'word/footer1.xml'), { timeout: 30_000 })
      .toContain('of the report')
    const xml = await part(docPath, 'word/footer1.xml')
    expect(savedFields(xml)).toEqual([{ instr: 'PAGE', result: 'iii', dirty: false }])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

const LINES = Array.from({ length: 150 }, (_, i) =>
  i % 12 === 5
    ? `<w:p><w:r><w:t xml:space="preserve">Line ${i + 1} is on </w:t></w:r>${field('PAGE', '0', true)}</w:p>`
    : para(`Line ${i + 1}`),
).join('')

/** opens 150 lines with a marked PAGE field in every twelfth, changes the document, saves; `check` gets the saved fields */
async function saveLines<T = undefined>(
  name: string,
  sectPr: string,
  zoom: string | null,
  check: (
    saved: ReturnType<typeof savedFields>,
    before: Awaited<ReturnType<typeof shown>>,
    probed: T,
  ) => void,
  body = LINES,
  probe?: (page: Page) => Promise<T>,
) {
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-page-fields-lines-')))
  const docPath = join(dir, name)
  writeFileSync(
    docPath,
    await docx(body + `<w:sectPr>${FOOTER_REF}${PAGE_SETUP}${sectPr}</w:sectPr>`),
  )
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'Line 150' }).first()).toBeVisible({
      timeout: 30_000,
    })
    if (zoom) {
      const zoomValue = page.locator('.zoom-value')
      await zoomValue.click()
      const dialog = page.locator('.zoom-dialog')
      await dialog.locator('.zoom-dialog-radio', { hasText: zoom }).click()
      await dialog.locator('button[type="submit"]').click()
      await expect(zoomValue).toHaveText(zoom)
    }
    // a closing dialog hands the focus back a moment later: type once the caret is in the text
    await expect(async () => {
      await caretIn(page, 'Line 1')
      await page.keyboard.type('!')
      await expect(page.locator('.doc-page p', { hasText: 'Line 1!' })).toHaveCount(1, {
        timeout: 1000,
      })
    }).toPass({ timeout: 15_000 })
    // the pass after the edit has run: the layout is the document's, columns or not
    await page.waitForTimeout(2500)
    const before = await shown(page)
    const probed = (await probe?.(page)) as T
    await saveFromMenu(app)
    await expect
      .poll(async () => /Line 1!+</.test(await part(docPath, 'word/document.xml')), {
        timeout: 30_000,
      })
      .toBe(true)
    check(savedFields(await part(docPath, 'word/document.xml')), before, probed)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
}

test('on a zoomed canvas each field still takes the number of the sheet it is drawn on', async () => {
  test.setTimeout(180_000)
  await saveLines('zoomed.docx', '<w:pgNumType w:start="5"/>', '200%', (saved, before) => {
    expect(before.numbers.length).toBeGreaterThan(2)
    expect(before.numbers[0]).toBe('5')
    expect(saved).toHaveLength(13)
    expect(saved.map((f) => f.dirty)).toEqual(saved.map(() => false))
    expect(saved.map((f) => f.result)).toEqual(before.fields.map((f) => before.numbers[f.sheet]))
    expect(new Set(saved.map((f) => f.result)).size).toBeGreaterThan(2)
  })
})

const sheetNumbers = (
  saved: ReturnType<typeof savedFields>,
  before: Awaited<ReturnType<typeof shown>>,
) => {
  expect(saved).toHaveLength(13)
  expect(saved.map((f) => f.dirty)).toEqual(saved.map(() => false))
  expect(saved.map((f) => f.result)).toEqual(before.fields.map((f) => before.numbers[f.sheet]))
  expect(new Set(saved.map((f) => f.result)).size).toBeGreaterThan(1)
}

test('a document set in the same columns throughout keeps the mark: its columns are drawn as one flow down every sheet', async () => {
  test.setTimeout(180_000)
  await saveLines('columns.docx', '<w:cols w:num="2" w:space="720"/>', null, (saved) => {
    expect(saved).toHaveLength(13)
    expect(saved.map((f) => `${f.result}${f.dirty ? ' dirty' : ''}`)).toEqual(
      saved.map(() => '0 dirty'),
    )
  })
})

test('in a section in columns that follows one in a single column each field takes the number of the sheet its line is on', async () => {
  test.setTimeout(180_000)
  const single = `<w:p><w:pPr><w:sectPr>${FOOTER_REF}${PAGE_SETUP}</w:sectPr></w:pPr></w:p>`
  const at = LINES.indexOf(para('Line 31'))
  await saveLines(
    'columns-after-single.docx',
    '<w:type w:val="continuous"/><w:cols w:num="2" w:space="720"/>',
    null,
    sheetNumbers,
    LINES.slice(0, at) + single + LINES.slice(at),
  )
})

test('in a section whose text runs sideways each field takes the number of the sheet its line is on', async () => {
  test.setTimeout(180_000)
  await saveLines('sideways.docx', '<w:textDirection w:val="tbRl"/>', null, sheetNumbers)
})

const DRAWING_NS =
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
/** a text box under its paragraph (`wrap` top and bottom) or floating over the text (none) */
const textBox = (id: number, wrap: 'TopAndBottom' | 'None', inner: string) =>
  `<w:p><w:r><mc:AlternateContent ${DRAWING_NS}><mc:Choice Requires="wps"><w:drawing>` +
  '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  `<wp:positionH relativeFrom="column"><wp:posOffset>${wrap === 'None' ? 3200000 : 0}</wp:posOffset></wp:positionH>` +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>300000</wp:posOffset></wp:positionV>' +
  `<wp:extent cx="2400000" cy="900000"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrap${wrap}/>` +
  `<wp:docPr id="${id}" name="Text Box ${id}"/><wp:cNvGraphicFramePr/>` +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2400000" cy="900000"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>' +
  `<wps:txbx><w:txbxContent><w:p>${inner}</w:p></w:txbxContent></wps:txbx>` +
  '<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r></w:p>'
const text = (value: string) => `<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`

const FOOTNOTES =
  `<?xml version="1.0" encoding="UTF-8"?><w:footnotes ${W}>` +
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:bookmarkStart w:id="90" w:name="Noted"/>' +
  `${text(' A note at the foot of the second sheet')}<w:bookmarkEnd w:id="90"/></w:p></w:footnote></w:footnotes>`

/** page fields that wait for their page: one that points into a footnote, and those of two text boxes on the second sheet */
const OUTSIDE_BODY =
  para('Line 1') +
  `<w:p>${text('Note page: ')}${field('PAGEREF Noted \\h', '0', true)}</w:p>` +
  PAGE_BREAK +
  `<w:p>${text('Second sheet')}<w:r><w:footnoteReference w:id="2"/></w:r></w:p>` +
  textBox(
    61,
    'TopAndBottom',
    text('Box page ') + field('PAGE', '0', true) + text(' of ') + field('NUMPAGES', '0', true),
  ) +
  textBox(62, 'None', text('Floating ') + field('PAGE', '0', true)) +
  para('Line 150')

test('page fields in text boxes and a page reference into a footnote are saved with their sheet, unmarked', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-page-fields-outside-')))
  const docPath = join(dir, 'outside.docx')
  const zip = await JSZip.loadAsync(
    await docx(OUTSIDE_BODY + `<w:sectPr>${FOOTER_REF}${PAGE_SETUP}</w:sectPr>`),
  )
  zip.file('word/footnotes.xml', FOOTNOTES)
  const rels = 'word/_rels/document.xml.rels'
  zip.file(
    rels,
    (await zip.file(rels)!.async('string')).replace(
      '</Relationships>',
      '<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>',
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
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'Line 150' }).first()).toBeVisible({
      timeout: 30_000,
    })
    await expect(async () => {
      await caretIn(page, 'Line 1')
      await page.keyboard.type('!')
      await expect(page.locator('.doc-page p', { hasText: 'Line 1!' })).toHaveCount(1, {
        timeout: 1000,
      })
    }).toPass({ timeout: 15_000 })
    await page.waitForTimeout(2500)
    const before = await shown(page)
    expect(before.numbers).toEqual(['1', '2'])
    await page.screenshot({ path: test.info().outputPath('page-fields-outside-body.png') })
    await saveFromMenu(app)
    await expect
      .poll(async () => /Line 1!+</.test(await part(docPath, 'word/document.xml')), {
        timeout: 30_000,
      })
      .toBe(true)
    const xml = await part(docPath, 'word/document.xml')
    expect(savedFields(xml)).toEqual([
      { instr: 'PAGEREF Noted \\h', result: '2', dirty: false },
      { instr: 'PAGE', result: '2', dirty: false },
      { instr: 'NUMPAGES', result: '2', dirty: false },
      { instr: 'PAGE', result: '2', dirty: false },
    ])
    // the footnote was not typed into: its entry and its bookmark are the file's
    expect(await part(docPath, 'word/footnotes.xml')).toBe(FOOTNOTES)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

const LONG = 14
const WORDS = Array.from({ length: 150 }, (_, i) => `word${i}`).join(' ')
/** paragraphs of some ten lines, each with a bookmark that opens before its last word */
const LONG_BODY =
  para('Line 1') +
  Array.from(
    { length: LONG },
    (_, i) =>
      `<w:p><w:r><w:t xml:space="preserve">Long ${i} ${WORDS} </w:t></w:r><w:bookmarkStart w:id="${40 + i}" w:name="Mid${i}"/><w:r><w:t>mark${i}</w:t></w:r><w:bookmarkEnd w:id="${40 + i}"/></w:p>`,
  ).join('') +
  Array.from(
    { length: LONG },
    (_, i) =>
      `<w:p><w:r><w:t xml:space="preserve">Ref ${i}: </w:t></w:r>${field(`PAGEREF Mid${i} \\h`, '0', true)}</w:p>`,
  ).join('') +
  `<w:p><w:r><w:t xml:space="preserve">Count: </w:t></w:r>${field('NUMPAGES \\* DBCHAR', '0', true)}</w:p>` +
  para('Line 150')

/** for each long paragraph: the sheets its first line, its last line and its bookmarked word are drawn on */
const longSheets = (page: Page) =>
  page.evaluate((count) => {
    const footers = [
      ...document.querySelectorAll('.page-gap-hf.page-hf-footer'),
      ...document.querySelectorAll('.page-wrap > .page-hf-footer'),
    ]
    const turns = footers.slice(0, -1).map((f) => f.getBoundingClientRect().top)
    const sheetAt = (y: number) => turns.filter((turn) => turn < y).length
    return Array.from({ length: count }, (_, i) => {
      const el = [...document.querySelectorAll('.doc-page p')].find((p) =>
        (p.textContent ?? '').startsWith(`Long ${i} `),
      )
      if (!el) return { first: -1, last: -1, mark: -1 }
      const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT)
      let mark = -1
      for (let node = walker.nextNode(); node; node = walker.nextNode()) {
        const at = (node as Text).data.indexOf(`mark${i}`)
        if (at < 0) continue
        const range = document.createRange()
        range.setStart(node, at)
        range.setEnd(node, at + 1)
        const box = range.getBoundingClientRect()
        mark = sheetAt((box.top + box.bottom) / 2)
      }
      const lines = [...el.getClientRects()]
      return {
        first: sheetAt(lines[0].top + 2),
        last: sheetAt(lines[lines.length - 1].bottom - 2),
        mark,
      }
    })
  }, LONG)

test('a page reference to a bookmark inside a paragraph that runs over two pages takes the page the bookmark opens on', async () => {
  test.setTimeout(180_000)
  await saveLines(
    'mid-bookmark.docx',
    '',
    null,
    (saved, before, sheets) => {
      expect(saved).toHaveLength(LONG + 1)
      expect(saved.map((f) => f.dirty)).toEqual(saved.map(() => false))
      // the case this is about: a paragraph whose bookmark is not on the page it starts on
      expect(sheets.some((s) => s.mark > s.first)).toBe(true)
      expect(saved.slice(0, LONG).map((f) => f.result)).toEqual(
        sheets.map((s) => before.numbers[s.mark]),
      )
      const fullWidth = String(before.numbers.length).replace(/\d/g, (d) =>
        String.fromCharCode(0xff10 + Number(d)),
      )
      expect(saved[LONG]).toEqual({ instr: 'NUMPAGES \\* DBCHAR', result: fullWidth, dirty: false })
    },
    LONG_BODY,
    longSheets,
  )
})
