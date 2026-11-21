/**
 * Two controls that used to be greyed out:
 * Insert > Index and Tables > Index > Columns puts the index in a section of
 * its own with that many columns, as Word does; Insert > Cross-reference
 * refers to a footnote or an endnote through a NOTEREF field and a bookmark
 * on the note's reference mark.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { existsSync, mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-index-columns-noteref'
const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/></w:sectPr>'

const para = (text: string, extra = '') =>
  `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r>${extra}</w:p>`
const field = (instr: string, result = '') =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  (result ? `<w:r><w:t>${result}</w:t></w:r>` : '') +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
const xe = (term: string) => field(`XE "${term}"`)
const noteMark = (id: number) =>
  `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`
const PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

const CONTENT_TYPES =
  '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
const FOOTNOTES_TYPE =
  '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'

async function docx(body: string, footnotes: string[] = []): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `${CONTENT_TYPES}${footnotes.length > 0 ? FOOTNOTES_TYPE : ''}</Types>`,
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  zip.file(
    'word/_rels/document.xml.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      (footnotes.length > 0
        ? '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>'
        : '') +
      '</Relationships>',
  )
  zip.file(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:styles ${NS}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`,
  )
  if (footnotes.length > 0)
    zip.file(
      'word/footnotes.xml',
      `<?xml version="1.0" encoding="UTF-8"?><w:footnotes ${NS}>` +
        '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
        '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
        footnotes
          .map(
            (text, i) =>
              `<w:footnote w:id="${i + 1}"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> ${text}</w:t></w:r></w:p></w:footnote>`,
          )
          .join('') +
        '</w:footnotes>',
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

type EditorWindow = { __aidocs: { editor: import('@tiptap/core').Editor } }

/** puts the caret at the end of the paragraph starting with `text` */
const caretIn = (page: Page, text: string) =>
  page.evaluate((text) => {
    const { editor } = (window as unknown as EditorWindow).__aidocs
    let at: number | null = null
    editor.state.doc.descendants((node, pos) => {
      if (at === null && node.isTextblock && node.textContent.startsWith(text))
        at = pos + node.nodeSize - 1
      return at === null
    })
    if (at === null) throw new Error(`no paragraph starts with ${text}`)
    editor.chain().focus().setTextSelection(at).run()
  }, text)

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

/** how many columns the index entries stand in: the distinct left edges of their lines on the page */
const indexColumns = (page: Page, terms: string[]) =>
  page.evaluate((terms) => {
    const lefts = new Set<number>()
    for (const term of terms) {
      // the innermost element that shows the entry: a generated line, or a line read from the file
      const line = Array.from(document.querySelectorAll('main.editor-scroll .doc-page *'))
        .filter((el) => new RegExp(`^${term}\\b`).test((el.textContent ?? '').trim()))
        .pop()
      if (line) lefts.add(Math.round(line.getBoundingClientRect().left))
    }
    return lefts.size
  }, terms)

/** the sections of the saved file in order: column count, and whether the section starts on the same page */
const sections = async (docPath: string) =>
  ((await documentXml(docPath)).match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g) ?? []).map(
    (sectPr) =>
      `${/<w:cols\b[^>]*w:num="(\d+)"/.exec(sectPr)?.[1] ?? '1'}` +
      (/<w:type w:val="continuous"\/>/.test(sectPr) ? ' continuous' : ''),
  )

const TERMS = ['Apple', 'Banana', 'Cherry', 'Damson', 'Elder', 'Fig', 'Grape', 'Hazel']

test('an index with columns sits in a continuous section of its own with that many columns', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-index-columns-')))
  const docPath = join(dir, 'index.docx')
  writeFileSync(
    docPath,
    await docx(
      para('Orchard notes') +
        TERMS.map((term) => para(`about ${term.toLowerCase()}`, xe(term))).join('') +
        para('Index goes here') +
        para('After the index'),
    ),
  )
  let session = await openDoc(docPath)
  try {
    let { app, page } = session
    const dialog = page.locator('.index-tables-dialog')
    const columns = dialog.locator('[data-it="columns"]')
    const openIndexTab = async (onIndex: boolean) => {
      await clickMenu(app, ['Insert', 'Index and Tables…'])
      await expect(dialog).toBeVisible()
      const tab = dialog.getByRole('tab', { name: 'Index', exact: true })
      if (onIndex) await expect(tab).toHaveClass(/active/)
      else await tab.click()
    }
    // an index read from the file: its lines are the field's own paragraphs
    const clickIndexLine = (term: string) =>
      page
        .locator('main.editor-scroll .doc-page .doc-col-block', {
          hasText: new RegExp(`^${term},`),
        })
        .first()
        .click()

    await caretIn(page, 'Index goes here')
    await openIndexTab(false)
    await expect(columns).toBeEnabled()
    await expect(columns).toHaveValue('1')
    await columns.fill('2')
    await page.screenshot({ path: test.info().outputPath('index-columns-dialog.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    // a section break in a saved document is written at once and the file read again
    await expect
      .poll(() => sections(docPath), { timeout: 30_000 })
      .toEqual(['1', '2 continuous', '1 continuous'])
    await expect(page.locator('.doc-page .doc-protected-field-text').first()).toBeVisible({
      timeout: 30_000,
    })
    await expect.poll(() => indexColumns(page, TERMS), { timeout: 30_000 }).toBe(2)
    // the text around the index stays in one column, on the same page
    const widthOf = async (text: string | RegExp) =>
      (
        await page
          .locator('.doc-page .doc-col-block, .doc-page p', { hasText: text })
          .first()
          .boundingBox()
      )?.width ?? 0
    const lineWidth = await widthOf(/^Apple,/)
    expect(lineWidth).toBeGreaterThan(0)
    expect(await widthOf('After the index')).toBeGreaterThan(lineWidth * 1.8)
    expect(await widthOf('Index goes here')).toBeGreaterThan(lineWidth * 1.8)
    // the index is balanced over its columns and the next section starts below it
    const topOf = async (text: string | RegExp) =>
      (
        await page
          .locator('.doc-page .doc-col-block, .doc-page p', { hasText: text })
          .first()
          .boundingBox()
      )?.y ?? 0
    expect(await topOf(/^Elder,/)).toBeCloseTo(await topOf(/^Apple,/), 0)
    expect(await topOf('After the index')).toBeGreaterThan(await topOf(/^Hazel,/))
    await expect(page.locator('.doc-page')).toHaveCount(1)
    await page.screenshot({ path: test.info().outputPath('index-two-columns.png') })

    await saveFromMenu(app)
    await expect
      .poll(() => sections(docPath), { timeout: 30_000 })
      .toEqual(['1', '2 continuous', '1 continuous'])
    let xml = await documentXml(docPath)
    expect(xml).toMatch(/<w:instrText[^>]*> INDEX [^<]*\\c "2" /)
    // the entries and their pages are the field's result: Word is not asked to rebuild the index
    expect(xml).not.toContain('w:dirty')
    // the breaks are around the field: one before its begin, one after its end
    const breaks = [...xml.matchAll(/<w:sectPr\b/g)].map((m) => m.index)
    const fieldStart = xml.indexOf(' INDEX ')
    const fieldEnd = xml.indexOf('w:fldCharType="end"', xml.lastIndexOf('Hazel'))
    expect(breaks[0]).toBeLessThan(fieldStart)
    expect(breaks[1]).toBeGreaterThan(fieldEnd)
    expect(xml.slice(breaks[0], fieldStart)).not.toContain('<w:t')
    expect(xml).toMatch(/<w:cols\b[^>]*w:num="2"[^>]*w:space="720"/)

    // the dialog opens on the index with its count; another count changes the switch and the section
    await clickIndexLine('Cherry')
    await openIndexTab(true)
    await expect(columns).toHaveValue('2')
    await columns.fill('3')
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect.poll(() => indexColumns(page, TERMS), { timeout: 30_000 }).toBe(3)
    await page.screenshot({ path: test.info().outputPath('index-three-columns.png') })
    await saveFromMenu(app)
    await expect
      .poll(() => sections(docPath), { timeout: 30_000 })
      .toEqual(['1', '3 continuous', '1 continuous'])
    xml = await documentXml(docPath)
    expect(xml).toMatch(/<w:instrText[^>]*> INDEX [^<]*\\c "3" /)
    expect(xml).not.toContain('\\c "2"')
    expect(xml).not.toContain('w:dirty')

    // read again from the file, the index is still in three columns and the dialog says so
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    session = await openDoc(docPath)
    ;({ app, page } = session)
    await expect(page.locator('.doc-page .doc-protected-field').first()).toBeVisible({
      timeout: 30_000,
    })
    await expect.poll(() => indexColumns(page, TERMS), { timeout: 30_000 }).toBe(3)
    const reopened = page.locator('.index-tables-dialog')
    await page
      .locator('main.editor-scroll .doc-page .doc-col-block', { hasText: /^Cherry,/ })
      .first()
      .click()
    await clickMenu(app, ['Insert', 'Index and Tables…'])
    await expect(reopened.locator('[data-it="columns"]')).toHaveValue('3')
    // back to one column: the switch goes, the section stays with one column
    await reopened.locator('[data-it="columns"]').fill('1')
    await reopened.locator('button.primary').click()
    await expect(reopened).toBeHidden()
    await expect.poll(() => indexColumns(page, TERMS), { timeout: 30_000 }).toBe(1)
    await saveFromMenu(app)
    await expect
      .poll(() => sections(docPath), { timeout: 30_000 })
      .toEqual(['1', '1 continuous', '1 continuous'])
    expect(await documentXml(docPath)).not.toMatch(/INDEX [^<]*\\c /)
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('an index that shares its section is rebuilt in place when it gets columns', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-index-columns-rewrap-')))
  const docPath = join(dir, 'one-index.docx')
  writeFileSync(
    docPath,
    await docx(
      para('Orchard notes') +
        TERMS.map((term) => para(`about ${term.toLowerCase()}`, xe(term))).join('') +
        para('Index goes here') +
        para('After the index'),
    ),
  )
  const { launched, app, page } = await openDoc(docPath)
  try {
    const dialog = page.locator('.index-tables-dialog')
    const columns = dialog.locator('[data-it="columns"]')
    await caretIn(page, 'Index goes here')
    await clickMenu(app, ['Insert', 'Index and Tables…'])
    await dialog.getByRole('tab', { name: 'Index', exact: true }).click()
    await expect(columns).toHaveValue('1')
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    const lines = page.locator('.doc-page .doc-toc-line')
    await expect(lines).toHaveCount(TERMS.length)

    // one column so far, in the section of the text around it: now two
    await lines.filter({ hasText: 'Cherry' }).locator('.doc-toc-title').click()
    await clickMenu(app, ['Insert', 'Index and Tables…'])
    await expect(dialog.getByRole('tab', { name: 'Index', exact: true })).toHaveClass(/active/)
    await columns.fill('2')
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect
      .poll(() => sections(docPath), { timeout: 30_000 })
      .toEqual(['1', '2 continuous', '1 continuous'])
    const xml = await documentXml(docPath)
    expect(xml.match(/<w:instrText[^>]*> INDEX /g)).toHaveLength(1)
    expect(xml).toMatch(/<w:instrText[^>]*> INDEX \\c "2" </)
    expect(xml).not.toContain('w:dirty')
    // every entry is listed once, in two columns
    for (const term of TERMS) expect(xml.split(`>${term}<`)).toHaveLength(2)
    await expect(page.locator('.doc-page .doc-protected-field-text').first()).toBeVisible({
      timeout: 30_000,
    })
    await expect.poll(() => indexColumns(page, TERMS), { timeout: 30_000 }).toBe(2)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a second index gets the section, not the index already in the document', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-index-columns-second-')))
  const docPath = join(dir, 'two-indexes.docx')
  writeFileSync(
    docPath,
    await docx(
      para('Orchard notes') +
        TERMS.map((term) => para(`about ${term.toLowerCase()}`, xe(term))).join('') +
        para('First index here') +
        para('Second index here') +
        para('After the indexes'),
    ),
  )
  const { launched, app, page } = await openDoc(docPath)
  try {
    const dialog = page.locator('.index-tables-dialog')
    const insertIndex = async (at: string, count: string) => {
      await caretIn(page, at)
      await clickMenu(app, ['Insert', 'Index and Tables…'])
      await dialog.getByRole('tab', { name: 'Index', exact: true }).click()
      await dialog.locator('[data-it="columns"]').fill(count)
      await dialog.locator('button.primary').click()
      await expect(dialog).toBeHidden()
    }
    await insertIndex('First index here', '1')
    await expect(page.locator('.doc-page .doc-toc-line')).toHaveCount(TERMS.length)
    await insertIndex('Second index here', '3')
    await expect
      .poll(() => sections(docPath), { timeout: 30_000 })
      .toEqual(['1', '3 continuous', '1 continuous'])
    const xml = await documentXml(docPath)
    const fields = [...xml.matchAll(/<w:instrText[^>]*> INDEX ([^<]*)</g)]
    expect(fields.map((m) => m[1].trim())).toEqual(['', '\\c "3"'])
    expect(xml).not.toContain('w:dirty')
    // both breaks are behind the first index: one before the second, one after it
    const breaks = [...xml.matchAll(/<w:sectPr\b/g)].map((m) => m.index)
    expect(breaks[0]).toBeGreaterThan(xml.indexOf('Second index here'))
    expect(breaks[0]).toBeLessThan(fields[1].index)
    expect(breaks[1]).toBeGreaterThan(fields[1].index)

    // the first index, which shares its section, gets columns too: it is rebuilt where it is, not added again
    await expect(page.locator('.doc-page .doc-protected-field-text').first()).toBeVisible({
      timeout: 30_000,
    })
    await page
      .locator('main.editor-scroll .doc-page .doc-col-block, main.editor-scroll .doc-page p', {
        hasText: /^Cherry,/,
      })
      .first()
      .click()
    await clickMenu(app, ['Insert', 'Index and Tables…'])
    await expect(dialog.getByRole('tab', { name: 'Index', exact: true })).toHaveClass(/active/)
    await expect(dialog.locator('[data-it="columns"]')).toHaveValue('1')
    await dialog.locator('[data-it="columns"]').fill('2')
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    const counts = async () => (await sections(docPath)).map((c) => c.split(' ')[0])
    await expect.poll(counts, { timeout: 30_000 }).toEqual(['1', '2', '1', '3', '1'])
    const both = [...(await documentXml(docPath)).matchAll(/<w:instrText[^>]*> INDEX ([^<]*)</g)]
    expect(both.map((m) => m[1].trim())).toEqual(['\\c "2"', '\\c "3"'])
    expect(await documentXml(docPath)).not.toContain('w:dirty')
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a new document that was never saved gets the index section too', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-index-columns-new-')))
  const out = join(dir, 'new-index.docx')
  const seed = join(dir, 'seed.docx')
  writeFileSync(seed, await docx(para('Seed document')))
  const { launched, app, page } = await openDoc(seed)
  try {
    const seeded = page.locator('.doc-page p', { hasText: 'Seed document' })
    await expect(seeded).toBeVisible({ timeout: 30_000 })
    await clickMenu(app, ['File', 'New Document'])
    await expect(seeded).toHaveCount(0, { timeout: 30_000 })
    await page.locator('.doc-page').first().click()
    // entries as Mark Index Entry leaves them, one paragraph each
    await page.evaluate((terms) => {
      const { editor } = (window as unknown as EditorWindow).__aidocs
      editor
        .chain()
        .focus()
        .insertContent([
          ...terms.map((term) => ({
            type: 'docParagraph',
            content: [
              { type: 'text', text: `about ${term.toLowerCase()}` },
              { type: 'docXeMark', attrs: { term } },
            ],
          })),
          { type: 'docParagraph', content: [{ type: 'text', text: 'Index goes here' }] },
          { type: 'docParagraph', content: [{ type: 'text', text: 'After the index' }] },
        ])
        .run()
    }, TERMS)
    await caretIn(page, 'Index goes here')
    await clickMenu(app, ['Insert', 'Index and Tables…'])
    const dialog = page.locator('.index-tables-dialog')
    await dialog.getByRole('tab', { name: 'Index', exact: true }).click()
    await dialog.locator('[data-it="columns"]').fill('2')
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect(page.locator('.doc-page .doc-toc-line')).toHaveCount(TERMS.length)
    // a section in a document that has no file yet is only laid out from its first save on:
    // the breaks are in the document and the status bar says so
    await expect(page.locator('.doc-page .doc-protected-sectbreak')).toHaveCount(2)
    await expect(page.locator('.status-msg')).toContainText('takes effect after saving')
    await page.screenshot({ path: test.info().outputPath('index-unsaved.png') })

    await app.evaluate(({ dialog }, out) => {
      dialog.showSaveDialog = async () => ({ canceled: false, filePath: out })
    }, out)
    await clickMenu(app, ['File', 'Save As…'])
    await expect.poll(() => existsSync(out), { timeout: 30_000 }).toBe(true)
    await expect
      .poll(() => sections(out), { timeout: 30_000 })
      .toEqual(['1', '2 continuous', '1 continuous'])
    expect(await documentXml(out)).toMatch(/<w:instrText[^>]*> INDEX [^<]*\\c "2" /)
    expect(await documentXml(out)).not.toContain('w:dirty')
    await expect.poll(() => indexColumns(page, TERMS), { timeout: 30_000 }).toBe(2)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a cross-reference to a footnote is a NOTEREF field on a bookmark around the note mark', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-noteref-')))
  const docPath = join(dir, 'notes.docx')
  writeFileSync(
    docPath,
    await docx(
      para('Opening words') +
        para('Alpha claim', noteMark(1)) +
        para('Bravo claim', noteMark(2)) +
        PAGE_BREAK +
        para('See note ') +
        para('Closing words'),
      ['First note text', 'Second note text that goes on for a while'],
    ),
  )
  let session = await openDoc(docPath)
  try {
    let { app, page } = session
    const crossRef = page.locator('.crossref-dialog')
    const refType = crossRef.locator('select[name="ref-type"]')
    const refContent = crossRef.locator('select[name="ref-content"]')
    const openCrossRef = async () => {
      await caretIn(page, 'See note')
      await clickMenu(app, ['Insert', 'Cross-reference…'])
      await expect(crossRef).toBeVisible()
    }
    const fields = () => page.locator('.doc-page [data-instr-field^="NOTEREF"]')
    const marks = () => page.locator('.doc-page sup.doc-note-ref').allTextContents()

    await openCrossRef()
    // numbered items need the list numbers cached, which is still not there: not offered
    await expect(crossRef.locator('option[value="numbered"]')).toHaveCount(0)
    await expect(crossRef.locator('option[value="footnote"]')).not.toHaveAttribute('disabled', '')
    await expect(crossRef.locator('option[value="endnote"]')).not.toHaveAttribute('disabled', '')
    await refType.selectOption('endnote')
    await expect(crossRef.locator('.bookmark-row')).toHaveCount(0)
    await expect(crossRef.locator('button[name="insert"]')).toBeDisabled()
    await refType.selectOption('footnote')
    await expect(refContent.locator('option')).toHaveText([
      'Footnote number',
      'Page Number',
      'Above/below',
      'Footnote number (formatted)',
    ])
    await expect(refContent).toHaveValue('noteNumber')
    await expect(crossRef.locator('.bookmark-row')).toHaveText([
      '1 First note text',
      '2 Second note text that goes on for a while',
    ])
    await crossRef.locator('.bookmark-row').nth(1).click()
    await page.screenshot({ path: test.info().outputPath('crossref-footnote.png') })
    await crossRef.locator('button[name="insert"]').click()
    await expect(crossRef).toBeHidden()
    await expect(fields()).toHaveText(['2'])

    // where the note is from here, then its number as a reference mark looks
    await openCrossRef()
    await refType.selectOption('footnote')
    await refContent.selectOption('aboveBelow')
    await crossRef.locator('.bookmark-row').nth(1).click()
    await crossRef.locator('button[name="insert"]').click()
    await openCrossRef()
    await refType.selectOption('footnote')
    await refContent.selectOption('noteNumberFormatted')
    await crossRef.locator('.bookmark-row').nth(1).click()
    await crossRef.locator('input[name="ref-hyperlink"]').uncheck()
    await crossRef.locator('button[name="insert"]').click()
    await expect(fields()).toHaveText(['2', 'above', '2'])
    // the formatted number is raised as a reference mark is; the plain one is not
    const raised = (nth: number) =>
      fields()
        .nth(nth)
        .evaluate((el) =>
          [el, el.parentElement, ...Array.from(el.querySelectorAll('*'))].some(
            (node) => node && getComputedStyle(node).verticalAlign === 'super',
          ),
        )
    expect(await raised(0)).toBe(false)
    expect(await raised(1)).toBe(false)
    expect(await raised(2)).toBe(true)
    await page.screenshot({ path: test.info().outputPath('crossref-inserted.png') })

    // a note added before the one referred to: its number moves, and the references with it
    await caretIn(page, 'Opening words')
    await clickMenu(app, ['Insert', 'Footnote…'])
    const footnote = page.locator('.footnote-dialog')
    await expect(footnote).toBeVisible()
    await footnote.locator('button.primary').click()
    await expect(footnote).toBeHidden()
    const prompt = page.locator('.modal textarea.prompt-textarea')
    await prompt.fill('An earlier note')
    await page.keyboard.press('ControlOrMeta+Enter')
    await expect(prompt).toBeHidden()
    await expect.poll(marks).toEqual(['1', '2', '3'])
    await expect(fields()).toHaveText(['3', 'above', '3'])

    // a reference inserted as a hyperlink goes to the note's mark on Cmd / Ctrl + click
    const target = page.locator('.doc-page sup.doc-note-ref').nth(2)
    await fields().first().scrollIntoViewIfNeeded()
    await page.evaluate(() => document.querySelector('main.editor-scroll')?.scrollBy(0, 100_000))
    await expect(target).not.toBeInViewport()
    await fields()
      .first()
      .click({ modifiers: ['ControlOrMeta'] })
    await expect(target).toBeInViewport()
    // the caret went with it: the mark is what is selected
    expect(
      await page.evaluate(() => {
        const { selection } = (window as unknown as EditorWindow).__aidocs.editor.state
        return selection.$from.nodeAfter?.type.name
      }),
    ).toBe('docNoteRef')

    await saveFromMenu(app)
    await expect
      .poll(async () => (await documentXml(docPath)).includes('NOTEREF'), { timeout: 30_000 })
      .toBe(true)
    const xml = await documentXml(docPath)
    const name = /NOTEREF (_Ref\d+) \\h /.exec(xml)?.[1] ?? ''
    expect(name).not.toBe('')
    // each reference is saved with the note's number as its result, none left for Word to update
    expect(xml).not.toContain('w:dirty')
    // one bookmark, around the run of the note's reference and nothing else
    expect(xml.match(new RegExp(`w:name="${name}"`, 'g'))).toHaveLength(1)
    expect(xml).toMatch(
      new RegExp(
        `<w:bookmarkStart w:id="(\\d+)" w:name="${name}"/><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r><w:bookmarkEnd w:id="\\1"/>`,
      ),
    )
    const instructions = [...xml.matchAll(/<w:instrText[^>]*>([^<]*)<\/w:instrText>/g)].map(
      (m) => m[1],
    )
    expect(instructions).toEqual([
      ` NOTEREF ${name} \\h `,
      ` NOTEREF ${name} \\p \\h `,
      ` NOTEREF ${name} \\f `,
    ])
    // each result is cached with the number the note has now
    expect(xml).toMatch(
      new RegExp(
        `NOTEREF ${name} \\\\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r>(?:<w:rPr>.*?</w:rPr>)?<w:t[^>]*>3</w:t>`,
      ),
    )

    // read again from the file: the fields show, the bookmark is found and reused
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    session = await openDoc(docPath)
    ;({ app, page } = session)
    const again = page.locator('.crossref-dialog')
    await expect(page.locator('.doc-page [data-instr-field^="NOTEREF"]')).toHaveText([
      '3',
      'above',
      '3',
    ])
    await caretIn(page, 'Closing words')
    await clickMenu(app, ['Insert', 'Cross-reference…'])
    await again.locator('select[name="ref-type"]').selectOption('footnote')
    await again.locator('.bookmark-row').nth(2).click()
    await again.locator('button[name="insert"]').click()
    await expect(again).toBeHidden()
    await expect(page.locator('.doc-page [data-instr-field^="NOTEREF"]')).toHaveCount(4)
    await saveFromMenu(app)
    await expect
      .poll(async () => (await documentXml(docPath)).split(`NOTEREF ${name} `).length - 1, {
        timeout: 30_000,
      })
      .toBe(4)
    expect((await documentXml(docPath)).match(/<w:bookmarkStart\b/g)).toHaveLength(1)
    expect(await documentXml(docPath)).not.toContain('w:dirty')
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
