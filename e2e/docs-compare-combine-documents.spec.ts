/**
 * Tools > Track Changes > Compare Documents and Tools > Combine Documents open
 * one dialog: an original and a revised document, the name the changes carry,
 * the comparison settings, and where the changes are shown. Shown in a new
 * document, the whole comparison (body, table cells, header, footnote,
 * formatting, a moved paragraph) opens as an untitled redline that saves as
 * real tracked changes; shown in the open document, the text differences are
 * written into it as one undoable edit.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { existsSync, mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const REVISER = 'Rae Visor'

const ORIGINAL = [
  'The quick fox jumps over the lazy dog.',
  'This clause goes away.',
  'Payment is due in thirty days.',
  'Signed by both parties.',
]
const REVISED = [
  'The quick brown fox jumps over the lazy dog.',
  'Payment is due in sixty days.',
  'Signed by both parties.',
  'A new closing clause.',
]

/** a document whose core properties name who saved it last */
async function docxSavedBy(paragraphs: string[], author: string): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await docxWith(paragraphs.map(para).join('')))
  zip.file(
    'docProps/core.xml',
    '<?xml version="1.0" encoding="UTF-8"?>' +
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">' +
      `<dc:creator>Someone Else</dc:creator><cp:lastModifiedBy>${author}</cp:lastModifiedBy>` +
      '</cp:coreProperties>',
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

/** the same package, its core properties naming who saved it last */
async function savedBy(docx: Buffer, author: string): Promise<Buffer> {
  const zip = await JSZip.loadAsync(docx)
  zip.file(
    'docProps/core.xml',
    '<?xml version="1.0" encoding="UTF-8"?>' +
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">' +
      `<cp:lastModifiedBy>${author}</cp:lastModifiedBy></cp:coreProperties>`,
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

const pickFile = (app: ElectronApplication, file: string) =>
  app.evaluate(({ dialog }, file) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [file] })
  }, file)

const partXml = async (path: string, part: string) =>
  (await (await JSZip.loadAsync(readFileSync(path))).file(part)?.async('string')) ?? ''
const documentXml = (path: string) => partXml(path, 'word/document.xml')

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const run = (text: string, rPr = '') =>
  `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`

/** a contract with a header, a footnote and a one-row table */
async function contract(parts: {
  body: string[]
  moved: string
  movedLast: boolean
  due: string
  bold: boolean
  price: string
  header: string
  note: string
  author?: string
}): Promise<Buffer> {
  const moved = `<w:p>${run(parts.moved)}</w:p>`
  const body =
    parts.body.map((text) => `<w:p>${run(text)}</w:p>`).join('') +
    (parts.movedLast ? '' : moved) +
    `<w:p>${run(`Payment is due in ${parts.due} days.`)}<w:r><w:footnoteReference w:id="1"/></w:r></w:p>` +
    `<w:p>${run('This line turns bold.', parts.bold ? '<w:b/>' : '')}</w:p>` +
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
    `<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p>${run('Price')}</w:p></w:tc>` +
    `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p>${run(parts.price)}</w:p></w:tc></w:tr></w:tbl>` +
    `<w:p>${run('Signed by both parties.')}</w:p>` +
    (parts.movedLast ? moved : '') +
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
  const zip = await JSZip.loadAsync(await docxWith(body))
  const types = await zip.file('[Content_Types].xml')!.async('string')
  zip.file(
    '[Content_Types].xml',
    types.replace(
      '</Types>',
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' +
        '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>' +
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>',
    ),
  )
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
  zip.file(
    'word/_rels/document.xml.rels',
    rels.replace(
      '</Relationships>',
      '<Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>' +
        '<Relationship Id="rIdNotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>',
    ),
  )
  zip.file(
    'word/header1.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:hdr ${W}><w:p>${run(parts.header)}</w:p></w:hdr>`,
  )
  zip.file(
    'word/footnotes.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:footnotes ${W}>` +
      '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
      '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
      `<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r>${run(` ${parts.note}`)}</w:p></w:footnote></w:footnotes>`,
  )
  if (parts.author)
    zip.file(
      'docProps/core.xml',
      '<?xml version="1.0" encoding="UTF-8"?>' +
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">' +
        `<cp:lastModifiedBy>${parts.author}</cp:lastModifiedBy></cp:coreProperties>`,
    )
  return zip.generateAsync({ type: 'nodebuffer' })
}

const MOVED = 'A clause that moves to the end of the contract.'
const contractOriginal = () =>
  contract({
    body: ['The quick fox jumps over the lazy dog.'],
    moved: MOVED,
    movedLast: false,
    due: 'thirty',
    bold: false,
    price: 'ten dollars',
    header: 'Draft contract',
    note: 'See the first edition.',
  })
const contractRevised = () =>
  contract({
    body: ['The quick brown fox jumps over the lazy dog.'],
    moved: MOVED,
    movedLast: true,
    due: 'sixty',
    bold: true,
    price: 'twelve dollars',
    header: 'Final contract',
    note: 'See the second edition.',
    author: REVISER,
  })

/** the dialog, with the revised document picked through its Browse button */
async function openDialog(app: ElectronApplication, page: Page, menu: string[], file: string) {
  await pickFile(app, file)
  await clickMenu(app, menu)
  const dialog = page.locator('.cmp-dialog')
  await expect(dialog).toBeVisible()
  const ok = dialog.getByRole('button', { name: 'OK', exact: true })
  await expect(ok).toBeDisabled()
  await expect(dialog.locator('select[name="compare-original"]')).toHaveValue('current')
  await expect(dialog.locator('button[name="compare-swap"]')).toBeDisabled()
  // a comparison takes the revised side's name only: the original's box is grey as in Word. Combining,
  // it names the original's own changes, which are not renamed while they are written into the open document
  await expect(dialog.locator('input[name="compare-label-original"]')).toBeDisabled()
  await dialog.locator('button[name="browse-revised"]').click()
  await expect(dialog.locator('select[name="compare-revised"] option:checked')).toHaveText(
    'revised.docx',
  )
  await expect(dialog.locator('input[name="compare-label-revised"]')).toHaveValue(REVISER)
  await expect(ok).toBeEnabled()
  await dialog.locator('button[name="compare-options"]').click()
  return { dialog, ok }
}

/** the tab the comparison opened: the docs page, other than the one it was run from, that shows changes */
async function resultPage(app: ElectronApplication, from: Page): Promise<Page> {
  let found: Page | undefined
  await expect
    .poll(
      async () => {
        for (const candidate of app.windows()) {
          if (candidate === from) continue
          const href = await candidate.evaluate(() => window.location.href).catch(() => '')
          if (!href.includes('://docs/')) continue
          const marks = await candidate
            .locator('.doc-page .doc-ins')
            .count()
            .catch(() => 0)
          if (marks > 0) found = candidate
        }
        return !!found
      },
      { timeout: 60_000 },
    )
    .toBe(true)
  return found!
}

const saveAs = async (app: ElectronApplication, out: string) => {
  await app.evaluate(({ dialog }, out) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath: out })
  }, out)
  await clickMenu(app, ['File', 'Save As…'])
  await expect.poll(() => existsSync(out), { timeout: 30_000 }).toBe(true)
}

const COMPARE_MENU = ['Tools', 'Track Changes', 'Compare Documents…']
const NEW_ONLY = ['moves', 'comments', 'headersFooters', 'notes', 'textBoxes', 'fields']

async function openReviewingPane(page: Page) {
  const toggle = page
    .locator('[data-ribbon-body]')
    .getByRole('button', { name: 'Reviewing', exact: true })
  // a click on the tab that is already showing folds the ribbon away
  if (!(await toggle.isVisible()))
    await page.locator('.ribbon-tab', { hasText: /^Review$/ }).click()
  await toggle.click()
  const pane = page.locator('.review-pane')
  await expect(pane).toBeVisible()
  return pane
}

test.describe('docs Compare Documents and Combine Documents', () => {
  let dir: string
  let docPath: string
  let revisedPath: string

  test.beforeEach(async () => {
    dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-combine-')))
    docPath = join(dir, 'contract.docx')
    revisedPath = join(dir, 'revised.docx')
    writeFileSync(docPath, await docxWith(ORIGINAL.map(para).join('')))
    writeFileSync(revisedPath, await docxSavedBy(REVISED, REVISER))
  })

  test.afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  test('the revised document comes in as tracked changes that are reviewed, saved and read back', async () => {
    test.setTimeout(240_000)
    let launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-combine-documents',
      openFile: docPath,
    })
    try {
      const { app } = launched
      const page = await waitForPageWithUrl(app, '://docs/')
      const paragraph = (text: string) => page.locator('.doc-page p', { hasText: text }).first()
      await expect(paragraph('Signed by both parties.')).toBeVisible({ timeout: 30_000 })
      const inserted = page.locator('.doc-page .doc-ins')
      const deleted = page.locator('.doc-page .doc-del')
      const status = page.locator('.status-msg')

      const { dialog, ok } = await openDialog(
        app,
        page,
        ['Tools', 'Combine Documents…'],
        revisedPath,
      )
      await expect(dialog.locator('h2')).toHaveText('Combine Documents')
      // Combine writes into the open document unless told otherwise: the text, formatting and table settings apply there
      const target = (value: string) =>
        dialog.locator(`input[name="compare-target"][value="${value}"]`)
      await expect(target('original')).toBeChecked()
      await expect(target('revised')).toBeDisabled()
      for (const name of NEW_ONLY)
        await expect(dialog.locator(`input[name="compare-${name}"]`)).toBeDisabled()
      for (const name of ['formatting', 'caseChanges', 'whiteSpace', 'tables'])
        await expect(dialog.locator(`input[name="compare-${name}"]`)).toBeChecked()
      // swapped, the open document is the revised one and the changes can go into it
      await dialog.locator('button[name="compare-swap"]').click()
      await expect(dialog.locator('select[name="compare-revised"]')).toHaveValue('current')
      await expect(target('original')).toBeDisabled()
      await expect(target('revised')).toBeEnabled()
      await expect(target('new')).toBeChecked()
      await dialog.locator('button[name="compare-swap"]').click()
      await expect(target('original')).toBeChecked()
      await ok.click()
      await expect(dialog).toHaveCount(0)
      await expect(status).toContainText('Combined with revised.docx: 3 inserted, 2 deleted, ')
      await expect(status).toContainText(/ \d+ formatting changes, marked as changes by /)
      await expect(status).toContainText(`marked as changes by ${REVISER}.`)
      await expect(inserted.filter({ hasText: 'brown' })).toHaveCount(1)
      await expect(deleted.filter({ hasText: 'thirty' })).toHaveCount(1)
      await expect(inserted.filter({ hasText: 'sixty' })).toHaveCount(1)
      await expect(deleted.filter({ hasText: 'This clause goes away.' })).toHaveCount(1)
      await expect(inserted.filter({ hasText: 'A new closing clause.' })).toHaveCount(1)

      // one undo takes the whole combine back
      await paragraph('Signed by both parties.').click()
      await clickMenu(app, ['Edit', 'Undo'])
      await expect(inserted).toHaveCount(0)
      await expect(deleted).toHaveCount(0)
      await expect(page.locator('.doc-page p', { hasText: 'A new closing clause.' })).toHaveCount(0)

      // Review > Compare > Combine Documents… is the same command
      await page.locator('.ribbon-tab', { hasText: /^Review$/ }).click()
      await page
        .locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: 'Compare' })
        .click()
      await page.locator('.layout-menu button', { hasText: 'Combine Documents…' }).click()
      await page.locator('.cmp-dialog button[name="browse-revised"]').click()
      await page.locator('.cmp-dialog').getByRole('button', { name: 'OK', exact: true }).click()
      await expect(inserted.filter({ hasText: 'A new closing clause.' })).toHaveCount(1)

      const pane = await openReviewingPane(page)
      const cards = pane.locator('.review-row')
      await expect(cards).toHaveCount(5)
      await expect(pane.locator('.review-pane-author')).toHaveText([REVISER])
      await expect(cards.nth(0)).toContainText('brown')
      await expect(cards.nth(1)).toContainText('This clause goes away.')
      await expect(cards.nth(2)).toContainText('thirty')
      await expect(cards.nth(3)).toContainText('sixty')
      await expect(cards.nth(4)).toContainText('A new closing clause.')

      // accept the deletion of "thirty", reject the new closing paragraph
      await cards.nth(2).click()
      await expect(cards.nth(2)).toHaveClass(/active/)
      await cards
        .nth(2)
        .getByRole('button', { name: /Accept/ })
        .click()
      await expect(cards).toHaveCount(4)
      await expect(page.locator('.doc-page p', { hasText: 'thirty' })).toHaveCount(0)
      await cards.nth(3).click()
      await expect(cards.nth(3)).toHaveClass(/active/)
      await cards
        .nth(3)
        .getByRole('button', { name: /Reject/ })
        .click()
      await expect(cards).toHaveCount(3)
      await expect(page.locator('.doc-page p', { hasText: 'A new closing clause.' })).toHaveCount(0)

      await saveFromMenu(app)
      const by = `w:author="${REVISER}"`
      await expect
        .poll(
          async () => {
            const xml = await documentXml(docPath)
            return [
              new RegExp(
                `<w:ins [^>]*${by}[^>]*><w:r>(?:<w:rPr>.*?</w:rPr>)?<w:t[^>]*>brown </w:t>`,
              ),
              new RegExp(
                `<w:ins [^>]*${by}[^>]*><w:r>(?:<w:rPr>.*?</w:rPr>)?<w:t[^>]*>sixty</w:t>`,
              ),
              new RegExp(
                `<w:del [^>]*${by}[^>]*><w:r>(?:<w:rPr>.*?</w:rPr>)?<w:delText[^>]*>This clause goes away\\.</w:delText>`,
              ),
            ].map((re) => re.test(xml))
          },
          { timeout: 30_000 },
        )
        .toEqual([true, true, true])
      const saved = await documentXml(docPath)
      expect(saved).not.toContain('thirty')
      expect(saved).not.toContain('A new closing clause.')
      expect(saved.match(/<w:ins /g)).toHaveLength(2)

      await closeAndSaveVideo(launched)
      launched = await launchShell({
        onboardingSeen: true,
        videoDir: 'docs-combine-documents-reopen',
        openFile: docPath,
      })
      const reopened = await waitForPageWithUrl(launched.app, '://docs/')
      await expect(
        reopened.locator('.doc-page p', { hasText: 'Signed by both parties.' }).first(),
      ).toBeVisible({ timeout: 30_000 })
      await expect(reopened.locator('.doc-page .doc-ins').filter({ hasText: 'brown' })).toHaveCount(
        1,
      )
      await expect(reopened.locator('.doc-page .doc-ins').filter({ hasText: 'sixty' })).toHaveCount(
        1,
      )
      await expect(
        reopened.locator('.doc-page .doc-del').filter({ hasText: 'This clause goes away.' }),
      ).toHaveCount(1)
      const reopenedPane = await openReviewingPane(reopened)
      await expect(reopenedPane.locator('.review-pane-author')).toHaveText([REVISER])
      await expect(reopenedPane.locator('.review-row').first()).toContainText('brown')
    } finally {
      await closeAndSaveVideo(launched)
    }
  })

  test('a document that shares no paragraph is not combined, and the status bar says so', async () => {
    test.setTimeout(120_000)
    writeFileSync(
      revisedPath,
      await docxSavedBy(['Recipe for lentil soup.', 'Simmer gently until soft.'], REVISER),
    )
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-combine-documents-unrelated',
      openFile: docPath,
    })
    try {
      const { app } = launched
      const page = await waitForPageWithUrl(app, '://docs/')
      await expect(
        page.locator('.doc-page p', { hasText: 'Signed by both parties.' }).first(),
      ).toBeVisible({ timeout: 30_000 })
      const { ok } = await openDialog(app, page, ['Tools', 'Combine Documents…'], revisedPath)
      await ok.click()
      await expect(page.locator('.status-msg')).toContainText(
        'revised.docx shares no paragraph with this document; nothing was combined.',
      )
      await expect(page.locator('.doc-page .doc-ins, .doc-page .doc-del')).toHaveCount(0)
      await expect(page.locator('.doc-page p')).toHaveText(ORIGINAL)
    } finally {
      await closeAndSaveVideo(launched)
    }
  })

  test('Compare Documents shows the whole comparison in a new document that saves as tracked changes', async () => {
    test.setTimeout(300_000)
    writeFileSync(docPath, await contractOriginal())
    writeFileSync(revisedPath, await contractRevised())
    const outPath = join(dir, 'redline.docx')
    let launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-compare-documents',
      openFile: docPath,
    })
    try {
      const { app } = launched
      const page = await waitForPageWithUrl(app, '://docs/')
      await expect(
        page.locator('.doc-page p', { hasText: 'Signed by both parties.' }).first(),
      ).toBeVisible({ timeout: 30_000 })

      const { dialog, ok } = await openDialog(app, page, COMPARE_MENU, revisedPath)
      await expect(dialog.locator('h2')).toHaveText('Compare Documents')
      await expect(dialog.locator('input[name="compare-target"][value="new"]')).toBeChecked()
      await expect(dialog.locator('input[name="compare-level"][value="word"]')).toBeChecked()
      for (const name of [...NEW_ONLY, 'caseChanges', 'whiteSpace', 'tables']) {
        await expect(dialog.locator(`input[name="compare-${name}"]`)).toBeEnabled()
        await expect(dialog.locator(`input[name="compare-${name}"]`)).toBeChecked()
      }
      await dialog.locator('input[name="compare-label-revised"]').fill('Opposing Counsel')
      await ok.click()
      await expect(page.locator('.status-msg')).toContainText(
        'Comparison opened in a new document:',
      )
      // the document the comparison was run from is untouched
      await expect(page.locator('.doc-page .doc-ins, .doc-page .doc-del')).toHaveCount(0)

      const result = await resultPage(app, page)
      const inserted = result.locator('.doc-ins')
      const deleted = result.locator('.doc-del')
      await expect(inserted.filter({ hasText: 'brown' })).toHaveCount(1)
      await expect(deleted.filter({ hasText: 'thirty' })).toHaveCount(1)
      await expect(inserted.filter({ hasText: 'sixty' })).toHaveCount(1)
      await expect(inserted.filter({ hasText: 'twelve' })).toHaveCount(1)
      await expect(deleted.filter({ hasText: 'ten' })).toHaveCount(1)
      await expect(inserted.filter({ hasText: 'Final' }).first()).toBeVisible()
      await expect(deleted.filter({ hasText: 'Draft' }).first()).toBeVisible()
      // the note area of a page shows a note as it reads with its changes accepted;
      // the changes themselves are listed in the Reviewing pane
      await expect(result.locator('.page-note-text')).toHaveText('See the second edition.')
      const resultPane = await openReviewingPane(result)
      for (const text of ['brown', 'Final', 'Draft', 'second', 'first', 'twelve', MOVED])
        await expect(resultPane.locator('.review-row', { hasText: text }).first()).toBeVisible()

      await saveAs(app, outPath)
      const by = 'w:author="Opposing Counsel"'
      const body = await documentXml(outPath)
      expect(body).toMatch(new RegExp(`<w:ins [^>]*${by}[^>]*><w:r>.*?<w:t[^>]*>brown </w:t>`))
      expect(body).toMatch(
        new RegExp(
          `<w:moveFrom [^>]*${by}[^>]*><w:r>.*?<w:t[^>]*>${MOVED.replace('.', '\\.')}</w:t>`,
        ),
      )
      expect(body).toMatch(
        new RegExp(
          `<w:moveTo [^>]*${by}[^>]*><w:r>.*?<w:t[^>]*>${MOVED.replace('.', '\\.')}</w:t>`,
        ),
      )
      expect(body).toMatch(
        new RegExp(
          `<w:rPrChange [^>]*${by}[^>]*>.*?</w:rPrChange></w:rPr><w:t[^>]*>This line turns bold\\.`,
        ),
      )
      const table = /<w:tbl>.*?<\/w:tbl>/s.exec(body)?.[0] ?? ''
      expect(table).toMatch(/<w:del [^>]*><w:r>.*?<w:delText[^>]*>ten<\/w:delText>/)
      expect(table).toMatch(/<w:ins [^>]*><w:r>.*?<w:t[^>]*>twelve<\/w:t>/)
      const header = await partXml(outPath, 'word/header1.xml')
      expect(header).toMatch(
        new RegExp(`<w:del [^>]*${by}[^>]*><w:r>.*?<w:delText[^>]*>Draft</w:delText>`),
      )
      expect(header).toMatch(new RegExp(`<w:ins [^>]*${by}[^>]*><w:r>.*?<w:t[^>]*>Final</w:t>`))
      const notes = await partXml(outPath, 'word/footnotes.xml')
      expect(notes).toMatch(
        new RegExp(`<w:del [^>]*${by}[^>]*><w:r>.*?<w:delText[^>]*>first</w:delText>`),
      )
      expect(notes).toMatch(new RegExp(`<w:ins [^>]*${by}[^>]*><w:r>.*?<w:t[^>]*>second</w:t>`))

      await closeAndSaveVideo(launched)
      launched = await launchShell({
        onboardingSeen: true,
        videoDir: 'docs-compare-documents-reopen',
        openFile: outPath,
      })
      const reopened = await waitForPageWithUrl(launched.app, '://docs/')
      await expect(
        reopened.locator('.doc-page p', { hasText: 'Signed by both parties.' }).first(),
      ).toBeVisible({ timeout: 30_000 })
      await expect(reopened.locator('.doc-ins').filter({ hasText: 'brown' })).toHaveCount(1)
      await expect(reopened.locator('.doc-ins').filter({ hasText: 'twelve' })).toHaveCount(1)
      await expect(reopened.locator('.doc-ins').filter({ hasText: 'Final' }).first()).toBeVisible()
      await expect(reopened.locator('.page-note-text')).toHaveText('See the second edition.')
      const pane = await openReviewingPane(reopened)
      await expect(pane.locator('.review-row', { hasText: 'second' }).first()).toBeVisible()
      await expect(pane.locator('.review-pane-author').first()).toHaveText('Opposing Counsel')
    } finally {
      await closeAndSaveVideo(launched)
    }
  })

  test('a comparison setting switched off leaves that part of the documents uncompared', async () => {
    test.setTimeout(240_000)
    writeFileSync(docPath, await contractOriginal())
    writeFileSync(revisedPath, await contractRevised())
    const outPath = join(dir, 'redline-text-only.docx')
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-compare-documents-settings',
      openFile: docPath,
    })
    try {
      const { app } = launched
      const page = await waitForPageWithUrl(app, '://docs/')
      await expect(
        page.locator('.doc-page p', { hasText: 'Signed by both parties.' }).first(),
      ).toBeVisible({ timeout: 30_000 })
      const { dialog, ok } = await openDialog(app, page, COMPARE_MENU, revisedPath)
      for (const name of ['moves', 'formatting', 'tables', 'headersFooters', 'notes'])
        await dialog.locator(`input[name="compare-${name}"]`).uncheck()
      await dialog.locator('input[name="compare-level"][value="char"]').check()
      await ok.click()

      const result = await resultPage(app, page)
      await expect(result.locator('.doc-ins').filter({ hasText: 'brown' })).toHaveCount(1)
      await saveAs(app, outPath)
      const body = await documentXml(outPath)
      expect(body).toMatch(/<w:ins [^>]*><w:r>.*?<w:t[^>]*>brown <\/w:t>/)
      // character level: "thirty" and "sixty" keep the letters they share
      expect(body).not.toContain('>thirty</w:delText>')
      expect(body).toMatch(/<w:delText[^>]*>[a-z]{1,3}<\/w:delText>/)
      expect(body).not.toContain('<w:moveFrom ')
      expect(body).toContain(`>${MOVED}</w:delText>`)
      expect(body).not.toContain('<w:rPrChange')
      const table = /<w:tbl>.*?<\/w:tbl>/s.exec(body)?.[0] ?? ''
      expect(table).toContain('twelve dollars')
      expect(table).not.toMatch(/<w:(?:ins|del) /)
      const header = await partXml(outPath, 'word/header1.xml')
      expect(header).toContain('Final contract')
      expect(header).not.toMatch(/<w:(?:ins|del) /)
      const notes = await partXml(outPath, 'word/footnotes.xml')
      expect(notes).toContain('See the second edition.')
      expect(notes).not.toMatch(/<w:(?:ins|del) /)
    } finally {
      await closeAndSaveVideo(launched)
    }
  })

  test('Compare writes a paragraph whose formatting changed as a tracked change that is listed, rejected and saved', async () => {
    test.setTimeout(240_000)
    const signed = (pPr: string) =>
      `<w:p><w:pPr>${pPr}</w:pPr>${run('Signed by both parties.')}</w:p>`
    writeFileSync(
      docPath,
      await docxWith(
        para('The quick fox jumps over the lazy dog.') + signed('<w:jc w:val="left"/>'),
      ),
    )
    writeFileSync(
      revisedPath,
      await savedBy(
        await docxWith(
          para('The quick brown fox jumps over the lazy dog.') +
            signed('<w:ind w:left="1440"/><w:jc w:val="center"/>'),
        ),
        REVISER,
      ),
    )
    const outPath = join(dir, 'redline.docx')
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-compare-paragraph-format',
      openFile: docPath,
    })
    try {
      const { app } = launched
      const page = await waitForPageWithUrl(app, '://docs/')
      await expect(
        page.locator('.doc-page p', { hasText: 'Signed by both parties.' }).first(),
      ).toBeVisible({ timeout: 30_000 })
      const { dialog, ok } = await openDialog(app, page, COMPARE_MENU, revisedPath)
      await expect(dialog.locator('input[name="compare-formatting"]')).toBeChecked()
      await ok.click()
      await expect(page.locator('.status-msg')).toContainText('1 formatting')

      const result = await resultPage(app, page)
      const paragraph = result
        .locator('.doc-page p', { hasText: 'Signed by both parties.' })
        .first()
      const align = () => paragraph.evaluate((el) => getComputedStyle(el).textAlign)
      await expect(paragraph).toHaveClass(/has-ppr-change/)
      expect(await align()).toBe('center')
      const pane = await openReviewingPane(result)
      const row = pane.locator('.review-row', { hasText: 'Paragraph formatted' })
      await expect(row).toHaveCount(1)
      await expect(pane.locator('.review-pane-author')).toHaveText([REVISER])
      if (process.env.CMP_SHOTS)
        await result.screenshot({ path: join(process.env.CMP_SHOTS, 'paragraph-format.png') })

      await saveAs(app, outPath)
      const record = new RegExp(
        `<w:pPr><w:ind w:left="1440"/><w:jc w:val="center"/><w:pPrChange w:id="\\d+" w:author="${REVISER}" w:date="[^"]+"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr>`,
      )
      expect(await documentXml(outPath)).toMatch(record)

      await row.click()
      await row.getByRole('button', { name: /Reject/ }).click()
      await expect(row).toHaveCount(0)
      await expect(result.locator('.doc-page .has-ppr-change')).toHaveCount(0)
      await expect.poll(align).toMatch(/^(left|start)$/)
      await saveFromMenu(app)
      await expect
        .poll(async () => (await documentXml(outPath)).includes('w:pPrChange'), { timeout: 30_000 })
        .toBe(false)
      const saved = await documentXml(outPath)
      expect(saved).toMatch(/<w:pPr><w:jc w:val="left"\/><\/w:pPr><w:r>.*?Signed by both parties\./)
      expect(saved).not.toContain('w:val="center"')
      expect(saved).toMatch(/<w:ins [^>]*><w:r>.*?<w:t[^>]*>brown <\/w:t>/)
    } finally {
      await closeAndSaveVideo(launched)
    }
  })

  const ABE = 'w:author="Abe Origin" w:date="2026-11-01T08:00:00Z"'
  const BEA = 'w:author="Bea Revise" w:date="2026-11-02T09:00:00Z"'
  const struck = (text: string) => `<w:r><w:delText xml:space="preserve">${text}</w:delText></w:r>`
  /** an original with a change Abe tracked, and a revised copy with Bea's and one untracked addition */
  const writeTrackedPair = async () => {
    writeFileSync(
      docPath,
      await docxWith(
        `<w:p>${run('The ')}<w:ins w:id="5" ${ABE}>${run('old ')}</w:ins>${run('cat sat on the mat.')}</w:p>` +
          para('Signed by both parties.'),
      ),
    )
    writeFileSync(
      revisedPath,
      await savedBy(
        await docxWith(
          `<w:p>${run('The cat sat ')}<w:del w:id="3" ${BEA}>${struck('on the mat')}</w:del>` +
            `<w:ins w:id="4" ${BEA}>${run('by the door')}</w:ins>${run('. Twice.')}</w:p>` +
            para('Signed by both parties.'),
        ),
        REVISER,
      ),
    )
  }
  const authorsIn = (pane: ReturnType<Page['locator']>) => async () =>
    (await pane.locator('.review-pane-author').allTextContents()).map((s) => s.trim()).sort()
  const expectAuthors = (xml: string) => {
    expect(xml).toMatch(new RegExp(`<w:ins w:id="\\d+" ${ABE}><w:r>.*?<w:t[^>]*>old </w:t>`))
    expect(xml).toMatch(
      new RegExp(`<w:del w:id="\\d+" ${BEA}><w:r>.*?<w:delText[^>]*>on the mat</w:delText>`),
    )
    expect(xml).toMatch(new RegExp(`<w:ins w:id="\\d+" ${BEA}><w:r>.*?<w:t[^>]*>by the door</w:t>`))
    expect(xml).toMatch(
      new RegExp(`<w:ins w:id="\\d+" w:author="${REVISER}"[^>]*><w:r>.*?<w:t[^>]*> Twice\\.</w:t>`),
    )
    expect(xml.match(/<w:(?:ins|del) /g)).toHaveLength(4)
  }

  test('Combine into a new document keeps the changes each document tracks under its author', async () => {
    test.setTimeout(240_000)
    await writeTrackedPair()
    const outPath = join(dir, 'combined.docx')
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-combine-authors-new',
      openFile: docPath,
    })
    try {
      const { app } = launched
      const page = await waitForPageWithUrl(app, '://docs/')
      await expect(
        page.locator('.doc-page p', { hasText: 'Signed by both parties.' }).first(),
      ).toBeVisible({ timeout: 30_000 })
      const { dialog, ok } = await openDialog(
        app,
        page,
        ['Tools', 'Combine Documents…'],
        revisedPath,
      )
      await dialog.locator('input[name="compare-target"][value="new"]').check()
      await ok.click()
      await expect(page.locator('.status-msg')).toContainText(
        'Comparison opened in a new document:',
      )
      // the open document still holds only its own change
      await expect(page.locator('.doc-page .doc-ins')).toHaveText(['old '])
      await expect(page.locator('.doc-page .doc-del')).toHaveCount(0)

      const result = await resultPage(app, page)
      await expect(result.locator('.doc-ins').filter({ hasText: 'by the door' })).toHaveCount(1)
      await expect(result.locator('.doc-del').filter({ hasText: 'on the mat' })).toHaveCount(1)
      await expect(result.locator('.doc-ins').filter({ hasText: 'old' })).toHaveCount(1)
      const pane = await openReviewingPane(result)
      await expect.poll(authorsIn(pane)).toEqual(['Abe Origin', 'Bea Revise', REVISER])
      await expect(pane.locator('.review-row')).toHaveCount(4)
      if (process.env.CMP_SHOTS)
        await result.screenshot({ path: join(process.env.CMP_SHOTS, 'combine-authors.png') })
      await saveAs(app, outPath)
      expectAuthors(await documentXml(outPath))
    } finally {
      await closeAndSaveVideo(launched)
    }
  })

  test("Combine into the open document brings the other document's changes in under their author", async () => {
    test.setTimeout(240_000)
    await writeTrackedPair()
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-combine-authors-open',
      openFile: docPath,
    })
    try {
      const { app } = launched
      const page = await waitForPageWithUrl(app, '://docs/')
      await expect(
        page.locator('.doc-page p', { hasText: 'Signed by both parties.' }).first(),
      ).toBeVisible({ timeout: 30_000 })
      const { dialog, ok } = await openDialog(
        app,
        page,
        ['Tools', 'Combine Documents…'],
        revisedPath,
      )
      await expect(dialog.locator('input[name="compare-target"][value="original"]')).toBeChecked()
      await ok.click()
      await expect(page.locator('.status-msg')).toContainText(
        'Combined with revised.docx: 1 inserted',
      )
      await expect(
        page.locator('.doc-page .doc-ins').filter({ hasText: 'by the door' }),
      ).toHaveCount(1)
      await expect(
        page.locator('.doc-page .doc-del').filter({ hasText: 'on the mat' }),
      ).toHaveCount(1)
      const pane = await openReviewingPane(page)
      await expect.poll(authorsIn(pane)).toEqual(['Abe Origin', 'Bea Revise', REVISER])
      await expect(pane.locator('.review-row')).toHaveCount(4)
      await saveFromMenu(app)
      await expect
        .poll(async () => (await documentXml(docPath)).includes('by the door'), { timeout: 30_000 })
        .toBe(true)
      expectAuthors(await documentXml(docPath))
    } finally {
      await closeAndSaveVideo(launched)
    }
  })

  const inlinePicture =
    '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="381000" cy="381000"/>' +
    '<wp:docPr id="7" name="Picture 7"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic><pic:nvPicPr><pic:cNvPr id="7" name="Picture 7"/><pic:cNvPicPr/></pic:nvPicPr>' +
    '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="381000" cy="381000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
    '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
  /** 1x1 blue PNG: another picture than the fixture's */
  const OTHER_PNG =
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPj/HwADBwIAMCbHYQAAAABJRU5ErkJggg=='
  const cell = (content: string) =>
    `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${content}</w:tc>`
  const priceTable = (pricePPr: string, extraRow: string) =>
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
    `<w:tr>${cell(para('Item'))}${cell(`<w:p>${pricePPr ? `<w:pPr>${pricePPr}</w:pPr>` : ''}${run('Price')}</w:p>`)}</w:tr>` +
    `<w:tr>${cell(para('Bolt'))}${cell(para('ten dollars'))}</w:tr>${extraRow}</w:tbl>`
  const SUPPLY = 'The supplier shall deliver the goods to the address of the customer.'
  const DAYS = 'Delivery takes place on working days only.'
  const INSPECT = 'The customer inspects the goods on arrival.'
  /** a pair that differs in paragraph formatting (body and cell), a table row, a picture and the place of a sentence */
  const writeLeftoversPair = async () => {
    const body = (parts: {
      sentences: string[]
      signedPPr: string
      pricePPr: string
      extraRow: string
    }) =>
      para(parts.sentences.join(' ')) +
      `<w:p><w:pPr>${parts.signedPPr}</w:pPr>${run('Signed by both parties.')}</w:p>` +
      priceTable(parts.pricePPr, parts.extraRow) +
      `<w:p>${run('Logo: ')}${inlinePicture}${run(' as shown.')}</w:p>`
    writeFileSync(
      docPath,
      await docxWith(
        body({
          sentences: [SUPPLY, DAYS, INSPECT],
          signedPPr: '<w:keepNext/><w:jc w:val="left"/>',
          pricePPr: '',
          extraRow: '',
        }),
      ),
    )
    const revised = await JSZip.loadAsync(
      await docxWith(
        body({
          sentences: [DAYS, SUPPLY, INSPECT],
          signedPPr: '<w:jc w:val="center"/>',
          pricePPr: '<w:jc w:val="right"/>',
          extraRow: `<w:tr>${cell(para('Nut'))}${cell(para('two dollars'))}</w:tr>`,
        }),
      ),
    )
    revised.file('word/media/image1.png', Buffer.from(OTHER_PNG, 'base64'))
    writeFileSync(
      revisedPath,
      await savedBy(await revised.generateAsync({ type: 'nodebuffer' }), REVISER),
    )
  }
  const mediaOf = async (path: string) => {
    const zip = await JSZip.loadAsync(readFileSync(path))
    const names = Object.keys(zip.files).filter((name) => name.startsWith('word/media/'))
    return Promise.all(names.map((name) => zip.file(name)!.async('base64')))
  }
  const BODY_RECORD = new RegExp(
    `<w:pPr><w:jc w:val="center"/><w:pPrChange w:id="\\d+" w:author="${REVISER}" w:date="[^"]+"><w:pPr><w:keepNext/><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr>`,
  )
  const CELL_RECORD =
    /<w:tc>(?:(?!<\/w:tc>).)*<w:pPr><w:jc w:val="right"\/><w:pPrChange [^>]*><w:pPr>(?:<\/w:pPr>|\/>)?<\/w:pPrChange><\/w:pPr><w:r>(?:(?!<\/w:tc>).)*Price/
  const ROW_RECORD = new RegExp(
    `<w:tr><w:trPr><w:ins w:id="\\d+" w:author="${REVISER}" w:date="[^"]+"/></w:trPr>`,
  )

  test('Compare into the open document tracks formatting, a table row and a replaced picture; they save, reject and reopen', async () => {
    test.setTimeout(300_000)
    await writeLeftoversPair()
    let launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-compare-leftovers-open',
      openFile: docPath,
    })
    try {
      let { app } = launched
      let page = await waitForPageWithUrl(app, '://docs/')
      const signedIn = (p: Page) =>
        p.locator('.doc-page p', { hasText: 'Signed by both parties.' }).first()
      await expect(signedIn(page)).toBeVisible({ timeout: 30_000 })
      const { dialog, ok } = await openDialog(app, page, COMPARE_MENU, revisedPath)
      await dialog.locator('input[name="compare-target"][value="original"]').check()
      // written into the open document, formatting is compared too
      await expect(dialog.locator('input[name="compare-formatting"]')).toBeEnabled()
      await expect(dialog.locator('input[name="compare-formatting"]')).toBeChecked()
      await expect(dialog.locator('input[name="compare-moves"]')).toBeDisabled()
      await ok.click()
      await expect(page.locator('.status-msg')).toContainText('Combined with revised.docx')
      // the formatting changes are counted beside the insertions and deletions
      await expect(page.locator('.status-msg')).toContainText(/ [1-9]\d* formatting changes, /)

      const align = (p: Page) => () => signedIn(p).evaluate((el) => getComputedStyle(el).textAlign)
      await expect(signedIn(page)).toHaveClass(/has-ppr-change/)
      await expect.poll(align(page)).toBe('center')
      await expect(page.locator('.doc-page table tr')).toHaveCount(3)
      // a cell paragraph carries the record and is listed; the change bar is drawn for body paragraphs
      await expect(page.locator('.doc-page td p', { hasText: 'Price' })).toHaveAttribute(
        'data-ppr-change',
        'true',
      )
      let pane = await openReviewingPane(page)
      await expect(pane.locator('.review-row', { hasText: 'Paragraph formatted' })).toHaveCount(2)
      await expect(pane.locator('.review-pane-author')).toHaveText([REVISER])
      if (process.env.CMP_SHOTS)
        await page.screenshot({ path: join(process.env.CMP_SHOTS, 'open-document.png') })

      await saveFromMenu(app)
      await expect
        .poll(async () => ROW_RECORD.test(await documentXml(docPath)), { timeout: 30_000 })
        .toBe(true)
      let xml = await documentXml(docPath)
      expect(xml).toMatch(BODY_RECORD)
      expect(xml).toMatch(CELL_RECORD)
      expect(xml).toMatch(/<w:del [^>]*><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:drawing>/)
      expect(xml).toMatch(/<w:ins [^>]*><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:drawing>/)
      const media = await mediaOf(docPath)
      expect(media).toHaveLength(2)
      expect(media).toContain(OTHER_PNG)

      // rejecting the paragraph's change brings back its alignment and keep with next
      const signedRow = pane
        .locator('.review-row', { hasText: 'Paragraph formatted' })
        .filter({ hasText: 'Signed by both parties.' })
      await signedRow.click()
      await signedRow.getByRole('button', { name: /Reject/ }).click()
      await expect(signedIn(page)).not.toHaveClass(/has-ppr-change/)
      await expect.poll(align(page)).toMatch(/^(left|start)$/)
      await saveFromMenu(app)
      await expect
        .poll(async () => BODY_RECORD.test(await documentXml(docPath)), { timeout: 30_000 })
        .toBe(false)
      xml = await documentXml(docPath)
      expect(xml).toMatch(
        /<w:pPr><w:keepNext\/><w:jc w:val="left"\/><\/w:pPr><w:r>(?:(?!<\/w:p>).)*Signed by both parties\./,
      )
      expect(xml).toMatch(CELL_RECORD)

      // reopened, what is still pending is listed again
      await closeAndSaveVideo(launched)
      launched = await launchShell({
        onboardingSeen: true,
        videoDir: 'docs-compare-leftovers-reopen',
        openFile: docPath,
      })
      app = launched.app
      page = await waitForPageWithUrl(app, '://docs/')
      await expect(signedIn(page)).toBeVisible({ timeout: 30_000 })
      await expect(page.locator('.doc-page table tr')).toHaveCount(3)
      // a cell paragraph carries the record and is listed; the change bar is drawn for body paragraphs
      await expect(page.locator('.doc-page td p', { hasText: 'Price' })).toHaveAttribute(
        'data-ppr-change',
        'true',
      )
      pane = await openReviewingPane(page)
      await expect(pane.locator('.review-row', { hasText: 'Paragraph formatted' })).toHaveCount(1)
      await expect(pane.locator('.review-pane-author')).toHaveText([REVISER])
    } finally {
      await closeAndSaveVideo(launched)
    }
  })

  test('Compare into a new document tracks a cell paragraph format, a table row, a replaced picture and a moved sentence', async () => {
    test.setTimeout(240_000)
    await writeLeftoversPair()
    const outPath = join(dir, 'redline.docx')
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-compare-leftovers-new',
      openFile: docPath,
    })
    try {
      const { app } = launched
      const page = await waitForPageWithUrl(app, '://docs/')
      await expect(
        page.locator('.doc-page p', { hasText: 'Signed by both parties.' }).first(),
      ).toBeVisible({ timeout: 30_000 })
      const { ok } = await openDialog(app, page, COMPARE_MENU, revisedPath)
      await ok.click()
      await expect(page.locator('.status-msg')).toContainText('2 formatting')
      await expect(page.locator('.status-msg')).toContainText('1 move')

      const result = await resultPage(app, page)
      const price = result.locator('.doc-page td p', { hasText: 'Price' })
      await expect(price).toHaveAttribute('data-ppr-change', 'true')
      await expect(result.locator('.doc-page table tr')).toHaveCount(3)
      const pane = await openReviewingPane(result)
      const formatted = pane.locator('.review-row', { hasText: 'Paragraph formatted' })
      await expect(formatted).toHaveCount(2)
      if (process.env.CMP_SHOTS)
        await result.screenshot({ path: join(process.env.CMP_SHOTS, 'new-document.png') })

      await saveAs(app, outPath)
      let xml = await documentXml(outPath)
      expect(xml).toMatch(BODY_RECORD)
      expect(xml).toMatch(CELL_RECORD)
      expect(xml).toMatch(ROW_RECORD)
      expect(xml.match(/<w:moveFrom /g)).toHaveLength(1)
      expect(xml.match(/<w:moveTo /g)).toHaveLength(1)
      expect(xml).toMatch(/<w:moveFrom [^>]*><w:r>(?:(?!<\/w:moveFrom>).)*Delivery takes place/)
      expect(xml).toMatch(/<w:del [^>]*><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:drawing>/)
      expect(xml).toMatch(/<w:ins [^>]*><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:drawing>/)
      expect(await mediaOf(outPath)).toHaveLength(2)

      // the cell paragraph's change is rejected from the pane and leaves the file
      const cellRow = formatted.filter({ hasText: 'Price' })
      await cellRow.click()
      await cellRow.getByRole('button', { name: /Reject/ }).click()
      await expect(price).not.toHaveAttribute('data-ppr-change', 'true')
      await saveFromMenu(app)
      await expect
        .poll(async () => CELL_RECORD.test(await documentXml(outPath)), { timeout: 30_000 })
        .toBe(false)
      xml = await documentXml(outPath)
      expect(xml).not.toContain('w:val="right"')
      expect(xml).toMatch(BODY_RECORD)
    } finally {
      await closeAndSaveVideo(launched)
    }
  })

  test('a table column only the revised document has is tracked as inserted cells, in a new document and in the open one', async () => {
    test.setTimeout(300_000)
    const grid = (rows: string[][]) =>
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>' +
      '<w:gridCol w:w="2000"/>'.repeat(rows[0].length) +
      `</w:tblGrid>${rows.map((cells) => `<w:tr>${cells.map((text) => cell(para(text))).join('')}</w:tr>`).join('')}</w:tbl>`
    const INTRO = 'The supplier lists the parts below.'
    const SIGNED = 'Signed by both parties.'
    writeFileSync(
      docPath,
      await docxWith(
        para(INTRO) +
          grid([
            ['Item', 'Price'],
            ['Bolt', 'ten dollars'],
          ]) +
          para(SIGNED),
      ),
    )
    writeFileSync(
      revisedPath,
      await savedBy(
        await docxWith(
          para(INTRO) +
            grid([
              ['Item', 'Stock', 'Price'],
              ['Bolt', 'many', 'ten dollars'],
            ]) +
            para(SIGNED),
        ),
        REVISER,
      ),
    )
    const cellsOf = (p: Page) =>
      p
        .locator('.doc-page table tr')
        .evaluateAll((rows) =>
          rows.map((row) =>
            Array.from(row.querySelectorAll('td, th'), (td) => (td.textContent ?? '').trim()),
          ),
        )
    const WIDE = [
      ['Item', 'Stock', 'Price'],
      ['Bolt', 'many', 'ten dollars'],
    ]
    const CELL_INS = new RegExp(`<w:cellIns w:id="\\d+" w:author="${REVISER}"`, 'g')
    const outPath = join(dir, 'columns.docx')
    let launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-compare-columns-new',
      openFile: docPath,
    })
    try {
      let { app } = launched
      let page = await waitForPageWithUrl(app, '://docs/')
      await expect(page.locator('.doc-page p', { hasText: SIGNED }).first()).toBeVisible({
        timeout: 30_000,
      })
      let { dialog, ok } = await openDialog(app, page, COMPARE_MENU, revisedPath)
      await ok.click()
      const result = await resultPage(app, page)
      await expect.poll(() => cellsOf(result)).toEqual(WIDE)
      await expect(result.locator('.doc-page table')).toHaveCount(1)
      await expect(result.locator('.doc-page td.cell-rev-ins')).toHaveText(['Stock', 'many'])
      if (process.env.CMP_SHOTS)
        await result.screenshot({ path: join(process.env.CMP_SHOTS, 'columns-new.png') })
      await saveAs(app, outPath)
      let xml = await documentXml(outPath)
      expect(xml.match(CELL_INS)).toHaveLength(2)
      expect(xml.match(/<w:gridCol /g)).toHaveLength(3)
      expect(xml).not.toMatch(/<w:(?:ins|del) [^>]*><w:tbl>/)

      // written into the open document instead
      await closeAndSaveVideo(launched)
      launched = await launchShell({
        onboardingSeen: true,
        videoDir: 'docs-compare-columns-open',
        openFile: docPath,
      })
      app = launched.app
      page = await waitForPageWithUrl(app, '://docs/')
      await expect(page.locator('.doc-page p', { hasText: SIGNED }).first()).toBeVisible({
        timeout: 30_000,
      })
      ;({ dialog, ok } = await openDialog(app, page, COMPARE_MENU, revisedPath))
      await dialog.locator('input[name="compare-target"][value="original"]').check()
      await ok.click()
      await expect(page.locator('.status-msg')).toContainText('Combined with revised.docx')
      await expect.poll(() => cellsOf(page)).toEqual(WIDE)
      await expect(page.locator('.doc-page td.cell-rev-ins')).toHaveText(['Stock', 'many'])
      if (process.env.CMP_SHOTS)
        await page.screenshot({ path: join(process.env.CMP_SHOTS, 'columns-open.png') })
      await saveFromMenu(app)
      await expect
        .poll(async () => (await documentXml(docPath)).match(CELL_INS)?.length ?? 0, {
          timeout: 30_000,
        })
        .toBe(2)

      // rejected, the column goes and the table is as it was
      await openReviewingPane(page)
      await page
        .locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: /^Reject$/ })
        .click()
      await page
        .locator('[data-rb-panel] button', { hasText: /^Reject All Changes$/ })
        .first()
        .click()
      await expect
        .poll(() => cellsOf(page))
        .toEqual([
          ['Item', 'Price'],
          ['Bolt', 'ten dollars'],
        ])
      await saveFromMenu(app)
      await expect
        .poll(async () => (await documentXml(docPath)).includes('w:cellIns'), { timeout: 30_000 })
        .toBe(false)
      xml = await documentXml(docPath)
      expect(xml.match(/<w:gridCol /g)).toHaveLength(2)
    } finally {
      await closeAndSaveVideo(launched)
    }
  })
  test('a column added beside merged cells is tracked as inserted cells and the merged cells stay merged, in a new document and in the open one', async () => {
    test.setTimeout(300_000)
    // a heading across the first two columns and a cell down the last column's two lower rows
    const merged = (text: string, width: number, tcPr: string) =>
      `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${tcPr}</w:tcPr>${para(text)}</w:tc>`
    const table = (extra: string[] | null) =>
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>' +
      '<w:gridCol w:w="2000"/>'.repeat(extra ? 4 : 3) +
      '</w:tblGrid>' +
      [
        [merged('Parts', 4000, '<w:gridSpan w:val="2"/>'), cell(para('Origin'))],
        [
          cell(para('Item')),
          cell(para('Price')),
          merged('Leeds', 2000, '<w:vMerge w:val="restart"/>'),
        ],
        [cell(para('Bolt')), cell(para('ten dollars')), merged('', 2000, '<w:vMerge/>')],
      ]
        .map((cells, r) => {
          const added = extra ? [cell(para(extra[r]))] : []
          return `<w:tr>${[...cells.slice(0, -1), ...added, cells[cells.length - 1]].join('')}</w:tr>`
        })
        .join('') +
      '</w:tbl>'
    const INTRO = 'The supplier lists the parts below.'
    const SIGNED = 'Signed by both parties.'
    writeFileSync(docPath, await docxWith(para(INTRO) + table(null) + para(SIGNED)))
    writeFileSync(
      revisedPath,
      await savedBy(
        await docxWith(para(INTRO) + table(['Kind', 'Stock', 'many']) + para(SIGNED)),
        REVISER,
      ),
    )
    /** each row's cells: text, columns covered, rows covered */
    const cellsOf = (p: Page) =>
      p
        .locator('.doc-page table tr')
        .evaluateAll((rows) =>
          rows.map((row) =>
            Array.from(row.querySelectorAll<HTMLTableCellElement>('td, th'), (td) => [
              (td.textContent ?? '').trim(),
              td.colSpan,
              td.rowSpan,
            ]),
          ),
        )
    const NARROW = [
      [
        ['Parts', 2, 1],
        ['Origin', 1, 1],
      ],
      [
        ['Item', 1, 1],
        ['Price', 1, 1],
        ['Leeds', 1, 2],
      ],
      [
        ['Bolt', 1, 1],
        ['ten dollars', 1, 1],
      ],
    ]
    const WIDE = [
      [NARROW[0][0], ['Kind', 1, 1], NARROW[0][1]],
      [NARROW[1][0], NARROW[1][1], ['Stock', 1, 1], NARROW[1][2]],
      [NARROW[2][0], NARROW[2][1], ['many', 1, 1]],
    ]
    const CELL_INS = new RegExp(`<w:cellIns w:id="\\d+" w:author="${REVISER}"`, 'g')
    const mergesIn = (xml: string) => [
      xml.match(/<w:gridSpan w:val="2"\/>/g)?.length ?? 0,
      xml.match(/<w:vMerge\b/g)?.length ?? 0,
    ]
    const outPath = join(dir, 'merged-columns.docx')
    let launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-compare-merged-columns-new',
      openFile: docPath,
    })
    try {
      let { app } = launched
      let page = await waitForPageWithUrl(app, '://docs/')
      await expect(page.locator('.doc-page p', { hasText: SIGNED }).first()).toBeVisible({
        timeout: 30_000,
      })
      let { dialog, ok } = await openDialog(app, page, COMPARE_MENU, revisedPath)
      await ok.click()
      const result = await resultPage(app, page)
      await expect.poll(() => cellsOf(result)).toEqual(WIDE)
      await expect(result.locator('.doc-page table')).toHaveCount(1)
      await expect(result.locator('.doc-page td.cell-rev-ins')).toHaveText([
        'Kind',
        'Stock',
        'many',
      ])
      await saveAs(app, outPath)
      let xml = await documentXml(outPath)
      expect(xml.match(CELL_INS)).toHaveLength(3)
      expect(xml.match(/<w:gridCol /g)).toHaveLength(4)
      expect(mergesIn(xml)).toEqual([1, 2])
      expect(xml).not.toMatch(/<w:(?:ins|del) [^>]*><w:tbl>/)

      // written into the open document instead
      await closeAndSaveVideo(launched)
      launched = await launchShell({
        onboardingSeen: true,
        videoDir: 'docs-compare-merged-columns-open',
        openFile: docPath,
      })
      app = launched.app
      page = await waitForPageWithUrl(app, '://docs/')
      await expect(page.locator('.doc-page p', { hasText: SIGNED }).first()).toBeVisible({
        timeout: 30_000,
      })
      ;({ dialog, ok } = await openDialog(app, page, COMPARE_MENU, revisedPath))
      await dialog.locator('input[name="compare-target"][value="original"]').check()
      await ok.click()
      await expect(page.locator('.status-msg')).toContainText('Combined with revised.docx')
      await expect.poll(() => cellsOf(page)).toEqual(WIDE)
      await expect(page.locator('.doc-page td.cell-rev-ins')).toHaveText(['Kind', 'Stock', 'many'])
      await saveFromMenu(app)
      await expect
        .poll(async () => (await documentXml(docPath)).match(CELL_INS)?.length ?? 0, {
          timeout: 30_000,
        })
        .toBe(3)
      expect(mergesIn(await documentXml(docPath))).toEqual([1, 2])

      // rejected, the column goes and the table is as it was, merged cells included
      await openReviewingPane(page)
      await page
        .locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: /^Reject$/ })
        .click()
      await page
        .locator('[data-rb-panel] button', { hasText: /^Reject All Changes$/ })
        .first()
        .click()
      await expect.poll(() => cellsOf(page)).toEqual(NARROW)
      await saveFromMenu(app)
      await expect
        .poll(async () => (await documentXml(docPath)).includes('w:cellIns'), { timeout: 30_000 })
        .toBe(false)
      xml = await documentXml(docPath)
      expect(xml.match(/<w:gridCol /g)).toHaveLength(3)
      expect(mergesIn(xml)).toEqual([1, 2])
    } finally {
      await closeAndSaveVideo(launched)
    }
  })

  test('a row and a column added in a table with a cell merged down are tracked in the open document, the merge running through the new row, and rejected back', async () => {
    test.setTimeout(240_000)
    const merged = (text: string, width: number, tcPr: string) =>
      `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${tcPr}</w:tcPr>${para(text)}</w:tc>`
    // a heading across the first two columns and a cell down the last column's lower rows
    const table = (extra: string[] | null) =>
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>' +
      '<w:gridCol w:w="2000"/>'.repeat(extra ? 4 : 3) +
      '</w:tblGrid>' +
      [
        [merged('Parts', 4000, '<w:gridSpan w:val="2"/>'), cell(para('Origin'))],
        [
          cell(para('Item')),
          cell(para('Price')),
          merged('Leeds', 2000, '<w:vMerge w:val="restart"/>'),
        ],
        ...(extra
          ? [[cell(para('Nut')), cell(para('two dollars')), merged('', 2000, '<w:vMerge/>')]]
          : []),
        [cell(para('Bolt')), cell(para('ten dollars')), merged('', 2000, '<w:vMerge/>')],
      ]
        .map((cells, r) => {
          const added = extra ? [cell(para(extra[r]))] : []
          return `<w:tr>${[...cells.slice(0, -1), ...added, cells[cells.length - 1]].join('')}</w:tr>`
        })
        .join('') +
      '</w:tbl>'
    const INTRO = 'The supplier lists the parts below.'
    const SIGNED = 'Signed by both parties.'
    writeFileSync(docPath, await docxWith(para(INTRO) + table(null) + para(SIGNED)))
    writeFileSync(
      revisedPath,
      await savedBy(
        await docxWith(para(INTRO) + table(['Kind', 'Stock', 'few', 'many']) + para(SIGNED)),
        REVISER,
      ),
    )
    const cellsOf = (p: Page) =>
      p
        .locator('.doc-page table tr')
        .evaluateAll((rows) =>
          rows.map((row) =>
            Array.from(row.querySelectorAll<HTMLTableCellElement>('td, th'), (td) => [
              (td.textContent ?? '').trim(),
              td.colSpan,
              td.rowSpan,
            ]),
          ),
        )
    const NARROW = [
      [
        ['Parts', 2, 1],
        ['Origin', 1, 1],
      ],
      [
        ['Item', 1, 1],
        ['Price', 1, 1],
        ['Leeds', 1, 2],
      ],
      [
        ['Bolt', 1, 1],
        ['ten dollars', 1, 1],
      ],
    ]
    // pending, the merged cell covers the new row too
    const PENDING = [
      [NARROW[0][0], ['Kind', 1, 1], NARROW[0][1]],
      [NARROW[1][0], NARROW[1][1], ['Stock', 1, 1], ['Leeds', 1, 3]],
      [
        ['Nut', 1, 1],
        ['two dollars', 1, 1],
        ['few', 1, 1],
      ],
      [NARROW[2][0], NARROW[2][1], ['many', 1, 1]],
    ]
    const count = (xml: string, re: RegExp) => xml.match(re)?.length ?? 0
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-compare-merged-rows-open',
      openFile: docPath,
    })
    try {
      const { app } = launched
      const page = await waitForPageWithUrl(app, '://docs/')
      await expect(page.locator('.doc-page p', { hasText: SIGNED }).first()).toBeVisible({
        timeout: 30_000,
      })
      const { dialog, ok } = await openDialog(app, page, COMPARE_MENU, revisedPath)
      await dialog.locator('input[name="compare-target"][value="original"]').check()
      await ok.click()
      await expect(page.locator('.status-msg')).toContainText('Combined with revised.docx')
      await expect.poll(() => cellsOf(page)).toEqual(PENDING)
      await saveFromMenu(app)
      await expect
        .poll(async () => count(await documentXml(docPath), /<w:cellIns /g), { timeout: 30_000 })
        .toBe(4)
      let xml = await documentXml(docPath)
      expect(count(xml, /<w:trPr><w:ins /g)).toBe(1)
      expect(count(xml, /<w:vMerge\b/g)).toBe(3)

      await openReviewingPane(page)
      await page
        .locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: /^Reject$/ })
        .click()
      await page
        .locator('[data-rb-panel] button', { hasText: /^Reject All Changes$/ })
        .first()
        .click()
      await expect.poll(() => cellsOf(page)).toEqual(NARROW)
      await saveFromMenu(app)
      await expect
        .poll(async () => (await documentXml(docPath)).includes('w:cellIns'), { timeout: 30_000 })
        .toBe(false)
      xml = await documentXml(docPath)
      expect(xml.match(/<w:gridCol /g)).toHaveLength(3)
      expect(xml.match(/<w:tr[ >]/g)).toHaveLength(3)
      expect(count(xml, /<w:vMerge\b/g)).toBe(2)
    } finally {
      await closeAndSaveVideo(launched)
    }
  })
})
