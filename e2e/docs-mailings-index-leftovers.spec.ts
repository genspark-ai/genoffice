/**
 * Mailings and References as Word has them: Preview Results works out the
 * rule fields for the record it shows, a SET that Word has run sits in an
 * editable paragraph and keeps its bookmark, a sheet of labels is saved as a
 * labels main document, the ribbon's Insert Index is the dialog's two-column
 * index, a page range entry lists its range, and a table of authorities puts
 * a sequence number before its pages.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-mailings-index-leftovers'

const part = async (file: string, name: string) =>
  (await (await JSZip.loadAsync(readFileSync(file))).file(name)?.async('string')) ?? ''
const documentXml = (file: string) => part(file, 'word/document.xml')

const fld = (type: string) => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`
const instr = (text: string) =>
  `<w:r><w:instrText xml:space="preserve"> ${text} </w:instrText></w:r>`
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const field = (code: string, result = '') =>
  fld('begin') + instr(code) + (result ? fld('separate') + result : '') + fld('end')

const tab = (page: Page, name: string) =>
  page.locator('.ribbon-tab', { hasText: new RegExp(`^${name}$`) }).click()
const big = (page: Page, name: string | RegExp) =>
  page.locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: name })

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

const stubDialogs = (app: ElectronApplication, list: string, out: string) =>
  app.evaluate(
    ({ dialog }, [list, out]) => {
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [list] })
      dialog.showSaveDialog = async () => ({ canceled: false, filePath: out })
    },
    [list, out],
  )

/** the text of each paragraph of a saved file that has any */
const paragraphs = async (file: string) =>
  [...(await documentXml(file)).matchAll(/<w:p[ >].*?<\/w:p>/g)]
    .map((p) => [...p[0].matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((t) => t[1]).join(''))
    .map((text) => text.replace(/\s+/g, ' ').trim())
    .filter(Boolean)

test('Preview Results works out the rule fields; a SET that Word has run is editable and keeps its bookmark', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-merge-preview-')))
  const docPath = join(dir, 'letter.docx')
  const ranSet = field(
    'SET Rate "5%"',
    '<w:bookmarkStart w:id="3" w:name="Rate"/>' + run('5%') + '<w:bookmarkEnd w:id="3"/>',
  )
  writeFileSync(
    docPath,
    await docxWith(
      para('Dear') +
        `<w:p>${run('Terms')}${ranSet}${run(' at ')}${field('REF Rate', run('5%'))}</w:p>` +
        para('Regards'),
    ),
  )
  const listPath = join(dir, 'people.csv')
  writeFileSync(listPath, 'Name,City\nAda,London\nAlan,Wilmslow\nGrace,London\nEdsger,Austin\n')
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
    // the paragraph of the SET is text like any other, not a locked block
    const terms = page.locator('.doc-page p', { hasText: 'Terms' }).first()
    await expect(terms).toBeVisible()
    await expect(terms.locator('[data-instr-field^="SET Rate"]')).toHaveCount(1)
    await expect(page.locator('.doc-page .doc-protected')).toHaveCount(0)

    await stubDialogs(app, listPath, outPath)
    await tab(page, 'Mailings')
    await big(page, 'Select Recipients').click()
    const record = page.locator('[data-ribbon-body] [data-testid="mailings-record-status"]')
    await expect(record).toHaveText(/Record 1 of 4/)

    const rules = big(page, 'Rules')
    const rule = async (name: RegExp) => {
      await rules.click()
      await page.locator('.layout-menu button', { hasText: name }).click()
    }
    const mergeField = async (name: string) => {
      await big(page, 'Insert Merge Field').click()
      await page.locator('.rb-field-item', { hasText: new RegExp(`^${name}$`) }).click()
    }
    // Dear «Name» no. <record>/<sequence>: <gift>, skipped for Wilmslow
    await greeting.click()
    await page.keyboard.press('End')
    await page.keyboard.type(' ')
    await mergeField('Name')
    await page.keyboard.type(' no. ')
    await rule(/^Merge Record #$/)
    await page.keyboard.type('/')
    await rule(/^Merge Sequence #$/)
    await page.keyboard.type(': ')
    await rule(/^Fill-in…$/)
    const dialog = page.locator('.rule-dialog')
    await dialog.locator('input[name="prompt"]').fill('Which gift?')
    await dialog.locator('input[name="value"]').fill('a card')
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await rule(/^Skip Record If…$/)
    await dialog.locator('select[name="field"]').selectOption('City')
    await dialog.locator('input[name="value"]').fill('Wilmslow')
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    // Regards, then the record after it
    await page.locator('.doc-page p', { hasText: 'Regards' }).first().click()
    await page.keyboard.press('End')
    await page.keyboard.type(' ')
    await rule(/^Next Record$/)
    await mergeField('Name')

    const preview = page.locator('.doc-page .doc-merge-preview')
    await expect(preview).toHaveCount(0)
    // the preview puts the Fill-in question of the record it shows; Enter takes the default
    const takeDefault = async () => {
      const question = page.locator('.modal', { hasText: 'Which gift?' }).last()
      await expect(question.locator('input')).toHaveValue('a card')
      await question.locator('input').press('Enter')
      await expect(question).toHaveCount(0)
    }
    await big(page, 'Preview Results').click()
    await takeDefault()
    await expect(preview).toHaveText(['Ada', '1', '1', 'a card', '5%', 'Alan'])
    await greeting.scrollIntoViewIfNeeded()
    await page.screenshot({ path: test.info().outputPath('preview-rules.png') })
    const next = page.locator('[data-ribbon-body] button[aria-label="Next"]')
    await next.click()
    await takeDefault()
    await expect(record).toHaveText(/Record 2 of 4/)
    await expect(preview).toHaveText(['Alan', '2', '2', 'a card', '5%', 'Grace'])
    // Ada's letter took Alan along, so Grace opens the second one written
    await next.click()
    await takeDefault()
    await expect(preview).toHaveText(['Grace', '3', '2', 'a card', '5%', 'Edsger'])
    await big(page, 'Preview Results').click()
    await expect(preview).toHaveCount(0)

    // the merge writes what the preview showed
    await big(page, 'Finish & Merge').click()
    const panel = page.locator('.modal-mailmerge')
    await panel.locator('input[name="mailmerge-output"]').nth(1).check()
    await panel.locator('button.btn-primary').click()
    for (let asked = 0; asked < 2; asked++) {
      const question = page.locator('.modal', { hasText: 'Which gift?' }).last()
      await expect(question.locator('input')).toHaveValue('a card')
      await question.locator('input').press('Enter')
    }
    await expect(panel.locator('.mailmerge-done')).toBeVisible({ timeout: 30_000 })
    await panel.getByRole('button', { name: 'Close' }).click()
    expect(await paragraphs(outPath)).toEqual([
      'Dear Ada no. 1/1: a card',
      'Terms at 5%',
      'Regards Alan',
      'Dear Grace no. 3/2: a card',
      'Terms at 5%',
      'Regards Edsger',
    ])

    // typing into the SET's paragraph rewrites it: the field keeps its bookmark around its result
    await terms.click()
    await page.keyboard.press('Home')
    await page.keyboard.type('Our ')
    await expect(page.locator('.doc-page p', { hasText: 'Our Terms' })).toHaveCount(1)
    await saveFromMenu(app)
    await expect.poll(() => documentXml(docPath), { timeout: 30_000 }).toContain('Our ')
    const saved = await documentXml(docPath)
    expect(saved).toMatch(
      / SET Rate (?:"|&quot;)5%(?:"|&quot;) <\/w:instrText><\/w:r><w:r><w:fldChar w:fldCharType="separate"\/><\/w:r><w:bookmarkStart w:id="(\d+)" w:name="Rate"\/><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:t[^>]*>5%<\/w:t><\/w:r><w:bookmarkEnd w:id="\1"\/><w:r><w:fldChar w:fldCharType="end"\/>/,
    )
    expect(saved.match(/w:name="Rate"/g)).toHaveLength(1)
    expect(saved).toContain(' MERGEREC ')
    if (process.env.KEEP_E2E_OUTPUT)
      writeFileSync(join(process.env.KEEP_E2E_OUTPUT, 'letter.docx'), readFileSync(docPath))
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a sheet of labels is saved as a labels main document and reopens as one', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-labels-main-')))
  const docPath = join(dir, 'labels.docx')
  const sectPr =
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
    '<w:cols w:space="720"/></w:sectPr>'
  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
  const zip = await JSZip.loadAsync(await docxWith(para('Dear') + sectPr))
  zip.file(
    'word/settings.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:settings ${NS}><w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><w:compat/></w:settings>`,
  )
  zip.file(
    'word/_rels/document.xml.rels',
    (await zip.file('word/_rels/document.xml.rels')!.async('string')).replace(
      '</Relationships>',
      '<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>',
    ),
  )
  zip.file(
    '[Content_Types].xml',
    (await zip.file('[Content_Types].xml')!.async('string')).replace(
      '</Types>',
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>',
    ),
  )
  writeFileSync(docPath, await zip.generateAsync({ type: 'nodebuffer' }))
  const listPath = join(dir, 'people.csv')
  writeFileSync(listPath, 'Name\nAda\nGrace\nAlan\n')
  const outPath = join(dir, 'merged.docx')
  const labelsRadio = (page: Page) =>
    page.locator('.modal-mailmerge input[name="mailmerge-output"]').nth(2)
  let launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const greeting = page.locator('.doc-page p', { hasText: 'Dear' }).first()
    await expect(greeting).toBeVisible({ timeout: 30_000 })
    await stubDialogs(app, listPath, outPath)
    await tab(page, 'Mailings')
    const panel = page.locator('.modal-mailmerge')
    await greeting.click()
    await page.keyboard.press('End')
    await page.keyboard.type(' ')
    await clickMenu(app, ['Tools', 'Labels…'])
    await panel.getByRole('button', { name: 'Select Recipient List…' }).click()
    await panel.locator('.mailmerge-field', { hasText: /^Name$/ }).click()
    await panel.locator('input[name="mailmerge-stock"]').nth(1).check()
    const grid = panel.locator('.mailmerge-pattern input[type="number"]')
    await grid.nth(0).fill('2')
    await grid.nth(1).fill('2')
    await panel.locator('button[name="createLabelSheet"]').click()
    await expect(panel).toBeHidden()
    await expect(page.locator('.doc-page td', { hasText: '«Name»' })).toHaveCount(4)

    await saveFromMenu(app)
    await expect
      .poll(() => part(docPath, 'word/settings.xml'), { timeout: 30_000 })
      .toContain('<w:mailMerge>')
    // at its place in CT_Settings: after the view settings, before the tab stops
    expect(await part(docPath, 'word/settings.xml')).toContain(
      '<w:zoom w:percent="100"/><w:mailMerge><w:mainDocumentType w:val="mailingLabels"/>' +
        '<w:dataType w:val="textFile"/></w:mailMerge><w:defaultTabStop w:val="720"/>',
    )
    if (process.env.KEEP_E2E_OUTPUT)
      writeFileSync(join(process.env.KEEP_E2E_OUTPUT, 'labels.docx'), readFileSync(docPath))
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
  }

  launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page td', { hasText: '«Name»' })).toHaveCount(4, {
      timeout: 30_000,
    })
    await stubDialogs(app, listPath, outPath)
    await tab(page, 'Mailings')
    await big(page, 'Select Recipients').click()
    await expect(
      page.locator('[data-ribbon-body] [data-testid="mailings-record-status"]'),
    ).toHaveText(/Record 1 of 3/)
    // the file says it is a labels main document: the merge opens on labels
    await big(page, 'Finish & Merge').click()
    const panel = page.locator('.modal-mailmerge')
    await expect(labelsRadio(page)).toBeChecked()
    await page.screenshot({ path: test.info().outputPath('labels-reopened.png') })
    await panel.locator('button.btn-primary').click()
    await expect(panel.locator('.mailmerge-done')).toBeVisible({ timeout: 30_000 })
    const merged = await documentXml(outPath)
    expect(merged.match(/Ada|Grace|Alan/g)).toEqual(['Ada', 'Grace', 'Alan'])
    expect(merged).not.toContain('NEXT')
    // the labels written are no main document
    expect(await part(outPath, 'word/settings.xml')).not.toContain('mailMerge')
    // saving the main document again leaves its one w:mailMerge alone
    await panel.getByRole('button', { name: 'Close' }).click()
    await page.locator('.doc-page td p').first().click()
    await page.keyboard.press('Home')
    await page.keyboard.type('To ')
    await saveFromMenu(app)
    await expect.poll(() => documentXml(docPath), { timeout: 30_000 }).toContain('To ')
    expect((await part(docPath, 'word/settings.xml')).match(/<w:mailMerge>/g)).toHaveLength(1)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Insert Index on the ribbon builds a two-column index in its own section; a page range lists its range; a table of authorities numbers its pages by chapter', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-index-leftovers-')))
  const docPath = join(dir, 'index.docx')
  const chapter = (n: number) =>
    `<w:p>${run('Chapter ')}${field('SEQ Chapter \\* ARABIC', run(String(n)))}</w:p>`
  const cite = field('TA \\l "Baldwin v. Alberti" \\c 1')
  const sectPr =
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
    '<w:cols w:space="720"/></w:sectPr>'
  writeFileSync(
    docPath,
    await docxWith(
      chapter(1) +
        `<w:p><w:bookmarkStart w:id="1" w:name="Orchard"/>${run('Orchard notes')}${field('XE "Apple"')}</w:p>` +
        `<w:p>${run('See Baldwin')}${cite}</w:p>` +
        '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' +
        chapter(2) +
        `<w:p>${run('More orchard notes')}${field('XE "Orchards" \\r Orchard')}<w:bookmarkEnd w:id="1"/></w:p>` +
        `<w:p>${run('Baldwin again')}${cite}</w:p>` +
        `<w:p>${field('TOA \\h \\c "1" \\s Chapter', run('stale'))}</w:p>` +
        para('Index goes here') +
        para('After the index') +
        sectPr,
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
    const spot = page.locator('.doc-page p', { hasText: 'Index goes here' })
    await expect(spot).toBeVisible({ timeout: 30_000 })
    const pageText = async () =>
      (await page.locator('main.editor-scroll .doc-page').allInnerTexts())
        .join('\n')
        .replace(/[ \t]+/g, ' ')
    await tab(page, 'References')

    // Update Table (the third of that name) rebuilds the table of authorities by its switches
    await big(page, /^Update Table$/)
      .nth(2)
      .click()
    await expect.poll(pageText).toMatch(/Baldwin v\. Alberti\s*\.*\s*1-1, 2-2/)
    expect(await pageText()).not.toContain('stale')

    await spot.click()
    await page.keyboard.press('End')
    await big(page, 'Insert Index').click()
    // the range runs from the page the bookmark opens on to the page it closes on
    await expect.poll(pageText, { timeout: 30_000 }).toMatch(/Orchards, 1\u20132/)
    expect(await pageText()).toMatch(/Apple, 1/)
    // the index is saved in a continuous section of two columns, as from the dialog
    const sections = async () =>
      [...(await documentXml(docPath)).matchAll(/<w:sectPr\b.*?<\/w:sectPr>/g)].map((m) => [
        /<w:cols\b[^>]*w:num="(\d+)"/.exec(m[0])?.[1] ?? '1',
        /<w:type w:val="([^"]+)"/.exec(m[0])?.[1] ?? '',
      ])
    await expect.poll(sections, { timeout: 30_000 }).toEqual([
      ['1', ''],
      ['2', 'continuous'],
      ['1', 'continuous'],
    ])
    const xml = await documentXml(docPath)
    expect(xml).toMatch(/<w:instrText[^>]*> INDEX \\c "2" </)
    expect(xml).toContain('1\u20132')
    expect(xml).toMatch(/> TOA \\h \\c "1" \\s Chapter </)
    // the two columns are drawn once the sections of the file are read back
    await expect
      .poll(
        () =>
          page.evaluate(() => {
            const lefts = Array.from(document.querySelectorAll('main.editor-scroll .doc-page *'))
              .filter(
                (el) => /^(Apple|Orchards), /.test(el.textContent ?? '') && !el.children.length,
              )
              .map((el) => Math.round(el.getBoundingClientRect().left))
            return new Set(lefts).size
          }),
        { timeout: 30_000 },
      )
      .toBe(2)
    await page.screenshot({ path: test.info().outputPath('index-two-columns.png') })
    if (process.env.KEEP_E2E_OUTPUT)
      writeFileSync(join(process.env.KEEP_E2E_OUTPUT, 'index.docx'), readFileSync(docPath))
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
