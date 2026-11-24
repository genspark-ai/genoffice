/**
 * References and Mailings as Word has them: citations are CITATION fields the
 * Style: box restyles together with the bibliography, Mark Index Entry and
 * Mark Citation take their options and Mark All, and the Rules menu's fields
 * are worked out when the merge runs.
 */
import { test, expect, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-references-mailings-gaps'

const part = async (file: string, name: string) =>
  (await (await JSZip.loadAsync(readFileSync(file))).file(name)?.async('string')) ?? ''
const documentXml = (file: string) => part(file, 'word/document.xml')

const fld = (type: string) => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`
const instr = (text: string) =>
  `<w:r><w:instrText xml:space="preserve"> ${text} </w:instrText></w:r>`

const BIB_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/bibliography'
/** the sources part as Word writes it: two authors, a city, and fields the app does not model */
const SOURCES =
  `<b:Sources SelectedStyle="\\APASixthEditionOfficeOnline.xsl" StyleName="APA" Version="6" xmlns:b="${BIB_NS}" xmlns="${BIB_NS}">` +
  '<b:Source><b:Tag>Mat11</b:Tag><b:SourceType>Book</b:SourceType><b:Guid>{AF584761-D8AB-49AE-8D03-6D025F356F0F}</b:Guid>' +
  '<b:Author><b:Author><b:NameList><b:Person><b:Last>Mattmann</b:Last><b:First>Chris</b:First></b:Person>' +
  '<b:Person><b:Last>Zitting</b:Last><b:First>Jukka</b:First></b:Person></b:NameList></b:Author></b:Author>' +
  '<b:Title>Tika in Action</b:Title><b:Year>2011</b:Year><b:City>Shelter Island</b:City>' +
  '<b:Publisher>Manning</b:Publisher><b:RefOrder>1</b:RefOrder></b:Source></b:Sources>'

/** a citation and a bibliography in the content controls Word wraps them in */
const WORD_CITED_BODY =
  '<w:p><w:r><w:t xml:space="preserve">As shown </w:t></w:r>' +
  '<w:sdt><w:sdtPr><w:id w:val="-2100173194"/><w:citation/></w:sdtPr><w:sdtContent>' +
  fld('begin') +
  instr('CITATION Mat11 \\l 1033') +
  fld('separate') +
  '<w:r><w:rPr><w:noProof/></w:rPr><w:t>(Mattmann &amp; Zitting, 2011)</w:t></w:r>' +
  fld('end') +
  '</w:sdtContent></w:sdt><w:r><w:t>.</w:t></w:r></w:p>' +
  para('Second thought') +
  '<w:sdt><w:sdtPr><w:id w:val="111145805"/><w:bibliography/></w:sdtPr><w:sdtContent>' +
  '<w:p><w:pPr><w:ind w:left="720" w:hanging="720"/></w:pPr>' +
  fld('begin') +
  instr('BIBLIOGRAPHY') +
  fld('separate') +
  '<w:r><w:t xml:space="preserve">Mattmann, C., &amp; Zitting, J. (2011). </w:t></w:r>' +
  '<w:r><w:rPr><w:i/></w:rPr><w:t>Tika in Action.</w:t></w:r>' +
  '<w:r><w:t xml:space="preserve"> Shelter Island: Manning.</w:t></w:r></w:p>' +
  `<w:p>${fld('end')}</w:p>` +
  '</w:sdtContent></w:sdt>' +
  para('Closing line')

async function withSources(docx: Buffer): Promise<Buffer> {
  const zip = await JSZip.loadAsync(docx)
  zip.file('customXml/item1.xml', SOURCES)
  return zip.generateAsync({ type: 'nodebuffer' })
}

const tab = (page: Page, name: string) =>
  page.locator('.ribbon-tab', { hasText: new RegExp(`^${name}$`) }).click()
const big = (page: Page, name: string | RegExp) =>
  page.locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: name })
const menuItem = (page: Page, name: string | RegExp) =>
  page.locator('[data-rb-panel] button', { hasText: name })
/** the page numbers the page draws in bold */
const boldPages = (page: Page) =>
  page.evaluate(() => {
    const out: string[] = []
    const walker = document.createTreeWalker(
      document.querySelector('.doc-page')!,
      NodeFilter.SHOW_TEXT,
    )
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      const text = (node.textContent ?? '').trim()
      if (!/^\d+$/.test(text) || !node.parentElement) continue
      if (Number(getComputedStyle(node.parentElement).fontWeight) >= 600) out.push(text)
    }
    return out
  })
const pageText = (page: Page) =>
  page.locator('.doc-page').evaluateAll((pages) => pages.map((p) => p.textContent).join('\n'))

test('Style: restyles the citations and the bibliography of a document from Word, and the file names the style', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-citations-')))
  const docPath = join(dir, 'paper.docx')
  writeFileSync(docPath, await withSources(await docxWith(WORD_CITED_BODY)))
  let launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const cited = page.locator('.doc-page p', { hasText: 'As shown' }).first()
    await expect(cited).toBeVisible({ timeout: 30_000 })
    await expect(cited).toContainText('(Mattmann & Zitting, 2011)')

    // the document's own style is the one the box names
    await tab(page, 'References')
    const style = big(page, 'Style')
    await expect(style).toBeEnabled()
    await expect(style).toContainText('APA')
    await style.click()
    await expect(menuItem(page, /./)).toHaveText([
      'APA ✓',
      'Chicago',
      'Harvard - Anglia',
      'IEEE',
      'MLA',
    ])
    await page.screenshot({ path: test.info().outputPath('citation-style-menu.png') })
    await menuItem(page, 'MLA').click()
    await expect(cited).toContainText('As shown (Mattmann and Zitting).')
    await expect(style).toContainText('MLA')
    await expect
      .poll(() => pageText(page))
      .toContain(
        'Mattmann, Chris and Jukka Zitting. Tika in Action. Shelter Island: Manning. 2011.',
      )

    // a second source, cited in the next paragraph
    await big(page, /^Insert Citation/).click()
    await menuItem(page, 'Add New Source…').click()
    const source = page.locator('.modal', { hasText: 'Author' }).first()
    await source.locator('input').nth(0).fill('Kay, Alan')
    await source.locator('input').nth(1).fill('The Early History of Smalltalk')
    await source.locator('input').nth(2).fill('1993')
    await source.locator('button.btn-primary').click()
    await expect(source).toBeHidden()
    const second = page.locator('.doc-page p', { hasText: 'Second thought' }).first()
    await second.click()
    await page.keyboard.press('End')
    await page.keyboard.type(' ')
    await big(page, /^Insert Citation/).click()
    await menuItem(page, 'The Early Hi').click()
    await expect(second).toContainText('Second thought (Kay)')
    // the new source joins the bibliography the document already has
    await expect.poll(() => pageText(page)).toContain('Kay, Alan. The Early History of Smalltalk.')

    // a numbered style numbers the sources in the order the text cites them
    await style.click()
    await menuItem(page, 'IEEE').click()
    await expect(cited).toContainText('As shown [1].')
    await expect(second).toContainText('Second thought [2]')
    await expect
      .poll(() => pageText(page))
      .toContain('[1] C. Mattmann and J. Zitting, Tika in Action, Shelter Island: Manning, 2011.')
    await page.screenshot({ path: test.info().outputPath('citation-style-ieee.png') })

    await saveFromMenu(app)
    await expect
      .poll(() => part(docPath, 'customXml/item1.xml'), { timeout: 30_000 })
      .toContain('StyleName="IEEE"')
    const sources = await part(docPath, 'customXml/item1.xml')
    expect(sources).toContain('SelectedStyle="\\IEEE2006OfficeOnline.xsl"')
    expect(sources).toContain('<b:Guid>{AF584761-D8AB-49AE-8D03-6D025F356F0F}</b:Guid>')
    expect(sources).toContain('<b:Last>Kay</b:Last><b:First>Alan</b:First>')
    const saved = await documentXml(docPath)
    // each citation is a field whose result is the styled text
    expect(saved).toMatch(
      /CITATION Mat11 \\l 1033 <\/w:instrText><\/w:r><w:r><w:fldChar w:fldCharType="separate"\/><\/w:r><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:t[^>]*>\[1\]<\/w:t>/,
    )
    expect(saved).toMatch(/CITATION Kay\S* \\l 1033 <\/w:instrText>.*?<w:t[^>]*>\[2\]<\/w:t>/)
    expect(saved.match(/ BIBLIOGRAPHY /g)).toHaveLength(1)
    expect(saved).toMatch(/<w:i\/><w:iCs\/><w:noProof\/><\/w:rPr><w:t[^>]*>Tika in Action<\/w:t>/)
    expect(saved).not.toContain('Mattmann, C., &amp; Zitting, J. (2011)')

    // reopened, the citations are still fields and take another style
    await closeAndSaveVideo(launched, VIDEO_DIR)
    launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
    const reopened = await waitForPageWithUrl(launched.app, '://docs/')
    const again = reopened.locator('.doc-page p', { hasText: 'As shown' }).first()
    await expect(again).toContainText('As shown [1].', { timeout: 30_000 })
    await tab(reopened, 'References')
    await expect(big(reopened, 'Style')).toContainText('IEEE')
    await big(reopened, 'Style').click()
    await menuItem(reopened, 'Chicago').click()
    await expect(again).toContainText('As shown (Mattmann and Zitting 2011).')
    await expect(
      reopened.locator('.doc-page p', { hasText: 'Second thought' }).first(),
    ).toContainText('(Kay 1993)')
    await expect
      .poll(() => pageText(reopened))
      .toContain('Kay, Alan. 1993. The Early History of Smalltalk.')
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Mark Index Entry takes a subentry, a cross-reference, bold pages and Mark All; Mark Citation takes a short citation', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-mark-entry-')))
  const docPath = join(dir, 'brief.docx')
  writeFileSync(
    docPath,
    await docxWith(
      para('Apples grow in the orchard') +
        para('Pears ripen later') +
        para('More Apples were picked') +
        para('In Baldwin v. Alberti the court agreed') +
        para('The rule in Baldwin was followed') +
        para('Closing line'),
    ),
  )
  let launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const line = (text: string) => page.locator('.doc-page p', { hasText: text }).first()
    await expect(line('Apples grow')).toBeVisible({ timeout: 30_000 })
    /** selects a word of a paragraph by double-clicking it */
    const selectWord = async (inLine: string, word: string) => {
      await line(inLine).scrollIntoViewIfNeeded()
      const box = await line(inLine).evaluate((p, word) => {
        const node =
          p.firstChild?.nodeType === 3 ? p.firstChild : p.querySelector('span')!.firstChild!
        const at = (p.textContent ?? '').indexOf(word)
        const range = document.createRange()
        range.setStart(node, at)
        range.setEnd(node, at + word.length)
        const r = range.getBoundingClientRect()
        return { x: r.left + r.width / 2, y: r.top + r.height / 2 }
      }, word)
      await page.mouse.dblclick(box.x, box.y)
    }
    await tab(page, 'References')

    // Mark All: a subentry with a bold page, at the first Apples of every paragraph that has one
    await selectWord('Apples grow', 'Apples')
    await big(page, 'Mark Entry').click()
    const entry = page.locator('.mark-entry-dialog')
    await expect(entry.locator('input[name="main"]')).toHaveValue('Apples')
    await entry.locator('input[name="main"]').fill('Fruit')
    await entry.locator('input[name="sub"]').fill('Apples')
    await entry.locator('input[name="bold"]').check()
    await expect(entry.locator('input[name="cross"]')).toBeDisabled()
    await page.screenshot({ path: test.info().outputPath('mark-index-entry.png') })
    await entry.getByRole('button', { name: 'Mark All' }).click()
    await expect(entry).toBeHidden()
    const marks = page.locator('.doc-page .doc-xe-mark:not(.doc-ta-mark)')
    await expect(marks).toHaveCount(2)
    await expect(marks.first()).toHaveAttribute('data-xe-term', 'Fruit:Apples')
    await expect(marks.first()).toHaveAttribute('data-xe-switches', '\\b')

    // a cross-reference in place of a page
    await selectWord('Pears ripen', 'Pears')
    await big(page, 'Mark Entry').click()
    await entry.getByLabel('Cross-reference', { exact: true }).first().check()
    await expect(entry.locator('input[name="cross"]')).toHaveValue('See ')
    await entry.locator('input[name="cross"]').fill('See Fruit')
    await expect(entry.getByRole('button', { name: 'Mark All' })).toBeDisabled()
    await entry.getByRole('button', { name: 'Mark', exact: true }).click()
    await expect(marks).toHaveCount(3)

    await line('Closing line').click()
    await page.keyboard.press('End')
    await big(page, 'Insert Index').click()
    await expect.poll(() => pageText(page)).toContain('Pears. See Fruit')
    await expect.poll(() => boldPages(page)).toEqual(['1'])
    await page.screenshot({ path: test.info().outputPath('index.png') })

    // Mark Citation: the long citation where it is selected, the short one at its later mention
    await selectWord('In Baldwin', 'Baldwin')
    await big(page, /^Mark Citation$/).click()
    const citation = page.locator('.mark-citation-dialog')
    await expect(citation.locator('input[name="short"]')).toHaveValue('Baldwin')
    await citation.locator('input[name="long"]').fill('Baldwin v. Alberti, 58 Wn. 2d 243 (1961)')
    await page.screenshot({ path: test.info().outputPath('mark-citation.png') })
    await citation.getByRole('button', { name: 'Mark All' }).click()
    await expect(citation).toBeHidden()
    await expect(page.locator('.doc-page .doc-ta-mark')).toHaveCount(2)
    await line('Closing line').click()
    await page.keyboard.press('Home')
    await big(page, 'Insert Table of Authorities').click()
    // both mentions are one line of the table, under the heading of their category
    await expect
      .poll(async () => (await pageText(page)).split('Baldwin v. Alberti, 58 Wn. 2d 243').length)
      .toBe(2)
    await expect(page.locator('.doc-page').getByText('Cases', { exact: true })).toHaveCount(1)

    // the index's section was saved when it went in: wait for this save by what was added since
    await saveFromMenu(app)
    await expect.poll(() => documentXml(docPath), { timeout: 30_000 }).toContain(' TOA \\h \\p ')
    const saved = await documentXml(docPath)
    expect(saved).toContain(' XE "Pears" \\t "See Fruit" ')
    expect(saved.match(/ XE "Fruit:Apples" \\b /g)).toHaveLength(2)
    expect(saved).toMatch(
      /TA \\l (?:"|&quot;)Baldwin v\. Alberti, 58 Wn\. 2d 243 \(1961\)(?:"|&quot;) \\s (?:"|&quot;)Baldwin(?:"|&quot;) \\c 1 /,
    )
    expect(saved).toMatch(/ TA \\s (?:"|&quot;)Baldwin(?:"|&quot;) </)
    expect(saved).toContain(' TOA \\h \\p ')
    // the bold page is a run of its own in the index
    expect(saved).toMatch(/<w:b\/><w:noProof\/><\/w:rPr><w:t xml:space="preserve">1<\/w:t>/)

    // reopened: the entries keep their switches, and the index its bold page
    await closeAndSaveVideo(launched, VIDEO_DIR)
    launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
    const reopened = await waitForPageWithUrl(launched.app, '://docs/')
    const kept = reopened.locator('.doc-page .doc-xe-mark:not(.doc-ta-mark)')
    await expect(kept).toHaveCount(3, { timeout: 30_000 })
    await expect(kept.first()).toHaveAttribute('data-xe-switches', '\\b')
    await expect(kept.nth(1)).toHaveAttribute('data-xe-switches', '\\t "See Fruit"')
    await expect.poll(() => boldPages(reopened)).toEqual(['1'])
    await tab(reopened, 'References')
    await big(reopened, 'Update Index').click()
    await expect.poll(() => pageText(reopened)).toContain('Pears. See Fruit')
    await expect.poll(() => boldPages(reopened)).toEqual(['1'])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('the Rules menu writes Merge Record #, Merge Sequence #, Skip Record If, Fill-in and Set Bookmark, and Finish & Merge works them out', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-merge-rules-')))
  const docPath = join(dir, 'letter.docx')
  writeFileSync(docPath, await docxWith(para('Dear') + para('Regards')))
  const listPath = join(dir, 'people.csv')
  writeFileSync(listPath, 'Name,City\nAda,London\nAlan,Wilmslow\nGrace,London\nEdsger,Austin\n')
  const outPath = join(dir, 'merged.docx')
  let launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
  const stubDialogs = (app: typeof launched.app) =>
    app.evaluate(
      ({ dialog }, [list, out]) => {
        dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [list] })
        dialog.showSaveDialog = async () => ({ canceled: false, filePath: out })
      },
      [listPath, outPath],
    )
  /** Finish & Merge into one document, answering each question the merge asks */
  const finish = async (page: Page, answers: Array<string | null>) => {
    await big(page, 'Finish & Merge').click()
    const panel = page.locator('.modal-mailmerge')
    await panel.locator('input[name="mailmerge-output"]').nth(1).check()
    await panel.locator('button.btn-primary').click()
    for (const answer of answers) {
      const question = page.locator('.modal', { hasText: 'Which gift?' }).last()
      await expect(question.locator('input')).toHaveValue('a card')
      if (answer !== null) await question.locator('input').fill(answer)
      await question.locator('input').press('Enter')
    }
    await expect(panel.locator('.mailmerge-done')).toBeVisible({ timeout: 30_000 })
    await panel.getByRole('button', { name: 'Close' }).click()
    await expect(panel).toBeHidden()
  }
  /** the letters of the merged file, one string per paragraph that has text */
  const letters = async () =>
    [...(await documentXml(outPath)).matchAll(/<w:p[ >].*?<\/w:p>/g)]
      .map((p) => [...p[0].matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((t) => t[1]).join(''))
      .map((text) => text.replace(/\s+/g, ' ').trim())
      .filter(Boolean)
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const greeting = page.locator('.doc-page p', { hasText: 'Dear' }).first()
    await expect(greeting).toBeVisible({ timeout: 30_000 })
    await stubDialogs(app)
    await tab(page, 'Mailings')
    await big(page, 'Select Recipients').click()
    await expect(
      page.locator('[data-ribbon-body] [data-testid="mailings-record-status"]', {
        hasText: 'Record 1 of 4',
      }),
    ).toHaveCount(1)

    const rules = big(page, 'Rules')
    const rule = async (name: RegExp) => {
      await rules.click()
      await page.locator('.layout-menu button', { hasText: name }).click()
    }
    await rules.click()
    await expect(page.locator('.layout-menu button')).toHaveText([
      'Ask…',
      'Fill-in…',
      'If…Then…Else…',
      'Merge Record #',
      'Merge Sequence #',
      'Next Record',
      'Next Record If…',
      'Set Bookmark…',
      'Skip Record If…',
    ])
    await page.screenshot({ path: test.info().outputPath('rules-menu.png') })
    await rules.click()

    // Dear «Name», no. <record>/<sequence>: <gift>
    await greeting.click()
    await page.keyboard.press('End')
    await page.keyboard.type(' ')
    await big(page, 'Insert Merge Field').click()
    await page.locator('.rb-field-item', { hasText: /^Name$/ }).click()
    await page.keyboard.type(' no. ')
    await rule(/^Merge Record #$/)
    await page.keyboard.type('/')
    await rule(/^Merge Sequence #$/)
    await page.keyboard.type(': ')
    await rule(/^Fill-in…$/)
    const dialog = page.locator('.rule-dialog')
    await dialog.locator('input[name="prompt"]').fill('Which gift?')
    await dialog.locator('input[name="value"]').fill('a card')
    await page.screenshot({ path: test.info().outputPath('rule-fill-in.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await rule(/^Skip Record If…$/)
    await dialog.locator('select[name="field"]').selectOption('City')
    await dialog.locator('input[name="value"]').fill('London')
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    // Set Bookmark needs a name Word accepts
    await rule(/^Set Bookmark…$/)
    await dialog.locator('input[name="bookmark"]').fill('2 bad')
    await expect(dialog.locator('button.primary')).toBeDisabled()
    await dialog.locator('input[name="bookmark"]').fill('Season')
    await dialog.locator('input[name="value"]').fill('spring')
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()

    // the two Londoners are skipped and ask nothing; the others are asked in turn
    await finish(page, ['a book', null])
    expect(await letters()).toEqual([
      'Dear Alan no. 2/1: a book',
      'Regards',
      'Dear Edsger no. 4/2: a card',
      'Regards',
    ])
    const merged = await documentXml(outPath)
    expect(merged).not.toMatch(/SKIPIF|MERGEREC|MERGESEQ|FILLIN|MERGEFIELD| SET /)

    // the rules are fields in the saved main document
    await saveFromMenu(app)
    await expect.poll(() => documentXml(docPath), { timeout: 30_000 }).toContain(' MERGEREC ')
    const saved = await documentXml(docPath)
    expect(saved).toContain(' MERGESEQ ')
    expect(saved).toMatch(
      / FILLIN (?:"|&quot;)Which gift\?(?:"|&quot;) \\d (?:"|&quot;)a card(?:"|&quot;) </,
    )
    expect(saved).toMatch(/<w:instrText[^>]*> SKIPIF <\/w:instrText>/)
    expect(saved).toMatch(/ SET Season (?:"|&quot;)spring(?:"|&quot;) </)

    // reopened, the paragraph is still editable text with its fields, and the merge reads them
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(outPath, { force: true })
    launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
    const reopened = await waitForPageWithUrl(launched.app, '://docs/')
    const again = reopened.locator('.doc-page p', { hasText: 'Dear' }).first()
    await expect(again).toBeVisible({ timeout: 30_000 })
    await expect(again.locator('[data-instr-field="MERGEREC"]')).toHaveCount(1)
    await stubDialogs(launched.app)
    await tab(reopened, 'Mailings')
    await big(reopened, 'Select Recipients').click()
    await expect(
      reopened.locator('[data-ribbon-body] [data-testid="mailings-record-status"]', {
        hasText: 'Record 1 of 4',
      }),
    ).toHaveCount(1)
    await finish(reopened, [null, 'a pen'])
    expect(await letters()).toEqual([
      'Dear Alan no. 2/1: a card',
      'Regards',
      'Dear Edsger no. 4/2: a pen',
      'Regards',
    ])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
