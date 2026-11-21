/**
 * An index inserted from the References tab keeps its entries and their page
 * numbers apart after the file is saved and opened again: the entry on the
 * left, the page right aligned behind the leader, on the page and in the
 * exported PDF.
 */
import { test, expect, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-index-entries-reopen'

const xe = (term: string) =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> XE "${term}" </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
const marked = (text: string, term: string) => para(text).replace('</w:p>', `${xe(term)}</w:p>`)

const documentXml = async (docPath: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file('word/document.xml')?.async('string') ?? ''

interface Line {
  title: string
  page: string
  leader: string
  /** px between the end of the entry text and the start of its page number */
  gap: number
}

const indexLines = (page: Page): Promise<Line[]> =>
  page.evaluate(() =>
    Array.from(document.querySelectorAll('main.editor-scroll .doc-page .doc-toc-line')).map(
      (el) => {
        const title = el.querySelector('.doc-toc-title')
        const number = el.querySelector('.doc-toc-page')
        return {
          title: title?.textContent ?? '',
          page: number?.textContent ?? '',
          leader: el.querySelector('.doc-toc-dots')?.textContent?.slice(0, 1) ?? '',
          gap:
            title && number
              ? number.getBoundingClientRect().left - title.getBoundingClientRect().right
              : 0,
        }
      },
    ),
  )

/** px by which the entry titled `sub` starts to the right of the one titled `main` */
const indentOf = (page: Page, sub: string, main: string) =>
  page.evaluate(
    ([sub, main]) => {
      const left = (title: string) =>
        Array.from(document.querySelectorAll('main.editor-scroll .doc-page .doc-toc-title'))
          .find((el) => el.textContent === title)!
          .getBoundingClientRect().left
      return left(sub) - left(main)
    },
    [sub, main],
  )

async function pdfText(path: string): Promise<string> {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs')
  const doc = await pdfjs.getDocument({ data: new Uint8Array(readFileSync(path)) }).promise
  let out = ''
  for (let i = 1; i <= doc.numPages; i++) {
    const content = await (await doc.getPage(i)).getTextContent()
    out += content.items.map((it) => ('str' in it ? it.str : '')).join('')
  }
  return out
}

test('an index from the References tab reopens with entry and page apart, on the page and in the PDF', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-index-reopen-')))
  const docPath = join(dir, 'index.docx')
  writeFileSync(
    docPath,
    await docxWith(
      marked('about zebras', 'Zebra') +
        marked('about apples', 'Apple') +
        marked('about mangos', 'Mango') +
        marked('about pears', 'Fruit:Pear') +
        para('Index goes here'),
    ),
  )
  const expected = [
    ['Apple', '1', '.'],
    ['Pear', '1', '.'],
    ['Mango', '1', '.'],
    ['Zebra', '1', '.'],
  ]
  const shown = async (page: Page) =>
    (await indexLines(page)).map((l) => [l.title, l.page, l.leader])

  const before = { subIndent: 0 }
  let launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    const last = page.locator('.doc-page p', { hasText: 'Index goes here' })
    await expect(last).toBeVisible({ timeout: 30_000 })
    await last.click()
    await page.keyboard.press('End')
    await page.locator('.ribbon-tab', { hasText: /^References$/ }).click()
    await page
      .locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: 'Insert Index' })
      .click()
    // the main entry of a subentry has no page of its own
    await expect
      .poll(() => shown(page))
      .toEqual([expected[0], ['Fruit', '', ''], ...expected.slice(1)])
    before.subIndent = await indentOf(page, 'Pear', 'Apple')
    expect(before.subIndent).toBeGreaterThan(10)
    await saveFromMenu(launched.app)
    await expect
      .poll(async () => /INDEX\b/.test(await documentXml(docPath)), { timeout: 30_000 })
      .toBe(true)
    expect(await documentXml(docPath)).toContain(' INDEX \\e "\t" \\c "2" ')
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
  }

  launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'Index goes here' })).toBeVisible({
      timeout: 30_000,
    })
    await expect.poll(() => shown(page)).toEqual(expected)
    // the page sits at the right tab, well clear of the entry
    for (const line of await indexLines(page)) expect(line.gap).toBeGreaterThan(100)
    // the subentry sits where it did before the save: one indent, not several stacked
    expect(Math.abs((await indentOf(page, 'Pear', 'Apple')) - before.subIndent)).toBeLessThan(1.5)
    await expect(page.locator('.doc-page', { hasText: 'Fruit' }).first()).toBeVisible()
    expect(await page.locator('.doc-page').first().innerText()).not.toMatch(
      /Apple1|Pear1|Mango1|Zebra1/,
    )
    await page.screenshot({ path: test.info().outputPath('index-reopened.png') })

    const pdfPath = join(dir, 'index.pdf')
    await launched.app.evaluate(({ dialog, shell }, out) => {
      dialog.showSaveDialog = async () => ({ canceled: false, filePath: out })
      shell.showItemInFolder = () => {}
    }, pdfPath)
    expect(
      await page.evaluate(() =>
        (window as unknown as { __exportPdf: () => Promise<boolean> }).__exportPdf(),
      ),
    ).toBe(true)
    const text = await pdfText(pdfPath)
    for (const entry of ['Apple', 'Pear', 'Mango', 'Zebra'])
      expect(text).toMatch(new RegExp(`${entry}\\s*\\.{5,}\\s*1`))
    expect(text).not.toMatch(/Apple1|Pear1|Mango1|Zebra1/)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('an index whose Index 2 style carries the indent, as Word writes it, shows the subentry at that indent', async () => {
  test.setTimeout(120_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-index-reopen-')))
  const docPath = join(dir, 'styled-index.docx')
  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
  const stop = '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="4310"/></w:tabs>'
  const entry = (style: string, text: string, pageNo: string, before = '', after = '') =>
    `<w:p><w:pPr><w:pStyle w:val="${style}"/>${stop}</w:pPr>${before}` +
    `<w:r><w:t>${text}</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>${pageNo}</w:t></w:r>${after}</w:p>`
  const begin =
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> INDEX \\e "\t" \\c "2" \\z "1033" </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
  const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
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
  zip.file(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:styles ${NS}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
      '<w:style w:type="paragraph" w:styleId="Index1"><w:name w:val="index 1"/><w:basedOn w:val="Normal"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Index2"><w:name w:val="index 2"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="480"/></w:pPr></w:style></w:styles>',
  )
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body>` +
      para('Body text') +
      entry('Index1', 'Apple', '1', begin) +
      entry('Index2', 'Pear', '2') +
      entry('Index1', 'Zebra', '3', '', end) +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>' +
      '</w:body></w:document>',
  )
  writeFileSync(docPath, await zip.generateAsync({ type: 'nodebuffer' }))
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(page.locator('.doc-page .doc-toc-line')).toHaveCount(3, { timeout: 30_000 })
    expect((await indexLines(page)).map((l) => [l.title, l.page, l.leader])).toEqual([
      ['Apple', '1', '.'],
      ['Pear', '2', '.'],
      ['Zebra', '3', '.'],
    ])
    // 480 twips = 32 px, once
    expect(Math.abs((await indentOf(page, 'Pear', 'Apple')) - 32)).toBeLessThan(1.5)
    await page.screenshot({ path: test.info().outputPath('index-style-indent.png') })
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
