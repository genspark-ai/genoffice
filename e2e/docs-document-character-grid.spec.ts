/**
 * Format > Document > Document Grid, the character side: Per line, Pitch and
 * Use default pitch on a line and character grid, Text snaps to character
 * grid, and Set Font… for the font the cells are counted from.
 */
import { test, expect, type ElectronApplication, type Locator, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-document-character-grid'
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
// 6.5" of text: 468pt
const PAGE =
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/>'
const KANJI = '\u6c38'.repeat(60)
const LATIN = 'The quick brown fox jumps over the lazy dog'

const para = (text: string, size?: number) =>
  `<w:p><w:r>${size ? `<w:rPr><w:sz w:val="${size}"/></w:rPr>` : ''}<w:t>${text}</w:t></w:r></w:p>`

/** a document whose default paragraph style is 10.5pt */
async function docx(body: string): Promise<Buffer> {
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
    `<?xml version="1.0" encoding="UTF-8"?><w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Songti SC"/><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>`,
  )
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body>${body}</w:body></w:document>`,
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

const part = async (docPath: string, name: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file(name)?.async('string') ?? ''

const grid = async (docPath: string) =>
  /<w:docGrid\b[^>]*\/>/.exec(await part(docPath, 'word/document.xml'))?.[0]

async function openDocs(docPath: string) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await expect(page.locator('.doc-page').first()).toContainText(LATIN, { timeout: 30_000 })
  return { launched, page }
}

/**
 * The nth paragraph's first line on an unzoomed page, px: how far each of its
 * first characters advances, where the first one starts from the paragraph's
 * edge, how many characters the line holds and how tall the paragraph is.
 */
const lineOf = (page: Page, index: number) =>
  page.evaluate((index) => {
    const paper = document.querySelector('.doc-page')!.getBoundingClientRect()
    const k = 816 / paper.width
    const p = document.querySelectorAll('.doc-page p')[index]!
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT)
    const node = walker.nextNode() as Text
    const range = document.createRange()
    const box = (i: number) => {
      range.setStart(node, i)
      range.setEnd(node, i + 1)
      return range.getBoundingClientRect()
    }
    const first = box(0)
    let count = 1
    while (count < node.length && Math.abs(box(count).top - first.top) < 2) count++
    const round = (n: number) => Math.round(n * k * 100) / 100
    return {
      advance: round(box(1).left - first.left),
      inset: round(first.left - p.getBoundingClientRect().left),
      count,
      height: round(p.getBoundingClientRect().height),
    }
  }, index)

async function openGrid(app: ElectronApplication, page: Page): Promise<Locator> {
  await page.locator('.doc-page p').first().click()
  await clickMenu(app, ['Format', 'Document…'])
  const dialog = page.locator('.doc-dialog')
  await dialog.getByRole('tab', { name: 'Document Grid' }).click()
  return dialog
}

async function accept(dialog: Locator) {
  await dialog.locator('button.primary').first().click()
  await expect(dialog).toBeHidden()
}

test('a line and character grid is set in characters per line or pitch', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-char-grid-')))
  const docPath = join(dir, 'grid.docx')
  writeFileSync(
    docPath,
    await docx(
      para(KANJI) +
        para(LATIN) +
        `<w:sectPr>${PAGE}<w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>`,
    ),
  )
  const first = await openDocs(docPath)
  try {
    const { launched, page } = first
    // 10.5pt characters, 44 to a line of 468pt
    await expect.poll(async () => (await lineOf(page, 0)).advance).toBeCloseTo(14, 1)
    expect((await lineOf(page, 0)).count).toBe(44)
    const latin = (await lineOf(page, 1)).advance
    const dialog = await openGrid(launched.app, page)
    const perLine = dialog.locator('input[name="charsPerLine"]')
    const pitch = dialog.locator('input[name="charPitch"]')
    const byDefault = dialog.locator('input[name="defaultCharPitch"]')
    for (const control of [perLine, pitch, byDefault]) await expect(control).toBeDisabled()
    await dialog.locator('input[name="gridType"][value="linesAndChars"]').check()
    await expect(byDefault).toBeChecked()
    await expect(perLine).toBeDisabled()
    await expect(perLine).toHaveValue('44')
    await expect(pitch).toHaveValue('10.5')
    await byDefault.uncheck()
    await perLine.fill('36')
    await expect(pitch).toHaveValue('13')
    await page.screenshot({ path: test.info().outputPath('character-grid-dialog.png') })
    await accept(dialog)
    // every character, Latin too, advances by the 2.5pt the pitch adds
    await expect.poll(async () => (await lineOf(page, 0)).advance).toBeCloseTo(17.33, 1)
    expect((await lineOf(page, 0)).count).toBe(36)
    expect((await lineOf(page, 1)).advance).toBeCloseTo(latin + 3.33, 1)
    await page.screenshot({ path: test.info().outputPath('character-grid-page.png') })
    await saveFromMenu(launched.app)
    await expect
      .poll(() => grid(docPath), { timeout: 30_000 })
      .toBe('<w:docGrid w:type="linesAndChars" w:linePitch="360" w:charSpace="10240"/>')

    // the pitch typed, in points
    await openGrid(launched.app, page)
    await expect(perLine).toHaveValue('36')
    await expect(byDefault).not.toBeChecked()
    await pitch.fill('12')
    await expect(perLine).toHaveValue('39')
    await accept(dialog)
    await expect.poll(async () => (await lineOf(page, 0)).advance).toBeCloseTo(16, 1)
    await saveFromMenu(launched.app)
    await expect
      .poll(() => grid(docPath), { timeout: 30_000 })
      .toBe('<w:docGrid w:type="linesAndChars" w:linePitch="360" w:charSpace="6144"/>')
  } finally {
    await closeAndSaveVideo(first.launched, VIDEO_DIR)
  }

  // the saved pitch reads back; the default pitch takes the attribute out again
  const second = await openDocs(docPath)
  try {
    const { launched, page } = second
    await expect.poll(async () => (await lineOf(page, 0)).advance).toBeCloseTo(16, 1)
    expect((await lineOf(page, 0)).count).toBe(39)
    const dialog = await openGrid(launched.app, page)
    await expect(dialog.locator('input[name="charsPerLine"]')).toHaveValue('39')
    await expect(dialog.locator('input[name="charPitch"]')).toHaveValue('12')
    await dialog.locator('input[name="defaultCharPitch"]').check()
    await expect(dialog.locator('input[name="charsPerLine"]')).toHaveValue('44')
    await accept(dialog)
    await expect.poll(async () => (await lineOf(page, 0)).advance).toBeCloseTo(14, 1)
    await saveFromMenu(launched.app)
    await expect
      .poll(() => grid(docPath), { timeout: 30_000 })
      .toBe('<w:docGrid w:type="linesAndChars" w:linePitch="360"/>')
  } finally {
    await closeAndSaveVideo(second.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('text snaps to the character grid, and Set Font changes the font its cells are counted from', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-char-grid-')))
  const docPath = join(dir, 'snap.docx')
  // cells of 10.5pt + 6144/4096pt = 12pt on lines of 18pt
  writeFileSync(
    docPath,
    await docx(
      para(KANJI) +
        para(LATIN) +
        para(KANJI, 32) +
        `<w:sectPr>${PAGE}<w:docGrid w:type="snapToChars" w:linePitch="360" w:charSpace="6144"/></w:sectPr>`,
    ),
  )
  const first = await openDocs(docPath)
  try {
    const { launched, page } = first
    // a 10.5pt character sits centred in one cell, a 16pt one takes two
    await expect.poll(async () => (await lineOf(page, 0)).advance).toBeCloseTo(16, 1)
    const small = await lineOf(page, 0)
    expect(small.inset).toBeCloseTo(1, 1)
    expect(small.count).toBe(39)
    const big = await lineOf(page, 2)
    expect(big.advance).toBeCloseTo(32, 1)
    expect(big.inset).toBeCloseTo(5.33, 1)
    expect(big.count).toBe(19)
    // Latin text keeps its natural advances
    const { advance: latin, height: latinLine } = await lineOf(page, 1)
    await page.screenshot({ path: test.info().outputPath('snap-grid-page.png') })

    const dialog = await openGrid(launched.app, page)
    await expect(dialog.locator('input[name="gridType"][value="snapToChars"]')).toBeChecked()
    await expect(dialog.locator('input[name="charsPerLine"]')).toHaveValue('39')
    await expect(dialog.locator('input[name="charPitch"]')).toHaveValue('12')
    // on a line and character grid the same pitch is added to every character instead
    await dialog.locator('input[name="gridType"][value="linesAndChars"]').check()
    await accept(dialog)
    await expect.poll(async () => (await lineOf(page, 2)).advance).toBeCloseTo(23.33, 1)
    expect((await lineOf(page, 1)).advance).toBeCloseTo(latin + 2, 1)
    expect((await lineOf(page, 0)).inset).toBeCloseTo(0, 1)
    // the snapped section stood on the same line grid: what a line holds depends on the font
    expect((await lineOf(page, 1)).height).toBeCloseTo(latinLine, 1)

    // back on the snapped grid, with a 12pt font: cells of 13.5pt
    await openGrid(launched.app, page)
    await dialog.locator('input[name="gridType"][value="snapToChars"]').check()
    await dialog.locator('button[name="gridSetFont"]').click()
    const font = page.locator('.grid-font-dialog')
    await font.locator('button.gs-dd-btn[aria-label="Size"]').click()
    await font.getByRole('option', { name: '12', exact: true }).click()
    await page.screenshot({ path: test.info().outputPath('snap-grid-set-font.png') })
    await font.locator('button.primary').click()
    await expect(font).toBeHidden()
    await expect(dialog.locator('input[name="charPitch"]')).toHaveValue('13.5')
    await expect(dialog.locator('input[name="charsPerLine"]')).toHaveValue('34')
    await accept(dialog)
    await expect.poll(async () => (await lineOf(page, 0)).advance).toBeCloseTo(18, 1)
    expect((await lineOf(page, 0)).count).toBe(34)
    expect((await lineOf(page, 1)).advance).toBeGreaterThan(latin)
    await saveFromMenu(launched.app)
    await expect
      .poll(() => grid(docPath), { timeout: 30_000 })
      .toBe('<w:docGrid w:type="snapToChars" w:linePitch="360" w:charSpace="6144"/>')
    await expect
      .poll(
        async () =>
          /<w:style\b[^>]*w:styleId="Normal"[\s\S]*?<\/w:style>/.exec(
            await part(docPath, 'word/styles.xml'),
          )?.[0],
      )
      .toContain('<w:sz w:val="24"/>')
  } finally {
    await closeAndSaveVideo(first.launched, VIDEO_DIR)
  }

  const second = await openDocs(docPath)
  try {
    const { launched, page } = second
    await expect.poll(async () => (await lineOf(page, 0)).advance).toBeCloseTo(18, 1)
    expect((await lineOf(page, 0)).count).toBe(34)
    const dialog = await openGrid(launched.app, page)
    await expect(dialog.locator('input[name="gridType"][value="snapToChars"]')).toBeChecked()
    await expect(dialog.locator('input[name="charPitch"]')).toHaveValue('13.5')
  } finally {
    await closeAndSaveVideo(second.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
