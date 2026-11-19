/**
 * More dialog controls that were shown greyed out and now work: Combine
 * Characters' Font and Size, Phonetic Guide's base text and alignment, the
 * Symbol dialog's Font, the Columns dialog's Apply to, and the Document
 * dialog's Default….
 */
import { test, expect, type ElectronApplication, type Locator, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-dialog-live-controls-2'
const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/></w:sectPr>'

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

/** selects document positions: the first paragraph's text starts at 1 */
const select = (page: Page, from: number, to: number) =>
  page.evaluate(
    ([from, to]) => {
      const { editor } = (
        window as never as {
          __aidocs: {
            editor: { chain(): { focus(): { setTextSelection(r: object): { run(): boolean } } } }
          }
        }
      ).__aidocs
      editor.chain().focus().setTextSelection({ from, to }).run()
    },
    [from, to],
  )

async function caretIn(paragraph: Locator) {
  await paragraph.click()
  await paragraph.page().keyboard.press('End')
}

const documentXml = async (docPath: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file('word/document.xml')?.async('string') ?? ''

test('Combine Characters takes a font and size, Phonetic Guide a base text and alignment, the Symbol dialog a symbol font', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-live-controls-2-')))
  const docPath = join(dir, 'layout.docx')
  writeFileSync(
    docPath,
    await docxWith(para('A grand day') + para('Before noon') + para('apple') + SECT_PR),
  )
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'A grand day' })).toBeVisible({
      timeout: 30_000,
    })

    // Combine Characters: "grand" in Arial at 18 pt
    await select(page, 3, 8)
    await clickMenu(app, ['Format', 'Asian Layout', 'Combine Characters…'])
    const combine = page.locator('.eq-layout-dialog')
    await expect(combine.locator('input[type=text]')).toHaveValue('grand')
    await combine.locator('select[name="font"]').selectOption('Arial')
    await combine.locator('select[name="size"]').selectOption('18')
    await page.screenshot({ path: test.info().outputPath('combine-font-size.png') })
    await combine.locator('button.primary').click()
    const combined = page.locator('.doc-page .doc-combine-chars')
    await expect(combined).toHaveCount(1)
    // two lines of half-size text: 9 pt
    await expect.poll(() => combined.evaluate((el) => getComputedStyle(el).fontSize)).toBe('12px')
    // opened on the combination again, the dialog shows what it was given
    await clickMenu(app, ['Format', 'Asian Layout', 'Combine Characters…'])
    await expect(combine.locator('select[name="font"]')).toHaveValue('Arial')
    await expect(combine.locator('select[name="size"]')).toHaveValue('18')
    await page.keyboard.press('Escape')
    await expect(combine).toBeHidden()

    // Phonetic Guide: a short reading set flush left, then spread, then centred, over other base text
    const before = page.locator('.doc-page p', { hasText: 'Before noon' }).first()
    await before
      .locator('text=Before')
      .first()
      .dblclick({ position: { x: 8, y: 8 } })
    await clickMenu(app, ['Format', 'Asian Layout', 'Phonetic Guide…'])
    const phonetic = page.locator('.phonetic-guide-dialog')
    await expect(phonetic.getByLabel('Base text')).toHaveValue(/^Before\s?$/)
    // with no base text there is nothing to annotate: OK waits for one
    await phonetic.getByLabel('Base text').fill('')
    await expect(phonetic.locator('button.primary')).toBeDisabled()
    await phonetic.getByLabel('Base text').fill('Earlier')
    await phonetic.getByLabel('Ruby text', { exact: true }).fill('x')
    await phonetic.locator('select[name="align"]').selectOption('left')
    await page.screenshot({ path: test.info().outputPath('phonetic-base-align.png') })
    await phonetic.locator('button.primary').click()
    const guide = page.locator('.doc-page ruby.doc-ruby')
    await expect(guide).toHaveAttribute('data-ruby-align', 'left')
    const placed = () => guide.evaluate((el) => getComputedStyle(el).getPropertyValue('ruby-align'))
    await expect.poll(placed).toBe('start')
    await page.screenshot({ path: test.info().outputPath('phonetic-left.png') })
    await clickMenu(app, ['Format', 'Asian Layout', 'Phonetic Guide…'])
    const align = phonetic.locator('select[name="align"]')
    await expect(align).toHaveValue('left')
    await expect(align.locator('option[value="right"]')).toBeEnabled()
    await align.selectOption('distributeLetter')
    await phonetic.locator('button.primary').click()
    await expect(guide).toHaveAttribute('data-ruby-align', 'distributeLetter')
    await expect.poll(placed).toBe('space-between')
    await clickMenu(app, ['Format', 'Asian Layout', 'Phonetic Guide…'])
    await align.selectOption('center')
    await phonetic.locator('button.primary').click()
    await expect.poll(placed).toBe('center')
    await page.screenshot({ path: test.info().outputPath('phonetic-result.png') })

    // Symbol: Wingdings lists its own characters; the check mark goes in as a symbol-font character
    const apple = page.locator('.doc-page p', { hasText: 'apple' }).first()
    await caretIn(apple)
    await clickMenu(app, ['Insert', 'Advanced Symbol…'])
    const symbol = page.locator('.symbol-dialog')
    const cells = symbol.locator('.symbol-grid .symbol-cell')
    await expect(cells.first()).toBeVisible()
    const plainCount = await cells.count()
    await symbol.locator('select[name="font"]').selectOption('Wingdings')
    await expect(cells).not.toHaveCount(plainCount)
    await expect(symbol.locator('[data-sym-char="F0FC"]')).toBeVisible()
    await symbol.locator('[data-sym-char="F0FC"]').click()
    await page.screenshot({ path: test.info().outputPath('symbol-font.png') })
    await symbol.locator('button.primary').click()
    await symbol.locator('[data-sym-char="F04A"]').dblclick()
    // back on normal text the grid is the Unicode one again
    await symbol.locator('select[name="font"]').selectOption('')
    await expect(cells).toHaveCount(plainCount)
    await page.keyboard.press('Escape')
    await expect(symbol).toBeHidden()
    await expect(apple.locator('[data-sym-char]')).toHaveCount(2)

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const xml = await documentXml(docPath)
          const field =
            /<w:fldChar w:fldCharType="begin"\/>[\s\S]*?<w:fldChar w:fldCharType="end"\/>/.exec(
              xml,
            )?.[0] ?? ''
          return [
            / eq \\o\(\\s\\up \d+\(/.test(field),
            /<w:rFonts\b[^>]*w:ascii="Arial"/.test(field),
            // the two lines are half of 18 pt
            /<w:sz w:val="18"\/>/.test(field),
            /<w:rubyAlign w:val="center"\/>/.test(xml),
            /<w:rubyBase>.*<w:t[^>]*>Earlier<\/w:t>.*<\/w:rubyBase>/.test(xml),
            /<w:sym w:font="Wingdings" w:char="F0FC"\/>/.test(xml),
            /<w:sym w:font="Wingdings" w:char="F04A"\/>/.test(xml),
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([true, true, true, true, true, true, true])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('the Columns dialog applies from the caret on and to the whole document', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-live-controls-2-')))
  const docPath = join(dir, 'columns.docx')
  writeFileSync(docPath, await docxWith(['apple', 'banana', 'cherry'].map(para).join('') + SECT_PR))
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const paragraph = (text: string) => page.locator('.doc-page p', { hasText: text }).first()
    await expect(paragraph('apple')).toBeVisible({ timeout: 30_000 })
    const dialog = page.locator('.columns-dialog')
    const applyTo = dialog.locator('select[name="applyTo"]')
    /** the w:cols of every section, in document order */
    const cols = async () =>
      ((await documentXml(docPath)).match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g) ?? []).map(
        (sectPr) =>
          `${/<w:cols\b[^>]*w:num="(\d+)"/.exec(sectPr)?.[1] ?? '1'}` +
          (/<w:type w:val="continuous"\/>/.test(sectPr) ? ' continuous' : ''),
      )

    // one section: the whole document is what the dialog acts on unless told otherwise
    await caretIn(paragraph('apple'))
    await clickMenu(app, ['Format', 'Columns…'])
    await expect(applyTo).toHaveValue('whole')
    await expect(applyTo.locator('option')).toHaveText(['This point forward', 'Whole document'])
    await dialog.locator('input[name="columns"]').nth(1).check()
    await applyTo.selectOption('forward')
    await page.screenshot({ path: test.info().outputPath('columns-apply-to.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    // the paragraph with the caret stays in one column; what follows starts two on the same page
    await expect.poll(cols, { timeout: 30_000 }).toEqual(['1', '2 continuous'])

    // two sections now: This section is offered and is the default
    const counts = async () => (await cols()).map((c) => c.split(' ')[0])
    await caretIn(paragraph('apple'))
    await clickMenu(app, ['Format', 'Columns…'])
    await expect(applyTo).toHaveValue('section')
    await expect(applyTo.locator('option')).toHaveText([
      'This section',
      'This point forward',
      'Whole document',
    ])
    // from the caret on reaches every section after the new break, not only the next one
    await dialog.locator('input[name="columns"]').nth(2).check()
    await applyTo.selectOption('forward')
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect.poll(counts, { timeout: 30_000 }).toEqual(['1', '3', '3'])

    await caretIn(paragraph('cherry'))
    await clickMenu(app, ['Format', 'Columns…'])
    await dialog.locator('input[name="columns"]').nth(1).check()
    await applyTo.selectOption('whole')
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await saveFromMenu(app)
    await expect.poll(counts, { timeout: 30_000 }).toEqual(['2', '2', '2'])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('the Document dialog keeps its page setup as the default for new blank documents', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-live-controls-2-')))
  const docPath = join(dir, 'letter.docx')
  writeFileSync(docPath, await docxWith(para('apple') + SECT_PR))
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const apple = page.locator('.doc-page p', { hasText: 'apple' }).first()
    await expect(apple).toBeVisible({ timeout: 30_000 })
    await caretIn(apple)
    const dialog = page.locator('.doc-dialog:not(.borders-shading-dialog)')
    const left = dialog.getByLabel('Left', { exact: true })

    await clickMenu(app, ['Format', 'Document…'])
    await left.fill('144pt')
    await left.blur()
    const kept = await left.inputValue()
    await page.screenshot({ path: test.info().outputPath('document-default.png') })
    await dialog.locator('button[name="default"]').click()
    await expect(dialog).toBeHidden()
    await expect(page.locator('.status-msg')).toContainText('this page setup')
    // Default… applies to this document as OK does
    await saveFromMenu(app)
    await expect
      .poll(async () => /<w:pgMar\b[^>]* w:left="2880"/.test(await documentXml(docPath)), {
        timeout: 30_000,
      })
      .toBe(true)

    // a new blank document starts with the kept margin
    await clickMenu(app, ['File', 'New Document'])
    await expect(apple).toHaveCount(0, { timeout: 30_000 })
    await page.locator('.doc-page').first().click()
    await clickMenu(app, ['Format', 'Document…'])
    await expect(left).toHaveValue(kept)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('the Document dialog applies from the caret on and sets the document grid', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-live-controls-2-')))
  const docPath = join(dir, 'letter.docx')
  writeFileSync(docPath, await docxWith(['apple', 'banana', 'cherry'].map(para).join('') + SECT_PR))
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const paragraph = (text: string) => page.locator('.doc-page p', { hasText: text }).first()
    await expect(paragraph('apple')).toBeVisible({ timeout: 30_000 })
    const dialog = page.locator('.doc-dialog:not(.borders-shading-dialog)')
    const applyTo = dialog.locator('select[name="applyTo"]')
    const sectPrs = async () =>
      (await documentXml(docPath)).match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g) ?? []

    // a wider left margin from the caret on: the first paragraph keeps the page it had
    await caretIn(paragraph('apple'))
    await clickMenu(app, ['Format', 'Document…'])
    await expect(applyTo).toHaveValue('whole')
    const left = dialog.getByLabel('Left', { exact: true })
    await left.fill('144pt')
    await left.blur()
    await applyTo.selectOption('forward')
    await page.screenshot({ path: test.info().outputPath('document-apply-to.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect
      .poll(
        async () =>
          (await sectPrs()).map(
            (sectPr) =>
              `${/<w:pgMar\b[^>]* w:left="(\d+)"/.exec(sectPr)?.[1]}` +
              (/<w:type\b/.test(sectPr) ? ' typed' : ''),
          ),
        { timeout: 30_000 },
      )
      .toEqual(['1440', '2880'])
    // the section after the break starts on a page of its own
    await expect(page.locator('.doc-page .page-gap')).toHaveCount(1)

    // the grid: lines only on the section with the caret, lines snapped to its pitch
    await caretIn(paragraph('cherry'))
    await clickMenu(app, ['Format', 'Document…'])
    await dialog.locator('[role="tab"]').nth(2).click()
    const grid = (type: string) => dialog.locator(`input[name="gridType"][value="${type}"]`)
    await expect(grid('default')).toBeChecked()
    await expect(grid('snapToChars')).toBeEnabled()
    const lineHeight = () =>
      paragraph('cherry').evaluate((el) => Math.round(el.getBoundingClientRect().height))
    const before = await lineHeight()
    await grid('lines').check()
    await page.screenshot({ path: test.info().outputPath('document-grid-type.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    // 360 twips of line pitch: 24 px, taller than the line was
    await expect.poll(lineHeight).toBeGreaterThan(before)
    await saveFromMenu(app)
    await expect
      .poll(
        async () => (await sectPrs()).map((sectPr) => /<w:docGrid\b[^>]*\/>/.exec(sectPr)?.[0]),
        {
          timeout: 30_000,
        },
      )
      .toEqual([undefined, '<w:docGrid w:type="lines" w:linePitch="360"/>'])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Page Numbers puts the number at the outside of facing pages', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-live-controls-2-')))
  const docPath = join(dir, 'pages.docx')
  const pageBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
  writeFileSync(docPath, await docxWith(para('apple') + pageBreak + para('banana') + SECT_PR))
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const apple = page.locator('.doc-page p', { hasText: 'apple' }).first()
    await expect(apple).toBeVisible({ timeout: 30_000 })
    await caretIn(apple)

    await clickMenu(app, ['Insert', 'Page Numbers…'])
    const dialog = page.locator('.page-numbers-dialog')
    await dialog.locator('select[name="align"]').selectOption('outside')
    await page.screenshot({ path: test.info().outputPath('page-numbers-outside.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()

    await saveFromMenu(app)
    /** the alignment of the footer each kind of page uses, read from the saved parts */
    const footers = async () => {
      const zip = await JSZip.loadAsync(readFileSync(docPath))
      const read = (path: string) => zip.file(path)?.async('string') ?? Promise.resolve('')
      const xml = await read('word/document.xml')
      const rels = await read('word/_rels/document.xml.rels')
      const out: Record<string, string> = {}
      for (const ref of xml.match(/<w:footerReference\b[^>]*\/>/g) ?? []) {
        const type = /w:type="([^"]+)"/.exec(ref)?.[1] ?? ''
        const id = /r:id="([^"]+)"/.exec(ref)?.[1] ?? ''
        const target = new RegExp(
          `Id="${id}"[^>]*Target="([^"]+)"|Target="([^"]+)"[^>]*Id="${id}"`,
        ).exec(rels)
        const part = await read(`word/${target?.[1] ?? target?.[2]}`)
        out[type] =
          `${/<w:jc w:val="([^"]+)"/.exec(part)?.[1] ?? 'left'}${/PAGE/.test(part) ? ' PAGE' : ''}`
      }
      return {
        ...out,
        evenAndOdd: /<w:evenAndOddHeaders\b/.test(await read('word/settings.xml')),
      }
    }
    await expect
      .poll(footers, { timeout: 30_000 })
      .toEqual({ default: 'right PAGE', even: 'left PAGE', evenAndOdd: true })
    await page.screenshot({
      path: test.info().outputPath('page-numbers-result.png'),
      fullPage: true,
    })
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
