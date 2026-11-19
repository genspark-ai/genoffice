/**
 * Controls that were shown greyed out in dialogs and dropdowns and now work:
 * the Document dialog's gutter, vertical alignment, columns and Borders…,
 * Line Numbers > Suppress for Current Paragraph, Page Numbers' Show
 * number on first page, the Document dialog's gutter position and section
 * start, and the Columns dialog's Left / Right presets, Line between and
 * Equal column width.
 */
import { test, expect, type ElectronApplication, type Locator } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

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

test('the Document dialog, Line Numbers and Page Numbers controls that were greyed out change the document', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-live-controls-')))
  const docPath = join(dir, 'letter.docx')
  writeFileSync(docPath, await docxWith(['apple', 'banana', 'cherry'].map(para).join('') + SECT_PR))
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-dialog-live-controls',
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const paragraph = (text: string) => page.locator('.doc-page p', { hasText: text }).first()
    await expect(paragraph('apple')).toBeVisible({ timeout: 30_000 })
    await paragraph('apple').click()

    // Format > Document: a gutter, the page centred top to bottom, two columns
    const dialog = page.locator('.doc-dialog:not(.borders-shading-dialog)')
    const tab = (index: number) => dialog.locator('[role="tab"]').nth(index)
    await clickMenu(app, ['Format', 'Document…'])
    await expect(dialog).toBeVisible()
    const gutter = dialog.getByLabel('Gutter', { exact: true })
    await gutter.fill('36pt')
    await gutter.blur()
    await page.screenshot({ path: test.info().outputPath('document-margins.png') })
    await tab(1).click()
    await dialog.locator('select[name="vAlign"]').selectOption('center')
    await page.screenshot({ path: test.info().outputPath('document-layout.png') })
    await tab(2).click()
    // a number of columns that cannot be used keeps OK from closing the dialog
    await dialog.locator('input[name="columns"]').fill('0')
    await expect(dialog.locator('button.primary')).toBeDisabled()
    await dialog.locator('input[name="columns"]').fill('2')
    await expect(dialog.locator('button.primary')).toBeEnabled()
    await page.screenshot({ path: test.info().outputPath('document-grid.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()

    // Borders… leaves for the Borders and Shading dialog
    await clickMenu(app, ['Format', 'Document…'])
    await tab(1).click()
    await dialog.locator('button', { hasText: /^Borders/ }).click()
    const borders = page.locator('.borders-shading-dialog')
    await expect(borders).toBeVisible()
    await expect(dialog).toBeHidden()
    await expect(borders.locator(':focus')).toHaveCount(1)
    await page.keyboard.press('Escape')
    await expect(borders).toBeHidden()

    // Line Numbers: every line counted, then the second paragraph taken out of the count
    await page.locator('.ribbon-tab', { hasText: /^Layout$/ }).click()
    const lineNumbers = page.locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', {
      hasText: 'Line Numbers',
    })
    const numbers = page.locator('.page-line-numbers')
    const counted = async () => (await numbers.allInnerTexts()).join(' ').match(/\d+/g) ?? []
    await lineNumbers.click()
    await page.locator('.layout-menu button', { hasText: /^Continuous$/ }).click()
    await expect.poll(counted).toEqual(['1', '2', '3'])
    await paragraph('banana').click()
    await lineNumbers.click()
    const suppress = page.locator('.layout-menu button', {
      hasText: /^Suppress for Current Paragraph$/,
    })
    await expect(suppress).not.toHaveClass(/active/)
    await page.screenshot({ path: test.info().outputPath('line-numbers.png') })
    await suppress.click()
    await expect.poll(counted).toEqual(['1', '2'])
    await lineNumbers.click()
    await expect(suppress).toHaveClass(/active/)
    await page.keyboard.press('Escape')

    // Page Numbers without the first page: that page gets a footer of its own, with no number
    await clickMenu(app, ['Insert', 'Page Numbers…'])
    const pageNumbers = page.locator('.page-numbers-dialog')
    const showFirst = pageNumbers.locator('input[name="showFirst"]')
    await expect(showFirst).toBeChecked()
    await showFirst.uncheck()
    await page.screenshot({ path: test.info().outputPath('page-numbers.png') })
    await pageNumbers.locator('button.primary').click()
    await expect(pageNumbers).toBeHidden()
    await expect(page.locator('.page-hf-footer', { hasText: '1' })).toHaveCount(0)

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const zip = await JSZip.loadAsync(readFileSync(docPath))
          const xml = await zip.file('word/document.xml')!.async('string')
          return [
            /<w:pgMar\b[^>]* w:left="1440"[^>]* w:gutter="720"/.test(xml),
            /<w:vAlign w:val="center"\/>/.test(xml),
            /<w:cols\b[^>]* w:num="2"/.test(xml),
            /<w:titlePg\/>/.test(xml),
            /<w:lnNumType\b/.test(xml),
            xml.match(/<w:suppressLineNumbers\/>/g)?.length ?? 0,
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([true, true, true, true, true, 1])
  } finally {
    await closeAndSaveVideo(launched, 'docs-dialog-live-controls')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Line Numbering Options, the default tab stop and Custom Paragraph Spacing change the document', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-live-controls-')))
  const docPath = join(dir, 'letter.docx')
  writeFileSync(
    docPath,
    await docxWith(
      ['apple', 'banana'].map(para).join('') +
        '<w:p><w:r><w:tab/><w:t>tabbed</w:t></w:r></w:p>' +
        SECT_PR,
    ),
  )
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-dialog-live-controls',
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const apple = page.locator('.doc-page p', { hasText: 'apple' }).first()
    await expect(apple).toBeVisible({ timeout: 30_000 })
    await apple.click()

    // Line Numbering Options: from 5, every second line shown, never starting over
    await page.locator('.ribbon-tab', { hasText: /^Layout$/ }).click()
    await page
      .locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: 'Line Numbers' })
      .click()
    await page.locator('.layout-menu button', { hasText: /^Line Numbering Options/ }).click()
    const lines = page.locator('.line-numbers-dialog')
    const on = lines.locator('input[name="on"]')
    await expect(on).not.toBeChecked()
    await expect(lines.locator('input[name="start"]')).toBeDisabled()
    await on.check()
    await lines.locator('input[name="start"]').fill('5')
    await lines.locator('input[name="countBy"]').fill('2')
    await lines.locator('input[name="restart"][value="continuous"]').check()
    await page.screenshot({ path: test.info().outputPath('line-numbering-options.png') })
    await lines.locator('button.primary').click()
    await expect(lines).toBeHidden()
    const numbers = page.locator('.page-line-numbers')
    await expect
      .poll(async () => (await numbers.allInnerTexts()).join(' ').match(/\d+/g) ?? [])
      .toEqual(['6'])

    // the Document dialog's Line Numbers… opens the same dialog on what was set
    await clickMenu(app, ['Format', 'Document…'])
    const dialog = page.locator('.doc-dialog:not(.borders-shading-dialog)')
    await dialog.locator('[role="tab"]').nth(1).click()
    await dialog.locator('button', { hasText: /^Line Numbers/ }).click()
    await expect(lines.locator('input[name="on"]')).toBeChecked()
    await expect(lines.locator('input[name="start"]')).toHaveValue('5')
    await expect(lines.locator(':focus')).toHaveCount(1)
    await page.keyboard.press('Escape')
    await expect(lines).toBeHidden()

    // how far into its paragraph the text after the tab starts, px
    const tabbedAt = () =>
      page.evaluate(() => {
        const p = Array.from(document.querySelectorAll('.doc-page p')).find((el) =>
          el.textContent?.includes('tabbed'),
        )
        if (!p) return null
        const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT)
        let text: Node | null = null
        while (walker.nextNode())
          if (walker.currentNode.textContent?.includes('tabbed')) text = walker.currentNode
        if (!text) return null
        const range = document.createRange()
        range.selectNodeContents(text)
        return Math.round(range.getBoundingClientRect().left - p.getBoundingClientRect().left)
      })
    await expect.poll(tabbedAt).toBe(48)

    // Format > Tabs: the document's default tab stop
    await apple.click()
    await clickMenu(app, ['Format', 'Tabs…'])
    const tabs = page.locator('.tabs-dialog')
    const defaultStop = tabs.getByLabel('Default stops')
    await expect(defaultStop).toHaveValue(/^(0\.5"|1\.27 cm)$/)
    await defaultStop.fill('72pt')
    await defaultStop.blur()
    await page.screenshot({ path: test.info().outputPath('tabs-default.png') })
    await tabs.locator('button.primary').click()
    await expect(tabs).toBeHidden()
    // the tab on the page is an inch wide now
    await expect.poll(tabbedAt).toBe(96)

    // Design > Paragraph Spacing > Custom: the default paragraph style's spacing, typed by hand
    await page.locator('.ribbon-tab', { hasText: /^Design$/ }).click()
    await page.locator('.ribbon button', { hasText: /^Paragraph Spacing$/ }).click()
    await page.locator('.layout-menu button', { hasText: /^Custom Paragraph Spacing/ }).click()
    const spacing = page.locator('.paragraph-spacing-dialog')
    await spacing.locator('input[name="before"]').fill('6')
    await spacing.locator('input[name="after"]').fill('18')
    await spacing.locator('input[name="line"]').fill('1.5')
    await page.screenshot({ path: test.info().outputPath('custom-paragraph-spacing.png') })
    await spacing.locator('button.primary').click()
    await expect(spacing).toBeHidden()
    await expect
      .poll(() => apple.evaluate((el) => Math.round(parseFloat(getComputedStyle(el).marginBottom))))
      .toBe(24)

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const zip = await JSZip.loadAsync(readFileSync(docPath))
          const part = async (name: string) => (await zip.file(name)?.async('string')) ?? ''
          const styles = await part('word/styles.xml')
          return [
            /<w:lnNumType w:countBy="2" w:start="4" w:restart="continuous"\/>/.test(
              await part('word/document.xml'),
            ),
            /<w:defaultTabStop w:val="1440"\/>/.test(await part('word/settings.xml')),
            /<w:spacing\b[^>]*w:before="120"/.test(styles),
            /<w:spacing\b[^>]*w:after="360"/.test(styles),
            /<w:spacing\b[^>]*w:line="360"/.test(styles),
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([true, true, true, true, true])
  } finally {
    await closeAndSaveVideo(launched, 'docs-dialog-live-controls')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('the Symbol dialog lists the special characters and inserts them', async () => {
  test.setTimeout(120_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-live-controls-')))
  const docPath = join(dir, 'letter.docx')
  writeFileSync(docPath, await docxWith(para('apple') + SECT_PR))
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-dialog-live-controls',
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const apple = page.locator('.doc-page p', { hasText: 'apple' }).first()
    await expect(apple).toBeVisible({ timeout: 30_000 })
    await apple.click()
    await page.keyboard.press('End')

    await clickMenu(app, ['Insert', 'Advanced Symbol…'])
    const dialog = page.locator('.symbol-dialog')
    await dialog.locator('[role="tab"]', { hasText: 'Special Characters' }).click()
    const rows = dialog.locator('.symbol-special-list button')
    await expect(rows).toHaveCount(18)
    await page.screenshot({ path: test.info().outputPath('special-characters.png') })
    // a double click inserts and leaves the dialog open; Insert takes the picked row
    await rows.filter({ hasText: /Em Dash$/ }).dblclick()
    await rows.filter({ hasText: /Nonbreaking Space$/ }).click()
    await dialog.locator('button.primary').click()
    await rows.filter({ hasText: /Trademark$/ }).click()
    await dialog.locator('button.primary').click()
    await page.keyboard.press('Escape')
    await expect(dialog).toBeHidden()
    await expect.poll(() => apple.evaluate((el) => el.textContent)).toBe('apple\u2014\u00a0\u2122')

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const zip = await JSZip.loadAsync(readFileSync(docPath))
          return (await zip.file('word/document.xml')!.async('string')).includes(
            'apple\u2014\u00a0\u2122',
          )
        },
        { timeout: 30_000 },
      )
      .toBe(true)
  } finally {
    await closeAndSaveVideo(launched, 'docs-dialog-live-controls')
    rmSync(dir, { recursive: true, force: true })
  }
})

const GUTTER_SECT_PR = SECT_PR.replace('w:gutter="0"', 'w:gutter="720"')
const sectionBreak = (sectPr: string) => `<w:p><w:pPr>${sectPr}</w:pPr></w:p>`
/**
 * Puts the caret in a paragraph. The click alone leaves the editor to pick the
 * caret up on its own time; a key press makes it read the caret first, so a
 * dialog opened next acts on this paragraph's section.
 */
async function caretIn(paragraph: Locator) {
  await paragraph.click()
  await paragraph.page().keyboard.press('End')
}
const partOf = async (docPath: string, part: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file(part)?.async('string') ?? ''

test('the Document dialog moves the gutter to the top and changes how a section starts', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-live-controls-')))
  const docPath = join(dir, 'sections.docx')
  writeFileSync(
    docPath,
    await docxWith(
      para('apple') +
        sectionBreak(GUTTER_SECT_PR) +
        para('banana') +
        sectionBreak(GUTTER_SECT_PR) +
        para('cherry') +
        GUTTER_SECT_PR,
    ),
  )
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-dialog-live-controls',
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const paragraph = (text: string) => page.locator('.doc-page p', { hasText: text }).first()
    await expect(paragraph('cherry')).toBeVisible({ timeout: 30_000 })
    const pageTurns = page.locator('.doc-page .page-gap')
    await expect(pageTurns).toHaveCount(2)
    // where a paragraph's text starts on the paper, in px of an unzoomed page
    const inset = (text: string) =>
      page.evaluate((text) => {
        const paper = document.querySelector('.doc-page')!.getBoundingClientRect()
        const p = Array.from(document.querySelectorAll('.doc-page p')).find((el) =>
          el.textContent?.includes(text),
        )!
        const r = p.getBoundingClientRect()
        const k = 816 / paper.width
        return {
          left: Math.round((r.left - paper.left) * k),
          top: Math.round((r.top - paper.top) * k),
        }
      }, text)
    const before = await inset('apple')
    expect(before.left).toBe(144)

    // the last section: the gutter of every section goes to the top, and it starts without a new page
    await caretIn(paragraph('cherry'))
    const dialog = page.locator('.doc-dialog:not(.borders-shading-dialog)')
    const tab = (index: number) => dialog.locator('[role="tab"]').nth(index)
    const position = dialog.locator('select[name="gutterPosition"]')
    const start = dialog.locator('select[name="sectionStart"]')
    await clickMenu(app, ['Format', 'Document…'])
    await expect(dialog).toHaveAttribute('data-section', '2')
    await expect(position).toHaveValue('left')
    await position.selectOption('top')
    await page.screenshot({ path: test.info().outputPath('gutter-position.png') })
    await tab(1).click()
    await expect(start).toHaveValue('nextPage')
    await expect(start.locator('option')).toHaveCount(5)
    await start.selectOption('continuous')
    await page.screenshot({ path: test.info().outputPath('section-start.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect(pageTurns).toHaveCount(1)
    await expect.poll(() => inset('apple')).toEqual({ left: 96, top: before.top + 48 })

    // a section in the middle of the document
    await caretIn(paragraph('banana'))
    await clickMenu(app, ['Format', 'Document…'])
    await expect(dialog).toHaveAttribute('data-section', '1')
    await expect(position).toHaveValue('top')
    await tab(1).click()
    await expect(start).toHaveValue('nextPage')
    await start.selectOption('continuous')
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect(pageTurns).toHaveCount(0)

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const xml = await partOf(docPath, 'word/document.xml')
          const sections = xml.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g) ?? []
          return [
            /<w:gutterAtTop\/>/.test(await partOf(docPath, 'word/settings.xml')),
            sections.map((sectPr) => /<w:type w:val="continuous"\/>/.test(sectPr)),
            sections.every((sectPr) =>
              /<w:pgMar\b[^>]* w:top="1440"[^>]* w:left="1440"[^>]* w:gutter="720"/.test(sectPr),
            ),
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([true, [false, true, true], true])

    // read back from the saved file
    await caretIn(paragraph('cherry'))
    await clickMenu(app, ['Format', 'Document…'])
    await expect(dialog).toHaveAttribute('data-section', '2')
    await expect(position).toHaveValue('top')
    await tab(1).click()
    await expect(start).toHaveValue('continuous')
    await page.keyboard.press('Escape')
    await expect(dialog).toBeHidden()
    await expect(pageTurns).toHaveCount(0)
    await expect.poll(() => inset('apple')).toEqual({ left: 96, top: before.top + 48 })
  } finally {
    await closeAndSaveVideo(launched, 'docs-dialog-live-controls')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('the Columns dialog sets unequal columns, the line between them and equal columns again', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-live-controls-')))
  const docPath = join(dir, 'columns.docx')
  const lines = Array.from({ length: 30 }, (_, i) => `line ${String(i + 1).padStart(2, '0')}`)
  writeFileSync(docPath, await docxWith(lines.map(para).join('') + SECT_PR))
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-dialog-live-controls',
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const first = page.locator('.doc-page p', { hasText: 'line 01' }).first()
    await expect(first).toBeVisible({ timeout: 30_000 })
    await caretIn(first)
    const cols = () =>
      partOf(docPath, 'word/document.xml').then(
        (xml) => /<w:cols\b[^>]*(?:\/>|>[\s\S]*?<\/w:cols>)/.exec(xml)?.[0],
      )
    // the rules and the paragraphs on the paper, in px of an unzoomed page
    const layout = () =>
      page.evaluate(() => {
        const paper = document.querySelector('.doc-page')!.getBoundingClientRect()
        const k = 816 / paper.width
        const box = (el: Element) => {
          const r = el.getBoundingClientRect()
          return {
            left: Math.round((r.left - paper.left) * k),
            top: Math.round((r.top - paper.top) * k),
            bottom: Math.round((r.bottom - paper.top) * k),
          }
        }
        const paras = Array.from(document.querySelectorAll('.doc-page p'))
          .filter((el) => el.textContent?.startsWith('line '))
          .map(box)
        return {
          rules: Array.from(document.querySelectorAll('.page-column-rule')).map(box),
          lefts: [...new Set(paras.map((p) => p.left))].sort((a, b) => a - b),
          top: Math.min(...paras.map((p) => p.top)),
          bottom: Math.max(...paras.map((p) => p.bottom)),
        }
      })

    const dialog = page.locator('.columns-dialog')
    const line = dialog.locator('input[name="lineBetween"]')
    const equal = dialog.locator('input[name="equalWidth"]')
    await clickMenu(app, ['Format', 'Columns…'])
    await expect(line).toBeDisabled()
    await expect(equal).toBeChecked()
    await expect(equal).toBeDisabled()
    // Left: two columns, the narrow one first, with a line between them
    await dialog.locator('input[name="columns"][value="left"]').check()
    await expect(equal).not.toBeChecked()
    await line.check()
    await page.screenshot({ path: test.info().outputPath('columns-left.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    // 1" margin, a 176 px column, a 48 px gap with the line in its middle
    await expect.poll(async () => (await layout()).lefts).toEqual([96, 320])
    await expect
      .poll(async () => {
        const { rules, top, bottom } = await layout()
        return (
          rules.length === 1 && [
            rules[0].left,
            Math.abs(rules[0].top - top) <= 3,
            Math.abs(rules[0].bottom - bottom) <= 3,
          ]
        )
      })
      .toEqual([296, true, true])
    await page.screenshot({ path: test.info().outputPath('columns-left-page.png') })
    await saveFromMenu(app)
    await expect
      .poll(cols, { timeout: 30_000 })
      .toBe(
        '<w:cols w:num="2" w:space="720" w:equalWidth="0" w:sep="1"><w:col w:w="2640" w:space="720"/><w:col w:w="6000"/></w:cols>',
      )

    // the sheets that print carry the line too
    await page.evaluate(() =>
      (window as unknown as { __openPagePreview?: () => void }).__openPagePreview?.(),
    )
    await expect(page.locator('.pv-page')).toHaveCount(1, { timeout: 30_000 })
    await expect(page.locator('.pv-column-rule')).toHaveCount(1)
    await page.keyboard.press('Escape')
    await expect(page.locator('.pv-page')).toHaveCount(0)

    // read back, then equal columns again: the line moves to the middle of the page
    await caretIn(first)
    await clickMenu(app, ['Format', 'Columns…'])
    await expect(dialog.locator('input[name="columns"][value="left"]')).toBeChecked()
    await expect(line).toBeChecked()
    await expect(equal).not.toBeChecked()
    await equal.check()
    await expect(dialog.locator('input[name="columns"]').nth(1)).toBeChecked()
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect.poll(async () => (await layout()).lefts).toEqual([96, 432])
    await expect.poll(async () => (await layout()).rules.map((rule) => rule.left)).toEqual([408])
    await saveFromMenu(app)
    await expect.poll(cols, { timeout: 30_000 }).toBe('<w:cols w:num="2" w:space="720" w:sep="1"/>')

    // without the line
    await caretIn(first)
    await clickMenu(app, ['Format', 'Columns…'])
    await expect(line).toBeChecked()
    await line.uncheck()
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect(page.locator('.page-column-rule')).toHaveCount(0)
    await saveFromMenu(app)
    await expect.poll(cols, { timeout: 30_000 }).toBe('<w:cols w:num="2" w:space="720"/>')
  } finally {
    await closeAndSaveVideo(launched, 'docs-dialog-live-controls')
    rmSync(dir, { recursive: true, force: true })
  }
})
