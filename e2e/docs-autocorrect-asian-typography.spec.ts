/**
 * Tools > AutoCorrect Options… with Word's four tabs (AutoCorrect, Math
 * AutoCorrect, AutoFormat As You Type, AutoText) and the Asian Typography tab
 * of Format > Paragraph…: what the dialogs switch shows while typing and is in
 * the saved file.
 */
import { test, expect, type ElectronApplication, type Locator, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-autocorrect-asian-typography'
const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '</w:sectPr>'
const FLAGS = /<w:(?:wordWrap|overflowPunct|autoSpaceD[EN])[^>]*>/g
const MIXED = '\u4e2d\u6587abc\u4e2d\u6587123\u4e2d\u6587'

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

async function caretAtEnd(paragraph: Locator) {
  await paragraph.click()
  await paragraph.page().keyboard.press('End')
}

/** puts the caret in front of the first character of the paragraph that starts with `text` */
const caretAtStartOf = (page: Page, text: string) =>
  page.evaluate((text) => {
    type Node = { isTextblock: boolean; textContent: string }
    const { editor } = (
      window as never as {
        __aidocs: {
          editor: {
            state: { doc: { descendants(f: (node: Node, pos: number) => boolean): void } }
            chain(): { focus(): { setTextSelection(pos: number): { run(): boolean } } }
          }
        }
      }
    ).__aidocs
    let at = -1
    editor.state.doc.descendants((node, pos) => {
      if (at < 0 && node.isTextblock && node.textContent.startsWith(text)) at = pos + 1
      return at < 0
    })
    editor.chain().focus().setTextSelection(at).run()
  }, text)

/** the End key scrolls the window on macOS instead of moving the caret */
const caretAtDocEnd = (page: Page) =>
  page.evaluate(() => {
    const { editor } = (
      window as never as { __aidocs: { editor: { commands: { focus(at: string): boolean } } } }
    ).__aidocs
    editor.commands.focus('end')
  })

const documentXml = async (docPath: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file('word/document.xml')?.async('string') ?? ''

const savedParagraph = async (docPath: string, text: string) =>
  (await documentXml(docPath)).split('</w:p>').find((p) => p.includes(text)) ?? ''

async function openAutoCorrect(app: ElectronApplication, page: Page, tab?: string) {
  await clickMenu(app, ['Tools', 'AutoCorrect Options…'])
  const dialog = page.getByRole('dialog', { name: 'AutoCorrect Options' })
  await expect(dialog).toBeVisible()
  if (tab) await dialog.getByRole('tab', { name: tab, exact: true }).click()
  return dialog
}

async function closeDialog(dialog: Locator) {
  await dialog.getByRole('button', { name: 'Close', exact: true }).click()
  await expect(dialog).toBeHidden()
}

test('AutoCorrect Options: the replace table, exceptions, math names, AutoFormat switches and AutoText', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-autocorrect-')))
  const docPath = join(dir, 'typing.docx')
  writeFileSync(docPath, await docxWith(para('Kind regards from the team') + para('') + SECT_PR))
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const first = page.locator('.doc-page p', { hasText: 'Kind regards' })
    await expect(first).toBeVisible({ timeout: 30_000 })
    const last = page.locator('.doc-page p').last()

    // AutoCorrect tab: a new entry in the Replace / With table, and an exception
    let dialog = await openAutoCorrect(app, page)
    await expect(dialog.getByRole('tab')).toHaveText([
      'AutoCorrect',
      'Math AutoCorrect',
      'AutoFormat As You Type',
      'AutoText',
    ])
    await dialog.getByRole('textbox', { name: 'Replace:' }).fill('teh')
    await dialog.getByRole('textbox', { name: 'With:' }).fill('the')
    await dialog.getByRole('button', { name: 'Add', exact: true }).click()
    await expect(dialog.getByRole('option', { name: 'teh the' })).toBeVisible()
    await dialog.getByRole('button', { name: 'Exceptions…' }).click()
    const exceptions = page.getByRole('dialog', { name: 'AutoCorrect Exceptions' })
    await exceptions.getByRole('tab', { name: 'INitial CAps' }).click()
    await exceptions.getByRole('textbox', { name: "Don't correct:" }).fill('MYsql')
    await exceptions.getByRole('button', { name: 'Add', exact: true }).click()
    await expect(exceptions.getByRole('option', { name: 'MYsql' })).toBeVisible()
    await page.screenshot({ path: test.info().outputPath('autocorrect-exceptions.png') })
    await exceptions.getByRole('button', { name: 'Close', exact: true }).click()
    await expect(exceptions).toBeHidden()
    await page.screenshot({ path: test.info().outputPath('autocorrect-tab.png') })

    // Math AutoCorrect tab: names work in body text once switched on
    await dialog.getByRole('tab', { name: 'Math AutoCorrect' }).click()
    await dialog
      .getByLabel('Use Math AutoCorrect rules outside of math regions')
      .check({ force: true })
    await page.screenshot({ path: test.info().outputPath('autocorrect-math.png') })

    // AutoFormat As You Type tab
    await dialog.getByRole('tab', { name: 'AutoFormat As You Type' }).click()
    await dialog.getByLabel('*Bold* and _italic_ with real formatting').check({ force: true })
    await expect(dialog.getByLabel('Automatic bulleted lists')).toBeChecked()
    await expect(dialog.getByLabel('Border lines')).toBeChecked()
    await expect(dialog.getByLabel('Date style')).toHaveCount(0)
    await page.screenshot({ path: test.info().outputPath('autocorrect-autoformat.png') })
    await closeDialog(dialog)

    await caretAtEnd(last)
    await page.keyboard.type('THe MYsql teh (c) \\alpha *bold* on monday ')
    await expect(last).toHaveText('The MYsql the \u00a9 \u03b1 bold on Monday')
    await page.keyboard.press('Enter')
    await page.keyboard.type('---')
    await page.keyboard.press('Enter')
    await page.keyboard.type('+--+--+--+')
    await page.keyboard.press('Enter')
    await expect(page.locator('.doc-page table')).toHaveCount(1)
    await expect(page.locator('.doc-page table tr').first().locator('td')).toHaveCount(3)

    // the master switch stops everything
    dialog = await openAutoCorrect(app, page)
    await dialog
      .getByLabel('Automatically correct spelling and formatting as you type')
      .uncheck({ force: true })
    await expect(dialog.getByLabel('Correct TWo INitial CApitals')).toBeDisabled()
    await closeDialog(dialog)
    await caretAtEnd(first)
    await page.keyboard.type(' (c) teh ')
    await expect(first).toHaveText('Kind regards from the team (c) teh')

    // AutoText tab: keep the selection under a name, then put it back in
    await first
      .locator('text=Kind')
      .first()
      .dblclick({ position: { x: 6, y: 8 } })
    dialog = await openAutoCorrect(app, page, 'AutoText')
    await dialog.getByRole('textbox', { name: 'Enter AutoText entries here:' }).fill('greeting')
    await dialog.getByRole('button', { name: 'Add', exact: true }).click()
    await expect(dialog.getByRole('option', { name: 'greeting' })).toBeVisible()
    await page.screenshot({ path: test.info().outputPath('autocorrect-autotext.png') })
    await closeDialog(dialog)
    await caretAtEnd(first)
    dialog = await openAutoCorrect(app, page, 'AutoText')
    await dialog.getByRole('option', { name: 'greeting' }).click()
    await dialog.getByRole('button', { name: 'Insert', exact: true }).click()
    await expect(dialog).toBeHidden()
    await expect(first).toHaveText(/teh ?Kind ?$/)

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const xml = await documentXml(docPath)
          const typed = await savedParagraph(docPath, 'MYsql')
          return {
            typed: typed.replace(/<[^>]+>/g, '').trim(),
            bold: /<w:b\/>(?:(?!<\/w:r>).)*<w:t[^>]*>bold</.test(typed),
            border: /<w:pBdr><w:bottom w:val="single" w:sz="6"/.test(xml),
            columns: (/<w:tblGrid>(.*?)<\/w:tblGrid>/.exec(xml)?.[1].match(/<w:gridCol/g) ?? [])
              .length,
          }
        },
        { timeout: 20_000 },
      )
      .toEqual({
        typed: 'The MYsql the \u00a9 \u03b1 bold on Monday',
        bold: true,
        border: true,
        columns: 3,
      })
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('AutoFormat switches that were listed disabled: indents, brackets, the AutoComplete tip and formatted entries', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-autoformat-')))
  const docPath = join(dir, 'autoformat.docx')
  const opening =
    '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Kind regards</w:t></w:r>' +
    '<w:r><w:t xml:space="preserve"> from the team</w:t></w:r></w:p>'
  writeFileSync(docPath, await docxWith(opening + para('') + SECT_PR))
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const first = page.locator('.doc-page p', { hasText: 'Kind regards' })
    await expect(first).toBeVisible({ timeout: 30_000 })

    // a selection is offered as the replacement, kept as formatted text
    await page.evaluate(() => {
      const { editor } = (
        window as never as {
          __aidocs: {
            editor: { chain(): { focus(): { setTextSelection(r: object): { run(): boolean } } } }
          }
        }
      ).__aidocs
      editor.chain().focus().setTextSelection({ from: 1, to: 13 }).run()
    })
    let dialog = await openAutoCorrect(app, page)
    await expect(dialog.getByRole('textbox', { name: 'With:' })).toHaveValue('Kind regards')
    await expect(dialog.getByLabel('Formatted text')).toBeChecked()
    await dialog.getByRole('textbox', { name: 'Replace:' }).fill('kr')
    await dialog.getByRole('button', { name: 'Add', exact: true }).click()
    await expect(dialog.getByRole('option', { name: 'kr Kind regards' })).toBeVisible()
    // picking the entry brings its stored formatting back, so a copy under another name keeps it
    await dialog.getByRole('textbox', { name: 'With:' }).fill('')
    await dialog.getByRole('option', { name: 'kr Kind regards' }).click()
    await expect(dialog.getByLabel('Formatted text')).toBeChecked()
    await dialog.getByRole('textbox', { name: 'Replace:' }).fill('kr2')
    await dialog.getByRole('button', { name: 'Add', exact: true }).click()
    await dialog.getByRole('tab', { name: 'AutoFormat As You Type' }).click()
    for (const label of [
      'Spaces at beginning of paragraph with first-line indent',
      'Match opening and closing parentheses',
      'Set left- and first-indent with tabs and backspaces',
      'Format beginning of list item like the one before it',
      'Insert closing phrase to match memo style',
    ])
      await expect(dialog.getByLabel(label)).toBeChecked()
    await expect(dialog.getByLabel('Built-in Heading styles')).not.toBeChecked()
    await page.screenshot({ path: test.info().outputPath('autoformat-live.png') })
    await dialog.getByRole('tab', { name: 'AutoText', exact: true }).click()
    await expect(dialog.getByLabel('Show AutoComplete tip for AutoText and dates')).toBeChecked()
    await closeDialog(dialog)

    const last = page.locator('.doc-page p').last()
    await caretAtEnd(last)
    await page.keyboard.type('  spaced kr2, \uff08note)')
    await page.keyboard.press('Enter')
    await page.keyboard.type('tabbed in Nove')
    const tip = page.locator('.doc-autocomplete-tip')
    await expect(tip).toHaveText('November (Press Return to Insert)')
    await page.screenshot({ path: test.info().outputPath('autocomplete-tip.png') })
    await page.keyboard.press('Enter')
    await expect(tip).toHaveCount(0)
    await caretAtStartOf(page, 'tabbed')
    await page.keyboard.press('Tab')
    await page.keyboard.press('Tab')
    await expect(page.locator('.doc-page p', { hasText: 'tabbed' })).toHaveText(
      'tabbed in November',
    )

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const spaced = await savedParagraph(docPath, 'spaced')
          const tabbed = await savedParagraph(docPath, 'tabbed')
          return {
            spaced: spaced.replace(/<[^>]+>/g, ''),
            spacedBold: /<w:b\/>(?:(?!<\/w:r>).)*<w:t[^>]*>Kind regards</.test(spaced),
            spacedIndent: /<w:ind [^>]*w:firstLineChars="100"/.test(spaced),
            tabbed: tabbed.replace(/<[^>]+>/g, ''),
            // the new paragraph kept the first-line indent, so both Tabs moved the left indent
            tabbedFirst: /<w:ind [^>]*w:firstLineChars="100"/.test(tabbed),
            tabbedLeft: /<w:ind [^>]*w:left="1440"/.test(tabbed),
          }
        },
        { timeout: 20_000 },
      )
      .toEqual({
        spaced: 'spaced Kind regards, \uff08note\uff09',
        spacedBold: true,
        spacedIndent: true,
        tabbed: 'tabbed in November',
        tabbedFirst: true,
        tabbedLeft: true,
      })

    // switched off, Tab types a tab again
    dialog = await openAutoCorrect(app, page, 'AutoFormat As You Type')
    await dialog
      .getByLabel('Set left- and first-indent with tabs and backspaces')
      .uncheck({ force: true })
    await closeDialog(dialog)
    await first.first().click()
    await caretAtStartOf(page, 'Kind regards')
    await page.keyboard.press('Tab')
    await saveFromMenu(app)
    await expect
      .poll(async () => /<w:tab\/>/.test(await savedParagraph(docPath, 'from the team')), {
        timeout: 20_000,
      })
      .toBe(true)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Paragraph > Asian Typography switches hanging punctuation and the automatic spacing, and saves them', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-asian-typography-')))
  const docPath = join(dir, 'asian.docx')
  writeFileSync(docPath, await docxWith(para(MIXED) + para('plain') + SECT_PR))
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const mixed = page.locator('.doc-page p', { hasText: 'abc' })
    await expect(mixed).toBeVisible({ timeout: 30_000 })
    const pads = () => mixed.locator('[class*="doc-autospace-pad"]').count()
    const padsBefore = await pads()
    expect(padsBefore).toBeGreaterThan(0)

    await caretAtEnd(mixed)
    await clickMenu(app, ['Format', 'Paragraph…'])
    const dialog = page.locator('.modal', { hasText: 'Paragraph' })
    await dialog.getByRole('tab', { name: 'Asian Typography' }).click()
    const hanging = dialog.getByLabel('Allow hanging punctuation')
    const latin = dialog.getByLabel('Automatically adjust space between Asian and Latin text')
    const digits = dialog.getByLabel('Automatically adjust space between Asian text and numbers')
    await expect(hanging).toBeChecked()
    await expect(latin).toBeChecked()
    await expect(digits).toBeChecked()
    await hanging.uncheck()
    await digits.uncheck()
    await dialog.getByLabel('Allow Latin text to wrap in the middle of a word').check()
    // Options…: the document-wide character spacing control
    await dialog.getByRole('button', { name: 'Options…' }).click()
    const typography = page.getByRole('dialog', { name: 'Typography' })
    await expect(typography.getByLabel('Do not compress')).toBeChecked()
    // Escape closes Typography only: the Paragraph dialog and its pending changes stay
    await expect(typography.getByLabel('Standard')).toBeFocused()
    await page.keyboard.press('Escape')
    await expect(typography).toBeHidden()
    await expect(hanging).not.toBeChecked()
    await dialog.getByRole('button', { name: 'Options…' }).click()
    await expect(typography.getByLabel('Custom')).toHaveCount(0)
    await typography.getByLabel('Compress punctuation only').check()
    await page.screenshot({ path: test.info().outputPath('typography-options.png') })
    await typography.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(typography).toBeHidden()
    await page.screenshot({ path: test.info().outputPath('paragraph-asian-typography.png') })
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog).toBeHidden()
    // the gaps at the digits are gone, the ones at the letters stay
    await expect.poll(pads).toBeLessThan(padsBefore)
    await expect.poll(pads).toBeGreaterThan(0)

    await saveFromMenu(app)
    await expect
      .poll(async () => (await savedParagraph(docPath, 'abc')).match(FLAGS)?.join(''), {
        timeout: 20_000,
      })
      .toBe('<w:wordWrap w:val="0"/><w:overflowPunct w:val="0"/><w:autoSpaceDN w:val="0"/>')
    const settings =
      (await JSZip.loadAsync(readFileSync(docPath))).file('word/settings.xml')?.async('string') ??
      ''
    expect(await settings).toContain('<w:characterSpacingControl w:val="compressPunctuation"/>')
    expect(await savedParagraph(docPath, 'plain')).not.toMatch(/overflowPunct|autoSpace/)

    // the dialog opens on what was saved; switching the spacing back on writes it explicitly
    await caretAtEnd(mixed)
    await clickMenu(app, ['Format', 'Paragraph…'])
    await dialog.getByRole('tab', { name: 'Asian Typography' }).click()
    await expect(hanging).not.toBeChecked()
    await expect(
      dialog.getByLabel('Allow Latin text to wrap in the middle of a word'),
    ).toBeChecked()
    await dialog.getByRole('button', { name: 'Options…' }).click()
    await expect(typography.getByLabel('Compress punctuation only')).toBeChecked()
    await typography.getByRole('button', { name: 'Cancel', exact: true }).click()
    await expect(latin).toBeChecked()
    await expect(digits).not.toBeChecked()
    await digits.check()
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect.poll(pads).toBe(padsBefore)
    await saveFromMenu(app)
    await expect
      .poll(async () => (await savedParagraph(docPath, 'abc')).match(FLAGS)?.join(''), {
        timeout: 20_000,
      })
      .toBe('<w:wordWrap w:val="0"/><w:overflowPunct w:val="0"/><w:autoSpaceDN/>')
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Asian Typography: the first and last character rules can be switched off, line-start brackets compress', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-kinsoku-')))
  const docPath = join(dir, 'kinsoku.docx')
  const HAN = '\u4e2d'
  // one of these lengths puts the full stop exactly where the second line starts
  const stops = Array.from({ length: 36 }, (_, i) => para(HAN.repeat(22 + i) + '\u3002' + HAN))
  writeFileSync(docPath, await docxWith(para('\uff08' + HAN.repeat(4)) + stops.join('') + SECT_PR))
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const bracket = page.locator('.doc-page p').first()
    await expect(bracket).toBeVisible({ timeout: 30_000 })

    /** [left edge of the character after the opening bracket in em, full stops that start a line] */
    const measure = () =>
      page.evaluate(() => {
        const charRect = (node: Text, i: number) => {
          const range = document.createRange()
          range.setStart(node, i)
          range.setEnd(node, i + 1)
          return range.getBoundingClientRect()
        }
        const paras = [...document.querySelectorAll<HTMLElement>('.doc-page p')]
        const texts = (el: HTMLElement) => {
          const out: Text[] = []
          const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT)
          while (walker.nextNode()) out.push(walker.currentNode as Text)
          return out
        }
        const first = paras[0]
        const chars = texts(first).flatMap((t) => [...t.data].map((_, i) => charRect(t, i)))
        const em = parseFloat(getComputedStyle(first).fontSize)
        let lineStartStops = 0
        for (const para of paras.slice(1)) {
          const rects = texts(para).flatMap((t) =>
            [...t.data].map((ch, i) => ({ ch, rect: charRect(t, i) })),
          )
          rects.forEach((c, i) => {
            if (c.ch === '\u3002' && i > 0 && rects[i - 1].rect.bottom <= c.rect.top + 1)
              lineStartStops++
          })
        }
        return {
          second: (chars[1].left - first.getBoundingClientRect().left) / em,
          lineStartStops,
        }
      })
    const before = await measure()
    expect(before.lineStartStops).toBe(0)

    await bracket.click()
    await page.keyboard.press('ControlOrMeta+a')
    await clickMenu(app, ['Format', 'Paragraph…'])
    const dialog = page.locator('.modal', { hasText: 'Paragraph' })
    await dialog.getByRole('tab', { name: 'Asian Typography' }).click()
    const rules = dialog.getByLabel('Use Asian rules for controlling first and last characters')
    const compress = dialog.getByLabel('Allow punctuation at the start of the line to compress')
    await expect(rules).toBeChecked()
    await expect(compress).not.toBeChecked()
    await rules.uncheck()
    await compress.check()
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog).toBeHidden()

    // Word: the bracket gives up half an em; a full stop may now start a line
    await expect.poll(async () => before.second - (await measure()).second).toBeGreaterThan(0.4)
    const after = await measure()
    expect(before.second - after.second).toBeLessThan(0.6)
    expect(after.lineStartStops).toBeGreaterThan(0)
    await page.screenshot({ path: test.info().outputPath('kinsoku-off.png') })

    await saveFromMenu(app)
    await expect
      .poll(
        async () =>
          (await savedParagraph(docPath, '\uff08')).match(/<w:(?:kinsoku|topLinePunct)[^>]*>/g),
        { timeout: 20_000 },
      )
      .toEqual(['<w:kinsoku w:val="0"/>', '<w:topLinePunct/>'])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a line-start bracket keeps its line when compression makes it narrow enough to fit the line above', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-topline-')))
  const docPath = join(dir, 'topline.docx')
  const HAN = '\u4e2d'
  // every leftover width in front of the bracket, in quarter-em steps
  const paras: string[] = []
  for (let n = 30; n < 44; n++)
    for (let narrow = 0; narrow < 4; narrow++)
      paras.push(para(HAN.repeat(n) + 'i'.repeat(narrow) + '\uff08' + HAN.repeat(4)))
  writeFileSync(docPath, await docxWith(paras.join('') + SECT_PR))
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const first = page.locator('.doc-page p').first()
    await expect(first).toBeVisible({ timeout: 30_000 })

    /** per paragraph: does the bracket start a line, and how far left of the text edge it sits (em) */
    const brackets = () =>
      page.evaluate(() =>
        [...document.querySelectorAll<HTMLElement>('.doc-page p')].map((para) => {
          const rects: Array<{ ch: string; rect: DOMRect }> = []
          const walker = document.createTreeWalker(para, NodeFilter.SHOW_TEXT)
          while (walker.nextNode()) {
            const node = walker.currentNode as Text
            for (let i = 0; i < node.data.length; i++) {
              const range = document.createRange()
              range.setStart(node, i)
              range.setEnd(node, i + 1)
              rects.push({ ch: node.data[i], rect: range.getBoundingClientRect() })
            }
          }
          const at = rects.findIndex((c) => c.ch === '\uff08')
          const em = parseFloat(getComputedStyle(para).fontSize)
          return {
            starts: rects[at - 1].rect.bottom <= rects[at].rect.top + 1,
            hang: (rects[0].rect.left - rects[at].rect.left) / em,
          }
        }),
      )
    const before = await brackets()
    const starting = before.filter((b) => b.starts).length
    expect(starting).toBeGreaterThan(3)

    await first.click()
    await page.keyboard.press('ControlOrMeta+a')
    await clickMenu(app, ['Format', 'Paragraph…'])
    const dialog = page.locator('.modal', { hasText: 'Paragraph' })
    await dialog.getByRole('tab', { name: 'Asian Typography' }).click()
    await dialog.getByLabel('Allow punctuation at the start of the line to compress').check()
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog).toBeHidden()

    // the same brackets start a line as before, and every one of them hangs half an em
    await expect
      .poll(async () => (await brackets()).filter((b) => b.starts && b.hang > 0.4).length)
      .toBe(starting)
    const after = await brackets()
    expect(after.map((b) => b.starts)).toEqual(before.map((b) => b.starts))
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('the AutoCorrect smart button changes a correction back or stops it; the equation box reads Math AutoCorrect and recognized functions', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-smart-button-')))
  const docPath = join(dir, 'smart.docx')
  writeFileSync(docPath, await docxWith(para('') + SECT_PR))
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const line = page.locator('.doc-page p').first()
    await expect(line).toBeVisible({ timeout: 30_000 })
    await line.click()
    await page.keyboard.type('THe (c) end')
    await expect(line).toHaveText('The \u00a9 end')

    const fixed = page.locator('.doc-page .doc-ac-fix')
    await expect(fixed).toHaveCount(2)
    const button = page.locator('.ac-smart-button')
    const menu = page.locator('.ac-smart-menu')
    await fixed.first().hover()
    await expect(button).toBeVisible()
    await button.click()
    await expect(menu.getByRole('menuitem')).toHaveText([
      'Change Back to "THe"',
      'Stop Automatically Correcting "THe"',
      'Control AutoCorrect Options…',
    ])
    await page.screenshot({ path: test.info().outputPath('smart-button.png') })
    await menu.getByRole('menuitem', { name: 'Change Back to "THe"' }).click()
    await expect(line).toHaveText('THe \u00a9 end')
    await expect(fixed).toHaveCount(1)

    await fixed.first().hover()
    await button.click()
    await menu.getByRole('menuitem', { name: 'Stop Automatically Correcting "(c)"' }).click()
    await expect(line).toHaveText('THe (c) end')
    await caretAtDocEnd(page)
    await page.keyboard.type(' (c) (r)')
    await expect(line).toHaveText('THe (c) end (c) \u00ae')

    // two corrections side by side: the menu is for the one under the pointer
    await page.keyboard.type('(tm)')
    await expect(line).toHaveText('THe (c) end (c) \u00ae\u2122')
    await page.locator('.doc-page .doc-ac-fix', { hasText: '\u2122' }).hover()
    await button.click()
    await menu.getByRole('menuitem', { name: 'Change Back to "(tm)"' }).click()
    await expect(line).toHaveText('THe (c) end (c) \u00ae(tm)')

    // the menu's last item opens the dialog; the switch takes the marks away
    await page.locator('.doc-page .doc-ac-fix').first().hover()
    await button.click()
    await menu.getByRole('menuitem', { name: 'Control AutoCorrect Options…' }).click()
    const dialog = page.getByRole('dialog', { name: 'AutoCorrect Options' })
    await dialog.getByLabel('Show AutoCorrect smart button').uncheck({ force: true })
    await dialog.getByRole('tab', { name: 'Math AutoCorrect' }).click()
    await expect(dialog.getByLabel('Replace text as you type')).toBeChecked()
    // the table is editable for the equation box alone, with body-text replacement off
    await expect(
      dialog.getByLabel('Use Math AutoCorrect rules outside of math regions'),
    ).not.toBeChecked()
    await expect(dialog.getByRole('textbox', { name: 'Replace:' })).toBeEnabled()
    await dialog.getByRole('button', { name: 'Recognized Functions…' }).click()
    const functions = page.getByRole('dialog', { name: 'Recognized Math Functions' })
    await expect(functions.getByRole('option', { name: 'sin', exact: true })).toBeVisible()
    await functions.getByRole('textbox').fill('erf')
    await functions.getByRole('button', { name: 'Add', exact: true }).click()
    await expect(functions.getByRole('option', { name: 'erf', exact: true })).toBeVisible()
    await page.screenshot({ path: test.info().outputPath('recognized-functions.png') })
    await functions.getByRole('button', { name: 'Close', exact: true }).click()
    await closeDialog(dialog)
    await expect(page.locator('.doc-page .doc-ac-fix')).toHaveCount(0)

    // Insert > Equation…: typed bare, the functions and the symbol become what the equation reads
    await caretAtDocEnd(page)
    await clickMenu(app, ['Insert', 'Equation…'])
    const equation = page.locator('.equation-modal')
    await equation.locator('input').first().pressSequentially('erf(x) <= sin y ')
    await expect(equation.locator('input').first()).toHaveValue(
      '\\operatorname{erf}(x) \u2264 \\sin y ',
    )
    await equation.locator('input').first().press('Enter')
    await expect(equation).toBeHidden()
    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const xml = await documentXml(docPath)
          const math = /<m:oMath>.*<\/m:oMath>/.exec(xml)?.[0] ?? ''
          return {
            upright: (math.match(/<m:sty m:val="p"\/><\/m:rPr><m:t[^>]*>(\w+)</g) ?? []).length,
            text: math.replace(/<[^>]+>/g, ''),
          }
        },
        { timeout: 20_000 },
      )
      .toEqual({ upright: 2, text: 'erf(x) \u2264 siny ' })
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Typography > Strict keeps small kana off the line start in Japanese text; Text Alignment moves the smaller runs', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-strict-')))
  const docPath = join(dir, 'strict.docx')
  const KANA = '\u3042'
  const SMALL = '\u3083'
  const ja = (text: string) =>
    `<w:p><w:r><w:rPr><w:lang w:eastAsia="ja-JP"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`
  const sized = (text: string, halfPoints: number) =>
    `<w:r><w:rPr><w:sz w:val="${halfPoints}"/></w:rPr><w:t>${text}</w:t></w:r>`
  // the smaller run after the bigger one, and before it
  const mixed =
    `<w:p>${sized('Big', 56)}${sized('small', 20)}</w:p>` +
    `<w:p>${sized('tiny', 20)}${sized('Huge', 56)}</w:p>`
  // one of these lengths puts the small kana exactly where the second line starts
  const sweep = Array.from({ length: 36 }, (_, i) =>
    ja(KANA.repeat(22 + i) + SMALL + KANA.repeat(3)),
  )
  writeFileSync(docPath, await docxWith(mixed + sweep.join('') + SECT_PR))
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const first = page.locator('.doc-page p').first()
    await expect(first).toBeVisible({ timeout: 30_000 })

    const measure = () =>
      page.evaluate(() => {
        const rectsOf = (para: Element) => {
          const out: Array<{ ch: string; rect: DOMRect }> = []
          const walker = document.createTreeWalker(para, NodeFilter.SHOW_TEXT)
          while (walker.nextNode()) {
            const node = walker.currentNode as Text
            for (let i = 0; i < node.data.length; i++) {
              const range = document.createRange()
              range.setStart(node, i)
              range.setEnd(node, i + 1)
              out.push({ ch: node.data[i], rect: range.getBoundingClientRect() })
            }
          }
          return out
        }
        const paras = [...document.querySelectorAll('.doc-page p')]
        const head = rectsOf(paras[0])
        const big = head.find((c) => c.ch === 'B')!.rect
        const small = head.find((c) => c.ch === 's')!.rect
        const second = rectsOf(paras[1])
        const huge = second.find((c) => c.ch === 'H')!.rect
        const tiny = second.find((c) => c.ch === 't')!.rect
        let smallKanaStarts = 0
        for (const para of paras.slice(2)) {
          const rects = rectsOf(para)
          rects.forEach((c, i) => {
            if (c.ch === '\u3083' && rects[i - 1].rect.bottom <= c.rect.top + 1) smallKanaStarts++
          })
        }
        return {
          smallKanaStarts,
          tops: small.top - big.top,
          smallFirstTops: tiny.top - huge.top,
          middles: (small.top + small.bottom) / 2 - (big.top + big.bottom) / 2,
          bottoms: small.bottom - big.bottom,
        }
      })
    const before = await measure()
    expect(before.smallKanaStarts).toBeGreaterThan(0)
    // on one baseline the smaller run's box sits inside the bigger one's
    expect(before.tops).toBeGreaterThan(5)

    await first.click()
    await page.keyboard.press('Shift+ArrowDown')
    await clickMenu(app, ['Format', 'Paragraph…'])
    const dialog = page.locator('.modal', { hasText: 'Paragraph' })
    await dialog.getByRole('tab', { name: 'Asian Typography' }).click()
    const pick = async (name: string) => {
      await dialog.locator('[aria-label="Text Alignment:"]').click()
      await page.getByRole('option', { name, exact: true }).click()
    }
    await pick('Top')
    await dialog.getByRole('button', { name: 'Options…' }).click()
    const typography = page.getByRole('dialog', { name: 'Typography' })
    await expect(typography.getByLabel('Standard')).toBeChecked()
    await typography.getByLabel('Strict').check()
    await typography.getByRole('button', { name: 'OK', exact: true }).click()
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog).toBeHidden()

    // Top: the ascent tops meet; Strict: no small kana starts a line any more
    await expect.poll(async () => Math.abs((await measure()).tops)).toBeLessThan(1.5)
    await expect.poll(async () => Math.abs((await measure()).smallFirstTops)).toBeLessThan(1.5)
    await expect.poll(async () => (await measure()).smallKanaStarts).toBe(0)
    await page.screenshot({ path: test.info().outputPath('text-alignment-top.png') })

    await saveFromMenu(app)
    await expect
      .poll(async () => /<w:textAlignment[^>]*>/.exec(await savedParagraph(docPath, 'Big'))?.[0], {
        timeout: 20_000,
      })
      .toBe('<w:textAlignment w:val="top"/>')
    const settings =
      (await JSZip.loadAsync(readFileSync(docPath))).file('word/settings.xml')?.async('string') ??
      ''
    expect(await settings).toContain('<w:strictFirstAndLastChars/>')

    // Center: the middles of the two boxes meet
    await first.click()
    await clickMenu(app, ['Format', 'Paragraph…'])
    await dialog.getByRole('tab', { name: 'Asian Typography' }).click()
    await pick('Center')
    // reopened with Strict in force, Typography puts the focus on Strict
    await dialog.getByRole('button', { name: 'Options…' }).click()
    await expect(typography.getByLabel('Strict')).toBeChecked()
    await expect(typography.getByLabel('Strict')).toBeFocused()
    await typography.getByRole('button', { name: 'Cancel', exact: true }).click()
    await expect(typography).toBeHidden()
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect.poll(async () => Math.abs((await measure()).middles)).toBeLessThan(1.5)

    // Auto takes the element away and the runs stand on one baseline again
    await first.click()
    await clickMenu(app, ['Format', 'Paragraph…'])
    await dialog.getByRole('tab', { name: 'Asian Typography' }).click()
    await pick('Auto')
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect.poll(async () => (await measure()).tops).toBeGreaterThan(5)
    await saveFromMenu(app)
    await expect
      .poll(async () => (await savedParagraph(docPath, 'Big')).includes('textAlignment'), {
        timeout: 20_000,
      })
      .toBe(false)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
