/**
 * An entry whose feature is not built is not shown: the menu bar, the ribbon
 * and the dialogs list none of the entries recorded in
 * `apps/docs/src/shared/unavailable-entries.ts` and as `planned` commands,
 * and what sits next to each of them still works.
 */
import { test, expect, type ElectronApplication, type Locator, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-unavailable-entries'
const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/></w:sectPr>'
const cell = (text: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${para(text)}</w:tc>`
const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
  `<w:tr>${cell('one')}${cell('3')}</w:tr><w:tr>${cell('two')}${cell('4')}</w:tr>` +
  `<w:tr>${cell('sum')}${cell('')}</w:tr></w:tbl>`
const BODY =
  '<w:p><w:pPr><w:adjustRightInd w:val="0"/></w:pPr><w:r><w:t>apple keeps its indent</w:t></w:r></w:p>' +
  para('banana') +
  TABLE +
  para('tail') +
  SECT_PR

interface Item {
  label: string
  type: string
  enabled: boolean
  checked: boolean
  submenu?: Item[]
}

const menuBar = (app: ElectronApplication): Promise<Item[]> =>
  app.evaluate(({ Menu }) => {
    type Node = {
      label: string
      type: string
      enabled: boolean
      checked: boolean
      submenu?: Node[]
    }
    const walk = (item: Electron.MenuItem): Node => ({
      label: item.label,
      type: item.type,
      enabled: item.enabled,
      checked: item.checked,
      ...(item.submenu ? { submenu: item.submenu.items.map(walk) } : {}),
    })
    return (Menu.getApplicationMenu()?.items ?? []).map(walk)
  })

const itemAt = (items: Item[], path: string[]): Item | undefined => {
  let found: Item | undefined
  let level: Item[] | undefined = items
  for (const label of path) {
    found = level?.find((item) => item.label === label)
    level = found?.submenu
  }
  return found
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

async function caretIn(target: Locator) {
  await target.click()
  await target.page().keyboard.press('End')
}

async function close(page: Page, dialog: Locator) {
  await page.keyboard.press('Escape')
  await expect(dialog).toBeHidden()
}

const absent = async (dialog: Locator, texts: string[]) => {
  for (const text of texts)
    await expect(dialog.getByText(text, { exact: true }), text).toHaveCount(0)
}

/** nothing in the open dialog is greyed out, apart from the controls named as waiting on a state */
const noneDisabled = (dialog: Locator, stateful = '') =>
  // polled: a dialog may still be reading what it shows (Paste Special reads the clipboard)
  expect
    .poll(() =>
      dialog
        .locator('button, input, select, textarea, option')
        .evaluateAll(
          (controls, stateful) =>
            controls
              .filter((el) => el.matches(':disabled') && !(stateful && el.matches(stateful)))
              .map(
                (el) =>
                  `${el.tagName.toLowerCase()} ${el.getAttribute('name') ?? el.textContent?.trim() ?? ''}`,
              ),
          stateful,
        ),
    )
    .toEqual([])

const part = async (docPath: string, name: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file(name)?.async('string') ?? ''

async function open(docPath: string) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await expect(page.locator('.doc-page p', { hasText: 'tail' })).toBeVisible({ timeout: 30_000 })
  return { launched, app: launched.app, page }
}

test('the menu bar and the ribbon list no entry that is not built', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-unavailable-')))
  const docPath = join(dir, 'menus.docx')
  writeFileSync(docPath, await docxWith(BODY))
  const { launched, app, page } = await open(docPath)
  try {
    await caretIn(page.locator('.doc-page p', { hasText: 'banana' }).first())
    await expect
      .poll(async () => itemAt(await menuBar(app), ['Insert', 'File…'])?.enabled)
      .toBe(true)
    const bar = await menuBar(app)
    const docs = ['File', 'Edit', 'View', 'Insert', 'Format', 'Tools', 'Table']
    // listed like Word's, and greyed while the document has nothing linked
    expect(itemAt(bar, ['Edit', 'Links…'])?.enabled).toBe(false)
    expect(itemAt(bar, ['Insert', 'Object…'])).toBeUndefined()

    // where an item left, its neighbours close up: no separator first, last or doubled
    const untidy: string[] = []
    const walk = (items: Item[], trail: string) => {
      items.forEach((item, i) => {
        const separator = item.type === 'separator'
        if (separator && (i === 0 || i === items.length - 1 || items[i - 1].type === 'separator'))
          untidy.push(`${trail} #${i}`)
        if (item.submenu) walk(item.submenu, `${trail} > ${item.label}`)
      })
    }
    for (const top of bar.filter((item) => docs.includes(item.label)))
      walk(top.submenu ?? [], top.label)
    expect(untidy).toEqual([])
    const labels = (path: string[]) =>
      (itemAt(bar, path)?.submenu ?? []).map((item) =>
        item.type === 'separator' ? '---' : item.label,
      )
    const edit = labels(['Edit'])
    expect(edit.slice(edit.indexOf('Find'))).toEqual([
      'Find',
      '---',
      'Links…',
      '---',
      'Select Data…',
      '---',
      'Toggle Drawing',
    ])
    const insert = labels(['Insert'])
    expect(
      insert.slice(insert.indexOf('Date and Time…'), insert.indexOf('Date and Time…') + 3),
    ).toEqual(['Date and Time…', 'File…', 'Field…'])
    expect(labels(['Tools', 'Spelling and Grammar']).slice(-5)).toEqual([
      '---',
      'Hide Spelling Errors',
      'Hide Grammar Errors',
      '---',
      'Reset Ignored Words and Grammar',
    ])

    // Hide Spelling Errors switches the squiggles
    const hideSpelling = ['Tools', 'Spelling and Grammar', 'Hide Spelling Errors']
    const hidden = async () => itemAt(await menuBar(app), hideSpelling)?.checked
    const before = await hidden()
    await clickMenu(app, hideSpelling)
    await expect.poll(hidden).toBe(!before)
    await clickMenu(app, hideSpelling)
    await expect.poll(hidden).toBe(before)

    // the ribbon's Object control keeps Text from File…, which opens the file picker
    await page
      .locator('.ribbon-tab', { hasText: /^Insert$/ })
      .first()
      .click()
    await page
      .locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: 'Object' })
      .click()
    const objectMenu = page.locator('[data-rb-panel]')
    await expect(objectMenu.getByRole('button')).toHaveText(['Text from File…'])
    await page.screenshot({ path: test.info().outputPath('object-dropdown.png') })
    await app.evaluate(({ dialog }) => {
      const state = globalThis as never as { pickerOpened: number }
      state.pickerOpened = 0
      dialog.showOpenDialog = async () => {
        state.pickerOpened += 1
        return { canceled: true, filePaths: [] }
      }
    })
    await objectMenu.getByRole('button', { name: 'Text from File…' }).click()
    await expect
      .poll(() =>
        app.evaluate(() => (globalThis as never as { pickerOpened: number }).pickerOpened),
      )
      .toBe(1)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('the dialogs list no control that is not built, and the controls next to them work', async () => {
  test.setTimeout(420_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-unavailable-')))
  const docPath = join(dir, 'dialogs.docx')
  writeFileSync(docPath, await docxWith(BODY))
  const { launched, app, page } = await open(docPath)
  try {
    const paragraph = (text: string) => page.locator('.doc-page p', { hasText: text }).first()
    const shot = (name: string) => page.screenshot({ path: test.info().outputPath(`${name}.png`) })

    // Paragraph: the right-indent box shows the paragraph's own off, Snap to grid next to it is live
    await caretIn(paragraph('apple'))
    await clickMenu(app, ['Format', 'Paragraph…'])
    const paragraphDialog = page.locator('.modal', { hasText: 'Paragraph' })
    await expect(paragraphDialog.locator('input[name="snapToGrid"]')).toBeChecked()
    await expect(paragraphDialog.locator('input[name="adjustRightInd"]')).not.toBeChecked()
    await shot('paragraph')
    await paragraphDialog.locator('input[name="snapToGrid"]').uncheck()
    // Typography, behind the Asian Typography tab: every level and spacing control is live
    await paragraphDialog.getByRole('tab', { name: 'Asian Typography' }).click()
    await paragraphDialog.getByRole('button', { name: 'Options…' }).click()
    const typography = page.getByRole('dialog', { name: 'Typography' })
    await expect(typography.getByRole('radio')).toHaveCount(6)
    await expect(typography.getByLabel('Compress punctuation and Japanese kana')).toBeEnabled()
    await expect(typography.getByLabel('Custom')).toBeEnabled()
    await noneDisabled(typography)
    await typography.getByLabel('Strict').check()
    await shot('typography')
    await typography.getByRole('button', { name: 'Cancel' }).click()
    await expect(typography).toBeHidden()
    await paragraphDialog.locator('button.btn-primary').click()
    await expect(paragraphDialog).toBeHidden()

    // Font > Advanced
    await clickMenu(app, ['Format', 'Font…'])
    const font = page.locator('.font-dialog')
    await font.getByRole('tab', { name: 'Advanced' }).click()
    await expect(font.getByLabel('Snap to grid when document grid is defined')).toBeChecked()
    await font.getByRole('checkbox', { name: 'Kerning for fonts' }).check()
    await expect(font.getByRole('spinbutton', { name: 'Kerning for fonts' })).toBeEnabled()
    await close(page, font)

    // Document > Document Grid
    await clickMenu(app, ['Format', 'Document…'])
    const documentDialog = page.locator('.doc-dialog')
    await documentDialog.getByRole('tab', { name: 'Document Grid' }).click()
    await expect(documentDialog.getByRole('button', { name: 'Drawing Grid…' })).toBeEnabled()
    await expect(documentDialog.getByRole('button', { name: 'Set Font…' })).toBeEnabled()
    await shot('document-grid')
    await close(page, documentDialog)

    // Bullets and Numbering, and its List Styles tab
    await clickMenu(app, ['Format', 'Bullets and Numbering…'])
    const bullets = page.locator('.bn-dialog')
    await expect(bullets.getByRole('button', { name: 'Customize…' })).toBeEnabled()
    await bullets.getByRole('tab', { name: 'List Styles' }).click()
    await expect(bullets.locator('.bn-styles-list [role=option]').first()).toBeVisible()
    await expect(bullets.getByRole('button', { name: 'Add' })).toBeEnabled()
    // Modify waits for a list style to be picked, Remove for one of the document's own
    await expect(bullets.locator('button[name="list-style-modify"]')).toBeDisabled()
    await bullets.locator('.bn-styles-list [role=option]', { hasText: '1 / a / i' }).click()
    await expect(bullets.locator('button[name="list-style-modify"]')).toBeEnabled()
    await noneDisabled(bullets, '[aria-label="Remove"]')
    await shot('list-styles')
    await close(page, bullets)

    // Cross-reference: headings, bookmarks, notes and caption labels
    await caretIn(paragraph('banana'))
    await clickMenu(app, ['Insert', 'Cross-reference…'])
    const crossRef = page.locator('.crossref-dialog')
    await expect(crossRef.locator('select[name="ref-type"]')).toHaveValue('heading')
    const refTypes = await crossRef
      .locator('select[name="ref-type"] option')
      .evaluateAll((options) => options.map((option) => (option as HTMLOptionElement).value))
    expect(refTypes.slice(0, 5)).toEqual(['numbered', 'heading', 'bookmark', 'footnote', 'endnote'])
    expect(refTypes.slice(5).every((value) => value.startsWith('caption:'))).toBe(true)
    // the separator waits for a paragraph number in full context
    await expect(crossRef.locator('input[name="ref-separator"]')).toBeDisabled()
    await crossRef.locator('select[name="ref-type"]').selectOption('numbered')
    await expect(crossRef.locator('select[name="ref-content"]')).toHaveValue('paraNumber')
    await crossRef.locator('select[name="ref-content"]').selectOption('paraNumberFullContext')
    await expect(crossRef.locator('input[name="ref-separator"]')).toBeEnabled()
    await crossRef.locator('select[name="ref-type"]').selectOption('bookmark')
    await crossRef.locator('input[name="ref-hyperlink"]').uncheck()
    await shot('cross-reference')
    await close(page, crossRef)

    // Footnote and Endnote
    await caretIn(paragraph('banana'))
    await clickMenu(app, ['Insert', 'Footnote…'])
    const footnote = page.locator('.footnote-dialog')
    await expect(footnote.locator('select[name="note-columns"] option')).toHaveText([
      'Match section layout',
      '1',
      '2',
      '3',
      '4',
    ])
    await expect(footnote.locator('select[name="note-restart"] option')).toHaveText([
      'Continuous',
      'Restart Each Section',
      'Restart Each Page',
    ])
    await expect(footnote.locator('select[name="note-apply-to"] option')).toHaveText([
      'Whole document',
      'This section',
    ])
    await footnote.locator('select[name="note-format"]').selectOption('upperRoman')
    await expect(footnote.locator('button[name="apply"]')).toBeEnabled()
    await shot('footnote')
    await close(page, footnote)

    // Index and Tables > Index
    await clickMenu(app, ['Insert', 'Index and Tables…'])
    const indexTables = page.locator('.index-tables-dialog')
    await indexTables.getByRole('tab', { name: 'Index', exact: true }).click()
    await expect(indexTables.locator('[data-it="columns"]')).toBeEnabled()
    await indexTables.locator('[data-it="run-in"]').check()
    await expect(indexTables.locator('[data-it="run-in"]')).toBeChecked()
    await shot('index')
    await close(page, indexTables)

    // AutoCorrect: two rules and the AutoText template list
    await clickMenu(app, ['Tools', 'AutoCorrect Options…'])
    const autoCorrect = page.locator('.autocorrect-dialog').first()
    await expect(
      autoCorrect.getByLabel('Automatically use suggestions from the spelling checker'),
    ).toBeEnabled()
    await expect(autoCorrect.getByLabel('Replace text as you type')).toBeEnabled()
    await autoCorrect.getByRole('tab', { name: 'AutoFormat As You Type' }).click()
    await expect(autoCorrect.getByLabel('Date style')).toBeEnabled()
    await expect(autoCorrect.getByLabel('Define styles based on your formatting')).toBeEnabled()
    const borderLines = autoCorrect.getByLabel('Border lines')
    const was = await borderLines.isChecked()
    await borderLines.setChecked(!was, { force: true })
    await expect(borderLines).toBeChecked({ checked: !was })
    await borderLines.setChecked(was, { force: true })
    await noneDisabled(autoCorrect)
    await shot('autocorrect')
    await autoCorrect.getByRole('tab', { name: 'AutoText' }).click()
    await expect(autoCorrect.locator('.ac-look-in select')).toBeEnabled()
    await expect(autoCorrect.locator('.ac-look-in select option')).toHaveText([
      'All active templates',
      'Normal',
    ])
    await close(page, autoCorrect)

    // Protect Document
    await clickMenu(app, ['Tools', 'Protect Document…'])
    const protect = page.locator('.protect-dialog')
    // the native inputs sit under their drawn boxes: click the label
    await protect.locator('label.protect-check', { hasText: 'Restrict editing' }).click()
    await expect(protect.getByRole('radio').first()).toBeEnabled()
    // Sections… waits for the Forms mode
    await noneDisabled(protect, '.protect-sections')
    await shot('protect')
    await protect.getByRole('button', { name: 'Cancel' }).click()
    await expect(protect).toBeHidden()

    // Compare Documents: a label box on each side, the original's grey as in Word
    await clickMenu(app, ['Tools', 'Track Changes', 'Compare Documents…'])
    const compare = page.locator('.cmp-dialog')
    await expect(compare.locator('input[name="compare-label-original"]')).toBeDisabled()
    await expect(compare.locator('.cmp-reason')).toHaveCount(0)
    await expect(compare.getByText('Label changes with')).toHaveCount(2)
    await expect(compare.locator('input[name="compare-label-revised"]')).toHaveCount(1)
    await shot('compare')
    await close(page, compare)

    // Formula
    await page.locator('.doc-page td').nth(5).click()
    // Formula is live in body text too: wait for the caret to be in the table before opening it
    await expect
      .poll(async () => itemAt(await menuBar(app), ['Table', 'Split Table'])?.enabled)
      .toBe(true)
    await clickMenu(app, ['Table', 'Formula…'])
    const formula = page.locator('.formula-dialog')
    await expect(formula.locator('input[name="formula"]')).toHaveValue('=SUM(ABOVE)')
    await expect(formula.locator('select[name="pasteBookmark"]')).toBeEnabled()
    await noneDisabled(formula)
    await shot('formula')
    await formula.locator('button.primary').click()
    await expect(formula).toBeHidden()
    await expect(page.locator('.doc-page td').nth(5)).toContainText('7')

    // Table Properties: the width stands without its checkbox
    await page.locator('.doc-page td', { hasText: 'one' }).first().click()
    await clickMenu(app, ['Table', 'Table Properties…'])
    const table = page.locator('.table-properties-dialog:not(.table-cell-options)')
    await table.getByRole('tab', { name: 'Column' }).click()
    await expect(table.locator('.table-properties-width input[type=checkbox]')).toHaveCount(0)
    await expect(table.locator('.table-properties-width')).toContainText('Preferred width')
    await table.locator('.table-properties-width select').selectOption('pct')
    await expect(table.locator('.table-properties-pct input')).toBeEnabled()
    await shot('table-properties')
    await close(page, table)

    // Paste Special
    await app.evaluate(({ clipboard }) =>
      clipboard.write({ text: 'pasted words', html: '<b>pasted words</b>' }),
    )
    await caretIn(paragraph('tail'))
    await clickMenu(app, ['Edit', 'Paste Special…'])
    const paste = page.locator('.paste-special-dialog')
    // the formats are listed once the clipboard has been read; OK waits for one
    const pasteAs = paste.locator('.paste-special-as select')
    await expect(pasteAs).toHaveValue('html')
    // Paste link waits for a clipboard copied from a saved document; Display as icon is not offered
    await expect(paste.getByRole('radio')).toHaveCount(2)
    await absent(paste, ['Display as icon'])
    await noneDisabled(paste, 'input[name="paste-special-mode"]')
    await pasteAs.selectOption('text')
    await shot('paste-special')
    await paste.locator('button.primary').click()
    await expect(paste).toBeHidden()
    await expect(paragraph('tail')).toContainText('pasted words')

    // saved: the paragraph keeps the value of the box that was left alone, next to the change
    await saveFromMenu(app)
    await expect
      .poll(async () => (await part(docPath, 'word/document.xml')).includes('pasted words'), {
        timeout: 20_000,
      })
      .toBe(true)
    const xml = await part(docPath, 'word/document.xml')
    const apple = xml.split('<w:p>').find((chunk) => chunk.includes('apple keeps')) ?? ''
    expect(apple).toMatch(/<w:adjustRightInd w:val="0"\/>/)
    expect(apple).toMatch(/<w:snapToGrid w:val="0"\/>/)
    expect(xml).toMatch(/SUM\(ABOVE\)/)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
