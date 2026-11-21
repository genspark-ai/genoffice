/**
 * Entry points greyed or live in the situations Word greys or enables them:
 * the menu bar by the caret, a selection, cells, a picture and a read-only
 * document; Table > Formula… in body text; Split Table from the first row;
 * the right-click menu on selected cells and on a picture in a line of text;
 * Picture Format > Align.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'
import {
  caretAtEnd,
  clickMenu,
  open,
  ribbonTab,
  savedBody,
  selectWord,
  SECT_PR,
} from './docs-word-acceptance-kit'

const VIDEO_DIR = 'docs-entry-states-like-word'

const cellXml = (text: string) => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`
const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>' +
  ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map((side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`)
    .join('') +
  '</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
  `<w:tr>${cellXml('A1')}${cellXml('B1')}${cellXml('C1')}</w:tr>` +
  `<w:tr>${cellXml('A2')}${cellXml('B2')}${cellXml('C2')}</w:tr></w:tbl>`

/** a paragraph holding one picture in the line of text */
const INLINE_PICTURE =
  '<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
  '<wp:extent cx="1143000" cy="762000"/><wp:docPr id="7" name="Picture 7"/><wp:cNvGraphicFramePr/>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:nvPicPr><pic:cNvPr id="7" name="Picture 7"/><pic:cNvPicPr/></pic:nvPicPr>' +
  '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1143000" cy="762000"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'

const BODY =
  para('Plain paragraph here') +
  TABLE +
  para('Between') +
  INLINE_PICTURE +
  '<w:p><w:bookmarkStart w:id="1" w:name="Rate"/><w:r><w:t>Rate 20</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>' +
  '<w:p><w:r><w:t xml:space="preserve">Total: </w:t></w:r></w:p>' +
  para('Tail')

const enabledOf = (app: ElectronApplication, paths: string[][]) =>
  app.evaluate(
    ({ Menu }, paths) =>
      paths.map((path) => {
        let items = Menu.getApplicationMenu()?.items
        let item: Electron.MenuItem | undefined
        for (const label of path) {
          item = items?.find((candidate) => candidate.label === label)
          items = item?.submenu?.items
        }
        return item ? item.enabled : null
      }),
    paths,
  )

const selectionKind = (page: Page) =>
  page.evaluate(() => {
    const { selection } = (
      window as never as {
        __aidocs: {
          editor: { state: { selection: { empty: boolean; toJSON(): { type: string } } } }
        }
      }
    ).__aidocs.editor.state
    return selection.empty ? 'caret' : selection.toJSON().type
  })

const cell = (page: Page, index: number) => page.locator('.doc-page .doc-table td').nth(index)
const plain = (page: Page) =>
  page.locator('.doc-page p', { hasText: 'Plain paragraph here' }).first()
const picture = (page: Page) => page.locator('.doc-page[contenteditable="true"] img').first()

const ASIAN = ['Format', 'Asian Layout']
/** each item with whether Word has it live: caret in body, text selected, caret in a cell, cells selected, picture selected */
const STATES: Array<[string[], string]> = [
  [['Edit', 'Select Data…'], '00000'],
  [['View', 'Footnotes'], '00000'],
  [['Insert', 'AutoText', 'New…'], '01011'],
  [[...ASIAN, 'Fit Text…'], '01010'],
  [['Format', 'Drop Cap…'], '11000'],
  [['Insert', 'Date and Time…'], '11100'],
  [['Insert', 'Field…'], '11100'],
  [['Insert', 'Advanced Symbol…'], '11100'],
  [[...ASIAN, 'Phonetic Guide…'], '11100'],
  [[...ASIAN, 'Enclose Characters…'], '11100'],
  [[...ASIAN, 'Horizontal in Vertical…'], '11100'],
  [[...ASIAN, 'Combine Characters…'], '11100'],
  [[...ASIAN, 'Two Lines in One…'], '11100'],
  [['Tools', 'Thesaurus…'], '11100'],
  [['Tools', 'Translate', 'Translate Selection'], '11100'],
  [['Table', 'Formula…'], '11110'],
  [['Table', 'Split Table'], '00110'],
  [['Edit', 'Find', 'Replace…'], '11111'],
]

test('the menu bar greys what Word greys at a caret, on a selection, in a table, on cells and on a picture', async () => {
  test.setTimeout(180_000)
  const session = await open(VIDEO_DIR, BODY, 'states.docx')
  const { app, page } = session
  try {
    await expect(plain(page)).toBeVisible({ timeout: 30_000 })
    const situations: Array<[string, string, () => Promise<void>]> = [
      ['caret in body', 'caret', () => plain(page).click()],
      [
        'text selected',
        'text',
        async () => {
          await plain(page).click()
          await plain(page).dblclick({ position: { x: 20, y: 8 } })
        },
      ],
      ['caret in a cell', 'caret', () => cell(page, 0).click()],
      [
        'cells selected',
        'cell',
        async () => {
          await cell(page, 0).click()
          await cell(page, 4).click({ modifiers: ['Shift'] })
        },
      ],
      ['picture selected', 'node', () => picture(page).click()],
    ]
    for (const [index, [name, kind, act]] of situations.entries()) {
      await act()
      await expect.poll(() => selectionKind(page), { message: name }).toBe(kind)
      await expect
        .poll(
          () =>
            enabledOf(
              app,
              STATES.map(([path]) => path),
            ),
          { message: name },
        )
        .toEqual(STATES.map(([, live]) => live[index] === '1'))
    }

    // the Home tab's Asian Layout items follow the same rules as the Format menu's
    await plain(page).click()
    await ribbonTab(page, 'Home')
    const asianButton = page.locator('[data-ribbon-body] button.rb-icon[aria-label="Asian Layout"]')
    await asianButton.click()
    const asianItems = page.locator('.layout-menu button')
    await expect(asianItems.filter({ hasText: 'Fit Text…' })).toBeDisabled()
    await expect(asianItems.filter({ hasText: 'Two Lines in One' })).toBeEnabled()
    await page.keyboard.press('Escape')

    // a footnote makes View > Footnotes live, a chart Edit > Select Data
    await caretAtEnd(page, 'Tail')
    await clickMenu(app, ['Insert', 'Chart', 'Bar'])
    const chart = page.locator('.modal-chart')
    await chart.locator('.modal-actions .btn-primary').click()
    await expect(chart).toBeHidden()
    await expect(page.locator('.doc-page .doc-chart-canvas')).toHaveCount(1)
    await plain(page).click()
    await expect.poll(() => enabledOf(app, [['Edit', 'Select Data…']])).toEqual([true])
    // from the text it selects the chart, whose data then shows
    await clickMenu(app, ['Edit', 'Select Data…'])
    await expect(page.locator('.doc-page .doc-chart-data')).toBeVisible()
    await caretAtEnd(page, 'Plain paragraph here')
    await page.evaluate(() =>
      (
        window as never as {
          __aidocs: { editor: { commands: { focus(): boolean } } }
        }
      ).__aidocs.editor.commands.focus(),
    )
    await clickMenu(app, ['Insert', 'Footnote…'])
    const footnote = page.locator('.footnote-dialog')
    await footnote.locator('button.primary').click()
    await page.locator('.modal .prompt-textarea').fill('a note')
    await page.keyboard.press('ControlOrMeta+Enter')
    await expect(page.locator('.doc-page sup.doc-note-ref')).toHaveCount(1)
    await expect.poll(() => enabledOf(app, [['View', 'Footnotes']])).toEqual([true])
  } finally {
    await session.close()
  }
})

test('Table > Formula… puts a formula field in body text, refuses cells there, and the file keeps it', async () => {
  test.setTimeout(180_000)
  const session = await open(VIDEO_DIR, BODY, 'formula.docx')
  const { app, page, docPath } = session
  try {
    await expect(plain(page)).toBeVisible({ timeout: 30_000 })
    await caretAtEnd(page, 'Total: ')
    await clickMenu(app, ['Table', 'Formula…'])
    const dialog = page.locator('.formula-dialog')
    const formula = dialog.locator('input[name="formula"]')
    // outside a table Word proposes no sum
    await expect(formula).toHaveValue('=')
    // a direction and a cell have nothing to read here
    for (const refused of ['=SUM(ABOVE)', '=A1*2']) {
      await formula.fill(refused)
      await dialog.locator('button.primary').click()
      await expect(dialog.getByRole('alert')).toContainText('Outside a table')
    }
    await page.screenshot({ path: test.info().outputPath('formula-body-refused.png') })
    await formula.fill('=(2+3)*Rate')
    await dialog.locator('input[name="format"]').fill('0.00')
    await page.screenshot({ path: test.info().outputPath('formula-body-dialog.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    const total = page.locator('.doc-page p', { hasText: 'Total: ' }).first()
    await expect(total).toHaveText('Total: 100.00')
    await page.screenshot({ path: test.info().outputPath('formula-body-page.png') })

    // with the caret in the field the dialog opens on its formula
    await caretAtEnd(page, 'Total: ')
    await clickMenu(app, ['Table', 'Formula…'])
    await expect(formula).toHaveValue('=(2+3)*Rate')
    await expect(dialog.locator('input[name="format"]')).toHaveValue('0.00')
    await page.keyboard.press('Escape')
    await expect(dialog).toBeHidden()

    // the bookmark changes: the result waits for Update Field, as in Word
    await caretAtEnd(page, 'Rate 20')
    await page.keyboard.type('0')
    await expect(total).toHaveText('Total: 100.00')
    await page.keyboard.press('F9')
    await expect(total).toHaveText('Total: 1000.00')

    await saveFromMenu(app)
    await expect
      .poll(() => savedBody(docPath), { timeout: 30_000 })
      .toMatch(/=\(2\+3\)\*Rate \\# &quot;0\.00&quot;|=\(2\+3\)\*Rate \\# "0\.00"/)
    const saved = await savedBody(docPath)
    expect(saved).toContain('1000.00')
    expect(saved).toMatch(/w:fldSimple|w:fldChar/)
    // the computed result is what the file carries: Word is not asked to update the field
    expect(saved).toMatch(
      /Rate \\# (?:&quot;|")0\.00(?:&quot;|") ?<\/w:instrText>[\s\S]*?<w:t[^>]*>1000\.00<\/w:t>/,
    )
    expect(saved).not.toContain('w:dirty')
  } finally {
    await session.close()
  }
})

test('Split Table from the first row puts a paragraph above the table, from the ribbon and the menu alike', async () => {
  test.setTimeout(180_000)
  const session = await open(VIDEO_DIR, TABLE + para('Tail'), 'split.docx')
  const { app, page, docPath } = session
  try {
    await expect(cell(page, 0)).toBeVisible({ timeout: 30_000 })
    await cell(page, 1).click()
    await ribbonTab(page, 'Table Layout')
    const split = page.locator('[data-ribbon-body]').getByRole('button', { name: 'Split Table' })
    await expect(split).toBeEnabled()
    await expect.poll(() => enabledOf(app, [['Table', 'Split Table']])).toEqual([true])
    await split.click()
    // one table still, the caret in a new paragraph above it
    await expect(page.locator('.doc-page .doc-table')).toHaveCount(1)
    await page.keyboard.type('Above the table')
    await expect(page.locator('.doc-page .doc-table td')).toHaveCount(6)
    await saveFromMenu(app)
    await expect
      .poll(() => savedBody(docPath), { timeout: 30_000 })
      .toMatch(/Above the table<\/w:t><\/w:r><\/w:p>\s*<w:tbl>/)

    // a lower row still cuts the table in two
    await cell(page, 3).click()
    await clickMenu(app, ['Table', 'Split Table'])
    await expect(page.locator('.doc-page .doc-table')).toHaveCount(2)
  } finally {
    await session.close()
  }
})

test('a right-click inside selected cells keeps them, so Merge Cells, Cut and Copy are live', async () => {
  test.setTimeout(180_000)
  const session = await open(VIDEO_DIR, BODY, 'cells.docx')
  const { app, page, docPath } = session
  try {
    await expect(plain(page)).toBeVisible({ timeout: 30_000 })
    await cell(page, 0).click()
    await cell(page, 2).click({ modifiers: ['Shift'] })
    await expect.poll(() => selectionKind(page)).toBe('cell')
    await cell(page, 1).click({ button: 'right' })
    const menu = page.locator('.ctx-menu')
    await menu.waitFor({ timeout: 15_000 })
    expect(await selectionKind(page)).toBe('cell')
    const item = (label: RegExp) =>
      menu.locator('.ctx-item', { has: page.locator('.ctx-label', { hasText: label }) }).first()
    for (const label of [/^Cut$/, /^Copy$/, /^Merge Cells$/])
      await expect(item(label)).toBeEnabled()
    await page.screenshot({ path: test.info().outputPath('cells-context-menu.png') })
    await item(/^Merge Cells$/).click()
    await expect(page.locator('.doc-page .doc-table tr').first().locator('td')).toHaveCount(1)
    await saveFromMenu(app)
    await expect
      .poll(() => savedBody(docPath), { timeout: 30_000 })
      .toContain('<w:gridSpan w:val="3"/>')

    // a right-click outside the selected cells still moves the caret there
    await cell(page, 1).click()
    await cell(page, 2).click({ modifiers: ['Shift'] })
    await expect.poll(() => selectionKind(page)).toBe('cell')
    await plain(page).click({ button: 'right', position: { x: 200, y: 8 } })
    await menu.waitFor({ timeout: 15_000 })
    expect(await selectionKind(page)).not.toBe('cell')
    await expect(item(/^Merge Cells$/)).toHaveCount(0)
  } finally {
    await session.close()
  }
})

test('a picture in a line of text: the stacking order is greyed in its menu, Align opens with its items greyed', async () => {
  test.setTimeout(180_000)
  const session = await open(VIDEO_DIR, BODY, 'picture.docx')
  const { page } = session
  try {
    await expect(picture(page)).toBeVisible({ timeout: 30_000 })
    const rightClick = async () => {
      const box = (await picture(page).boundingBox())!
      await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2, { button: 'right' })
      await page.locator('.ctx-menu').waitFor({ timeout: 15_000 })
    }
    const menu = page.locator('.ctx-menu')
    const top = (label: RegExp) => menu.locator('.ctx-item-wrap > .ctx-item', { hasText: label })
    await rightClick()
    await expect(top(/^Bring to Front/)).toBeDisabled()
    await expect(top(/^Send to Back/)).toBeDisabled()
    await expect(top(/^Wrap Text/)).toBeEnabled()
    await page.screenshot({ path: test.info().outputPath('inline-picture-menu.png') })
    await page.keyboard.press('Escape')

    await ribbonTab(page, 'Picture Format')
    const align = page.locator('[data-ribbon-body] button.rb-align-objects')
    await expect(align).toBeEnabled()
    await align.click()
    const items = page.locator('[data-rb-panel].layout-menu button')
    await expect(items).toHaveCount(11)
    for (const label of ['Align Left', 'Align Middle', 'Distribute Horizontally'])
      await expect(items.filter({ hasText: label })).toBeDisabled()
    await page.screenshot({ path: test.info().outputPath('inline-picture-align.png') })
    await page.keyboard.press('Escape')

    // once the picture floats it has a place in the stack
    await rightClick()
    await top(/^Wrap Text/).hover()
    await page.locator('.ctx-submenu .ctx-item', { hasText: 'In Front of Text' }).click()
    await rightClick()
    await expect(top(/^Bring to Front/)).toBeEnabled()
    await expect(top(/^Send to Back/)).toBeEnabled()
  } finally {
    await session.close()
  }
})

/** `body` as a document whose settings enforce read-only protection */
async function readOnlyDocx(body: string): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await docxWith(body + SECT_PR))
  zip.file(
    'word/settings.xml',
    '<?xml version="1.0" encoding="UTF-8"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:documentProtection w:edit="readOnly" w:enforcement="1"/></w:settings>',
  )
  const types = await zip.file('[Content_Types].xml')!.async('string')
  zip.file(
    '[Content_Types].xml',
    types.replace(
      '</Types>',
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>',
    ),
  )
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
  zip.file(
    'word/_rels/document.xml.rels',
    rels.replace(
      '</Relationships>',
      '<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>',
    ),
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

test('a read-only document keeps Replace… and Thesaurus… live; neither changes it', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-entry-states-')))
  const docPath = join(dir, 'protected.docx')
  writeFileSync(docPath, await readOnlyDocx(BODY))
  const before = readFileSync(docPath)
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    await expect(plain(page)).toBeVisible({ timeout: 30_000 })
    // a protected document does not hold a selection for long, so it is made again on each try
    const paths = [
      ['Edit', 'Find', 'Replace…'],
      ['Tools', 'Thesaurus…'],
      ['Tools', 'Translate', 'Translate Selection'],
      ['Table', 'Formula…'],
      ['Insert', 'Date and Time…'],
    ]
    await expect(async () => {
      await selectWord(page, 'Plain')
      expect(await selectionKind(page)).toBe('text')
      expect(await enabledOf(app, paths)).toEqual([true, true, false, false, false])
    }).toPass({ timeout: 30_000 })

    // Replace… opens the Find panel, which has no replacing to offer here
    await clickMenu(app, ['Edit', 'Find', 'Replace…'])
    const find = page.locator('.find-panel')
    await expect(find).toBeVisible()
    await expect(find.getByRole('button', { name: /Replace/ })).toHaveCount(0)
    await page.keyboard.press('Escape')

    // Thesaurus… asks about the selected word in the AI panel
    await expect(async () => {
      await selectWord(page, 'Plain')
      await clickMenu(app, ['Tools', 'Thesaurus…'])
      await expect(page.locator('.ai-dock')).not.toHaveClass(/collapsed/, { timeout: 2_000 })
      await expect(page.locator('.ai-dock')).toContainText('synonyms', { timeout: 2_000 })
    }).toPass({ timeout: 30_000 })
    await expect(plain(page)).toHaveText('Plain paragraph here')
    expect(readFileSync(docPath).equals(before)).toBe(true)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
