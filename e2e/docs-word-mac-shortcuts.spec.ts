/**
 * The Word for Mac menu shortcuts Docs lacked (New from Template, Paste
 * Special, Find Next / Previous, Toggle Drawing, Focus, Ribbon, Equation,
 * Format Object, Selection Pane, Thesaurus, Delete Rows) show on their menu
 * items and run from the keyboard, a chord the user gave to another command
 * still wins, and the table-cell and picture right-click menus carry Word's
 * groups.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import { mkdtempSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-word-mac-shortcuts'
const IS_MAC = process.platform === 'darwin'
const KITCHEN_SINK = resolve(__dirname, '../fixtures/generated/kitchen-sink.docx')

const cell = (text: string) => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`
const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
  `<w:tr>${cell('A1')}${cell('B1')}</w:tr><w:tr>${cell('A2')}${cell('B2')}</w:tr></w:tbl>`

const BODY =
  para('alpha one') + para('alpha two') + para('alpha three') + '<w:p/>' + TABLE + para('After')

/** [menu path, the chord Word for Mac shows on that item] */
const WORD_CHORDS: [string[], string][] = [
  [['File', 'New from Template…'], 'Shift+Cmd+P'],
  [['Edit', 'Paste Special…'], 'Ctrl+Cmd+V'],
  [['Edit', 'Find', 'Find Next'], 'Cmd+G'],
  [['Edit', 'Find', 'Find Previous'], 'Shift+Cmd+G'],
  [['Edit', 'Toggle Drawing'], 'Ctrl+Cmd+Z'],
  [['View', 'Focus'], 'Ctrl+Shift+Cmd+F'],
  [['View', 'Ribbon'], 'Alt+Cmd+R'],
  [['Insert', 'Equation…'], 'Ctrl+='],
  [['Format', 'Format Object…'], 'Shift+Cmd+1'],
  [['Format', 'Selection Pane…'], 'Alt+Cmd+U'],
  [['Tools', 'Thesaurus…'], 'Ctrl+Alt+Cmd+R'],
  [['Table', 'Delete', 'Rows'], 'Ctrl+Cmd+X'],
]

const menuItem = (app: ElectronApplication, path: string[]) =>
  app.evaluate(({ Menu }, path) => {
    let items = Menu.getApplicationMenu()?.items
    let item: Electron.MenuItem | undefined
    for (const label of path) {
      item = items?.find((candidate) => candidate.label === label)
      items = item?.submenu?.items
    }
    return item
      ? { accelerator: String(item.accelerator ?? ''), checked: item.checked }
      : { accelerator: 'missing', checked: false }
  }, path)

/** puts the caret in the paragraph or cell that reads `text`, `offset` characters in */
const caretIn = (page: Page, text: string, offset = 0) =>
  page.evaluate(
    ({ text, offset }) => {
      interface Node {
        isTextblock: boolean
        textContent: string
      }
      const { editor } = (
        window as never as {
          __aidocs: {
            editor: {
              state: { doc: { descendants(visit: (node: Node, pos: number) => void): void } }
              chain(): { focus(): { setTextSelection(pos: number): { run(): boolean } } }
            }
          }
        }
      ).__aidocs
      let at = -1
      editor.state.doc.descendants((node, pos) => {
        if (at < 0 && node.isTextblock && node.textContent === text) at = pos + 1
      })
      if (at < 0) throw new Error(`no paragraph reads "${text}"`)
      editor
        .chain()
        .focus()
        .setTextSelection(at + offset)
        .run()
    },
    { text, offset },
  )

async function openDoc(docPath: string, userDataDir?: string) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
    userDataDir,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await page.waitForFunction(
    () => Boolean((window as unknown as { __aidocs?: { editor?: unknown } }).__aidocs?.editor),
    undefined,
    { timeout: 30_000 },
  )
  return { launched, page }
}

function tempDoc(body: Buffer): { dir: string; docPath: string } {
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-word-mac-shortcuts-')))
  const docPath = join(dir, 'shortcuts.docx')
  writeFileSync(docPath, body)
  return { dir, docPath }
}

test('each Word for Mac menu shortcut shows on its item and runs from the keyboard', async () => {
  test.skip(!IS_MAC, 'these chords are Word for Mac’s; other platforms keep their own')
  test.setTimeout(240_000)
  const { dir, docPath } = tempDoc(await docxWith(BODY))
  const { launched, page } = await openDoc(docPath)
  const { app } = launched
  try {
    await expect(page.locator('.doc-page p', { hasText: 'alpha three' })).toBeVisible({
      timeout: 30_000,
    })
    for (const [path, chord] of WORD_CHORDS)
      await expect
        .poll(async () => (await menuItem(app, path)).accelerator, { message: path.join(' > ') })
        .toBe(chord)
    // Word for Mac gives Word Count no chord
    expect((await menuItem(app, ['Tools', 'Word Count…'])).accelerator).toBe('')

    const status = page.locator('.status-msg')
    await caretIn(page, 'alpha one')

    await page.keyboard.press('Shift+Meta+KeyP')
    const gallery = page.locator('.modal', { hasText: 'New from Template' })
    await expect(gallery).toBeVisible()
    await page.keyboard.press('Escape')
    await expect(gallery).toBeHidden()

    await caretIn(page, 'alpha one')
    await page.keyboard.press('Control+Meta+KeyV')
    const pasteSpecial = page.locator('.paste-special-dialog')
    await expect(pasteSpecial).toBeVisible()
    await page.keyboard.press('Escape')
    await expect(pasteSpecial).toBeHidden()

    // Find Next / Find Previous step through the matches; Shift+Cmd+G is no longer Word Count
    await page.keyboard.press('Meta+KeyF')
    const find = page.locator('.find-panel')
    await find.locator('input').first().fill('alpha')
    const count = find.locator('.find-count')
    await expect(count).toHaveText('1/3')
    await page.keyboard.press('Meta+KeyG')
    await expect(count).toHaveText('2/3')
    await page.keyboard.press('Meta+KeyG')
    await expect(count).toHaveText('3/3')
    await page.keyboard.press('Shift+Meta+KeyG')
    await expect(count).toHaveText('2/3')
    await expect(page.locator('.word-count-dialog')).toHaveCount(0)
    await page.keyboard.press('Escape')
    await expect(find).toBeHidden()

    await caretIn(page, 'alpha one')
    const ink = page.locator('.ink-overlay.ink-active')
    await page.keyboard.press('Control+Meta+KeyZ')
    await expect(ink).toHaveCount(1)
    await page.keyboard.press('Control+Meta+KeyZ')
    await expect(ink).toHaveCount(0)

    // Focus: its own view, and not the Find panel that any Cmd+F used to open
    const focus = page.locator('.app.focus-mode')
    await page.keyboard.press('Control+Shift+Meta+KeyF')
    await expect(focus).toHaveCount(1)
    await expect(find).toBeHidden()
    await page.keyboard.press('Control+Shift+Meta+KeyF')
    await expect(focus).toHaveCount(0)

    const ribbonShown = async () => (await menuItem(app, ['View', 'Ribbon'])).checked
    await caretIn(page, 'alpha one')
    await expect.poll(ribbonShown).toBe(true)
    await page.keyboard.press('Alt+Meta+KeyR')
    await expect.poll(ribbonShown).toBe(false)
    await page.keyboard.press('Alt+Meta+KeyR')
    await expect.poll(ribbonShown).toBe(true)

    await caretIn(page, 'alpha one')
    await page.keyboard.press('Control+Equal')
    const equation = page.locator('.equation-modal')
    await expect(equation).toBeVisible()
    await page.keyboard.press('Escape')
    await expect(equation).toBeHidden()

    // nothing is selected that Format Object could format: it says so
    await caretIn(page, 'alpha one')
    await page.keyboard.press('Shift+Meta+Digit1')
    await expect(status).toContainText('Select a picture, text box or shape first')

    const selectionPane = page.locator('.selection-pane')
    await page.keyboard.press('Alt+Meta+KeyU')
    await expect(selectionPane).toBeVisible()
    await page.keyboard.press('Alt+Meta+KeyU')
    await expect(selectionPane).toBeHidden()

    // the caret is in no word: Thesaurus asks for one
    await caretIn(page, '')
    await page.keyboard.press('Control+Alt+Meta+KeyR')
    await expect(status).toContainText('Select a word or a short phrase first')

    // Delete Rows is greyed out outside a table, and the key does nothing there
    const rows = page.locator('.doc-page table tr')
    await expect(rows).toHaveCount(2)
    await caretIn(page, 'alpha one')
    await page.keyboard.press('Control+Meta+KeyX')
    await expect(page.locator('.doc-page p', { hasText: 'alpha one' })).toBeVisible()
    await expect(rows).toHaveCount(2)
    await caretIn(page, 'A1')
    await expect
      .poll(async () =>
        app.evaluate(({ Menu }) => {
          const table = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Table')
          const del = table?.submenu?.items.find((item) => item.label === 'Delete')
          return del?.submenu?.items.find((item) => item.label === 'Rows')?.enabled
        }),
      )
      .toBe(true)
    await page.keyboard.press('Control+Meta+KeyX')
    await expect(rows).toHaveCount(1)
    await expect(page.locator('.doc-page table')).toContainText('A2')
    await expect(page.locator('.doc-page table')).not.toContainText('A1')
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a chord the user gave to another command wins over the Word default', async () => {
  test.skip(!IS_MAC, 'these chords are Word for Mac’s; other platforms keep their own')
  test.setTimeout(180_000)
  const { dir, docPath } = tempDoc(await docxWith(BODY))
  const userDataDir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-')))
  // as Customize Keyboard stores it: Word Count takes Shift+Cmd+G, which Find Previous gives up
  writeFileSync(
    join(userDataDir, 'docs-keyboard.json'),
    JSON.stringify({ 'word-count': ['Shift+Cmd+G'], 'edit-find-previous': [] }),
  )
  const { launched, page } = await openDoc(docPath, userDataDir)
  const { app } = launched
  try {
    await expect(page.locator('.doc-page p', { hasText: 'alpha three' })).toBeVisible({
      timeout: 30_000,
    })
    await expect
      .poll(async () => (await menuItem(app, ['Tools', 'Word Count…'])).accelerator)
      .toBe('Shift+Cmd+G')
    expect((await menuItem(app, ['Edit', 'Find', 'Find Previous'])).accelerator).toBe('')
    expect((await menuItem(app, ['Edit', 'Find', 'Find Next'])).accelerator).toBe('Cmd+G')

    await caretIn(page, 'alpha one')
    await page.keyboard.press('Meta+KeyF')
    const find = page.locator('.find-panel')
    await find.locator('input').first().fill('alpha')
    const count = find.locator('.find-count')
    await expect(count).toHaveText('1/3')
    await page.keyboard.press('Meta+KeyG')
    await expect(count).toHaveText('2/3')
    // from the Find box the chord is no longer Find Previous
    await page.keyboard.press('Shift+Meta+KeyG')
    await page.keyboard.press('Meta+KeyG')
    await expect(count).toHaveText('3/3')
    await caretIn(page, 'alpha one')
    await expect
      .poll(() => page.evaluate(() => !!document.activeElement?.closest('.ProseMirror')))
      .toBe(true)
    await page.keyboard.press('Shift+Meta+KeyG')
    await expect(page.locator('.word-count-dialog')).toBeVisible()
    await expect(count).toHaveText('3/3')
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
    rmSync(userDataDir, { recursive: true, force: true })
  }
})

/** the rows of the open right-click menu, a separator as '---' */
const menuOutline = (page: Page) =>
  page
    .locator('.ctx-menu')
    .evaluate((menu) =>
      [...menu.children].map((row) =>
        row.classList.contains('ctx-sep')
          ? '---'
          : ((row.matches('.ctx-item') ? row : row.querySelector(':scope > .ctx-item'))
              ?.querySelector('.ctx-label')
              ?.textContent?.replace(/^✓ /, '') ?? ''),
      ),
    )

test('the table-cell and picture right-click menus carry Word’s groups', async () => {
  test.setTimeout(120_000)
  const { launched, page } = await openDoc(KITCHEN_SINK)
  try {
    const menu = page.locator('.ctx-menu')
    const cellA1 = page.locator('.doc-page table td', { hasText: 'A1' }).first()
    await cellA1.scrollIntoViewIfNeeded()
    await cellA1.click()
    await cellA1.click({ button: 'right' })
    await menu.waitFor({ timeout: 15_000 })
    await page.screenshot({ path: test.info().outputPath('table-cell-menu.png') })
    // Word's cell menu and nothing else, but for the item that withholds a selection from AI
    expect(await menuOutline(page)).toEqual([
      'Cut',
      'Copy',
      'Paste',
      '---',
      'Reply To Comment',
      'Delete Comment',
      'Resolve Comment',
      '---',
      'Insert',
      'Delete Cells…',
      'Split Cells…',
      '---',
      'Borders and Shading…',
      'Text Direction…',
      'Cell Alignment',
      'Table Properties…',
      '---',
      'Hide the selection from AI',
      '---',
      'New Comment',
    ])
    // one item, as in Word: no submenu of directions
    await expect(
      menu.locator('.ctx-item:has(.ctx-label:text-is("Text Direction…")) .ctx-arrow'),
    ).toHaveCount(0)
    await menu.locator('.ctx-item', { hasText: 'Delete Cells…' }).click()
    const deleteCells = page.locator('.modal', { hasText: 'Delete Cells' })
    await expect(deleteCells).toBeVisible()
    await page.keyboard.press('Escape')
    await expect(deleteCells).toBeHidden()
    await cellA1.click({ button: 'right' })
    await menu.locator('.ctx-item:has(.ctx-label:text-is("Text Direction…"))').click()
    const textDirection = page.locator('.modal', { hasText: 'Text Direction' })
    await expect(textDirection).toBeVisible()
    await page.screenshot({ path: test.info().outputPath('text-direction-dialog.png') })
    await textDirection.locator('button', { hasText: 'Cancel' }).click()
    await expect(textDirection).toBeHidden()

    const img = page.locator('.doc-page[contenteditable="true"] img').first()
    await img.scrollIntoViewIfNeeded()
    const box = (await img.boundingBox())!
    await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2, { button: 'right' })
    await menu.waitFor({ timeout: 15_000 })
    await page.screenshot({ path: test.info().outputPath('picture-menu.png') })
    expect(await menuOutline(page)).toEqual([
      'Cut',
      'Copy',
      'Paste',
      '---',
      'Save as Picture…',
      'Change Picture',
      '---',
      'Group',
      'Bring to Front',
      'Send to Back',
      '---',
      'Hyperlink…',
      '---',
      'Insert Caption…',
      'Wrap Text',
      '---',
      'View Alt Text…',
      'Crop',
      'Auto Crop',
      'Size and Position…',
      'Format Picture…',
      '---',
      'View Image',
    ])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
  }
})
