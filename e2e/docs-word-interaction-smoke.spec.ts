import { test, expect, type ElectronApplication, type Locator, type Page } from '@playwright/test'
import { mkdtempSync, readdirSync, readFileSync, realpathSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, type LaunchedApp } from './helpers'

/**
 * One document session that chains the Word-aligned interactions (style
 * gallery, selection bar, format painter, list keys, font boxes, contextual
 * table tabs, ruler and units, Layout galleries, paste modes, zoom and status
 * bar, navigation pane, hyperlink menu, save round trip) so that the pieces
 * are exercised against each other's state, not only in isolation.
 */

const POLL = { timeout: 15_000, intervals: [250, 500, 1000] }

const WORD_GALLERY = [
  'Normal',
  'No Spacing',
  'Heading 1',
  'Heading 2',
  'Title',
  'Subtitle',
  'Subtle Emphasis',
  'Emphasis',
  'Intense Emphasis',
  'Strong',
  'Quote',
  'Intense Quote',
  'Subtle Reference',
  'Intense Reference',
  'Book Title',
  'List Paragraph',
]

const LONG = Array.from({ length: 60 }, (_, i) => `gamma${i}`).join(' ')

interface AidocsWindow {
  __aidocs?: {
    editor?: any
    save?: () => Promise<unknown>
    openPath?: (p: string) => void
  }
}

interface Sel {
  from: number
  to: number
  docSize: number
  text: string
}

interface Run {
  text: string
  bold: boolean
  sizeHalfPoints: number | null
}

async function selection(editor: Page): Promise<Sel> {
  return editor.evaluate(() => {
    const ed = (window as unknown as AidocsWindow).__aidocs!.editor!
    const s = ed.state.selection
    return {
      from: s.from,
      to: s.to,
      docSize: ed.state.doc.content.size,
      text: ed.state.doc.textBetween(s.from, s.to, '\n'),
    }
  })
}

async function runs(editor: Page): Promise<Run[]> {
  return editor.evaluate(() => {
    const ed = (window as unknown as AidocsWindow).__aidocs!.editor!
    const out: Run[] = []
    ed.state.doc.descendants((n: any) => {
      if (!n.isText) return true
      const ts = n.marks.find((m: any) => m.type.name === 'docTextStyle')
      out.push({
        text: n.text ?? '',
        bold: n.marks.some((m: any) => m.type.name === 'bold'),
        sizeHalfPoints: (ts?.attrs.sizeHalfPoints as number | null) ?? null,
      })
      return true
    })
    return out
  })
}

const boldText = async (editor: Page) =>
  (await runs(editor)).filter((r) => r.bold).map((r) => r.text)

/** block types + text in document order */
async function blocks(editor: Page): Promise<string[]> {
  return editor.evaluate(() => {
    const ed = (window as unknown as AidocsWindow).__aidocs!.editor!
    const out: string[] = []
    ed.state.doc.forEach((n: any) => {
      const style = n.attrs?.styleId ? `/${n.attrs.styleId}` : ''
      out.push(`${n.type.name}${style}:${n.textContent}`)
    })
    return out
  })
}

/** select the first occurrence of `text` through the editor and focus it */
async function selectText(editor: Page, text: string): Promise<void> {
  const hit = await editor.evaluate((w) => {
    const ed = (window as unknown as AidocsWindow).__aidocs!.editor!
    let found: { from: number; to: number } | null = null
    // per textblock: a word split across runs (a resized first word) is still one match
    ed.state.doc.descendants((node: any, pos: number) => {
      if (found || !node.isTextblock) return !found
      const idx = (node.textContent as string).indexOf(w)
      if (idx !== -1) found = { from: pos + 1 + idx, to: pos + 1 + idx + w.length }
      return false
    })
    if (found) ed.chain().focus().setTextSelection(found).run()
    return found
  }, text)
  expect(hit, `text not found: ${text}`).not.toBeNull()
  await expect.poll(async () => (await selection(editor)).text, POLL).toBe(text)
  await expectEditorFocused(editor)
}

/** Tiptap's focus() lands on the next animation frame; shortcuts sent before it hit the ribbon field instead */
async function expectEditorFocused(editor: Page): Promise<void> {
  await expect
    .poll(() => editor.evaluate(() => !!document.activeElement?.closest('.ProseMirror')), POLL)
    .toBe(true)
}

async function caretAt(editor: Page, text: string, offset = 0): Promise<void> {
  await selectText(editor, text)
  await editor.evaluate((o) => {
    const ed = (window as unknown as AidocsWindow).__aidocs!.editor!
    ed.chain()
      .focus()
      .setTextSelection(ed.state.selection.from + o)
      .run()
  }, offset)
  await expect.poll(async () => (await selection(editor)).text, POLL).toBe('')
  await expectEditorFocused(editor)
}

async function caretAtEnd(editor: Page): Promise<void> {
  await editor.evaluate(() => {
    const ed = (window as unknown as AidocsWindow).__aidocs!.editor!
    ed.chain()
      .focus()
      .setTextSelection(ed.state.doc.content.size - 1)
      .run()
  })
  await expect.poll(async () => (await selection(editor)).text, POLL).toBe('')
  await expectEditorFocused(editor)
}

/** viewport center of the first occurrence of `word`, scrolled clear of the ribbon first */
async function wordCenter(editor: Page, word: string): Promise<{ x: number; y: number }> {
  return editor.evaluate((w) => {
    const walker = document.createTreeWalker(
      document.querySelector('.doc-page')!,
      NodeFilter.SHOW_TEXT,
    )
    let tn: Node | null
    while ((tn = walker.nextNode())) {
      const i = (tn.textContent ?? '').indexOf(w)
      if (i < 0) continue
      const r = document.createRange()
      r.setStart(tn, i)
      r.setEnd(tn, i + w.length)
      tn.parentElement?.scrollIntoView({ block: 'center' })
      const b = r.getBoundingClientRect()
      return { x: b.left + b.width / 2, y: b.top + b.height / 2 }
    }
    throw new Error(`word not found: ${w}`)
  }, word)
}

async function clickWord(editor: Page, word: string): Promise<void> {
  const c = await wordCenter(editor, word)
  await editor.mouse.click(c.x, c.y)
  await expect.poll(async () => (await selection(editor)).text, POLL).toBe('')
  await editor.waitForTimeout(50)
}

async function dragBy(page: Page, handle: Locator, dx: number): Promise<void> {
  const box = (await handle.boundingBox())!
  const x = box.x + box.width / 2
  const y = box.y + box.height / 2
  await page.mouse.move(x, y)
  await page.mouse.down()
  await page.mouse.move(x + dx / 2, y, { steps: 3 })
  await page.mouse.move(x + dx, y, { steps: 3 })
  await page.mouse.up()
}

/** the context menu opens through a renderer/main handshake; a click while the window settles can be swallowed */
async function openContextMenu(editor: Page, x: number, y: number): Promise<Locator> {
  const menu = editor.locator('.ctx-menu')
  for (let attempt = 0; attempt < 3; attempt++) {
    await editor.mouse.click(x, y, { button: 'right' })
    if (await menu.isVisible({ timeout: 3_000 }).catch(() => false)) return menu
    await editor.waitForTimeout(500)
  }
  await expect(menu).toBeVisible()
  return menu
}

async function documentXml(path: string): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(path))
  return zip.file('word/document.xml')!.async('string')
}

test.describe.serial('docs Word interaction smoke', () => {
  test.describe.configure({ timeout: 60_000 })
  let dir: string
  let launched: LaunchedApp
  let app: ElectronApplication
  let editor: Page
  let savedPath: string

  const tab = (name: string) => editor.locator('.ribbon-tab', { hasText: new RegExp(`^${name}$`) })
  const activeTab = () => editor.locator('.ribbon-tab.active')
  // clicking the active tab collapses the ribbon (Word for Mac), so only switch when needed
  const showTab = async (name: string) => {
    if ((await activeTab().textContent()) !== name) await tab(name).click()
    await expect(activeTab()).toHaveText(name)
  }
  const gallery = () => editor.locator('.style-gallery')
  const docPage = () => editor.locator('.doc-page').first()
  const para = (text: string) => editor.locator('.doc-page p', { hasText: text }).first()
  const layoutMenuItem = (label: string) =>
    editor
      .locator('[data-rb-panel] button')
      .filter({ has: editor.locator('b', { hasText: new RegExp(`^${label}$`) }) })
  const paraAttrs = () =>
    editor.evaluate(() =>
      (window as unknown as AidocsWindow).__aidocs!.editor!.getAttributes('docParagraph'),
    )
  const textAttrs = () =>
    editor.evaluate(() =>
      (window as unknown as AidocsWindow).__aidocs!.editor!.getAttributes('docTextStyle'),
    )
  const undo = () => editor.keyboard.press('ControlOrMeta+z')
  const redo = () => editor.keyboard.press('ControlOrMeta+Shift+z')

  test.beforeAll(async () => {
    test.setTimeout(90_000)
    dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-word-smoke-')))
    launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-word-interaction-smoke',
      settings: { defaultSaveDir: dir },
    })
    app = launched.app
    const { page } = launched
    await expect(page.locator('.quick-card').first()).toContainText('AI Docs')
    await page.locator('.quick-card').first().click()
    editor = await waitForPageWithUrl(app, '://docs/')
    await editor.waitForFunction(
      () => Boolean((window as unknown as AidocsWindow).__aidocs?.editor),
      undefined,
      { timeout: 30_000 },
    )
    await editor.locator('.doc-page[contenteditable="true"]').waitFor()
  })

  test.afterAll(async () => {
    if (launched) await closeAndSaveVideo(launched, 'docs-word-interaction-smoke')
    if (dir) rmSync(dir, { recursive: true, force: true })
  })

  test('1 new document: Word gallery, title, heading, body and list', async () => {
    await expect(gallery().locator('.style-card-label')).toHaveText(WORD_GALLERY)
    await docPage().click()
    await editor.keyboard.type('Smoke Title', { delay: 10 })
    await gallery().locator('.style-card[data-style-id="Title"]').click()
    expect((await paraAttrs()).styleId).toBe('Title')
    await editor.keyboard.press('End')
    await editor.keyboard.press('Enter')
    await editor.keyboard.type('Overview', { delay: 10 })
    await editor.keyboard.press('Alt+ControlOrMeta+Digit1')
    await expect(docPage().locator('h1')).toHaveText(['Overview'])
    await editor.keyboard.press('End')
    await editor.keyboard.press('Enter')
    await editor.keyboard.type('alpha one two three', { delay: 10 })
    await editor.keyboard.press('Enter')
    await editor.keyboard.type('beta four five six', { delay: 10 })
    await editor.keyboard.press('Enter')
    // 400+ synthetic keystrokes drop characters; the long paragraph goes in as one insert
    await editor.evaluate(
      (text) => (window as unknown as AidocsWindow).__aidocs!.editor!.commands.insertContent(text),
      LONG,
    )
    await editor.keyboard.press('Enter')
    await editor.getByRole('button', { name: 'Bullets', exact: true }).first().click()
    await editor.keyboard.type('item one', { delay: 10 })
    await editor.keyboard.press('Enter')
    await editor.keyboard.type('item two', { delay: 10 })
    await editor.keyboard.press('Enter')
    await editor.keyboard.type('item three', { delay: 10 })
    await expect(docPage().locator('.doc-li')).toHaveCount(3)
    // Enter after a heading returns to Normal (Word's next-style rule)
    expect(await blocks(editor)).toEqual([
      'docParagraph/Title:Smoke Title',
      'docHeading/Heading1:Overview',
      'docParagraph:alpha one two three',
      'docParagraph:beta four five six',
      `docParagraph:${LONG}`,
      'docListItem:item one',
      'docListItem:item two',
      'docListItem:item three',
    ])
  })

  test('2 selection bar: line, paragraph, Strong, document, bold toggles', async () => {
    const pageBox = (await docPage().boundingBox())!
    const barX = pageBox.x + 30
    const rowY = async (text: string) => {
      await para(text).scrollIntoViewIfNeeded()
      return (await para(text).boundingBox())!.y + 8
    }

    await editor.mouse.click(barX, await rowY('beta four'))
    expect((await selection(editor)).text).toBe('beta four five six')

    // the first visual line of the wrapped paragraph, then the whole paragraph
    await editor.mouse.click(barX, await rowY('gamma0 '))
    let sel = await selection(editor)
    expect(sel.text.startsWith('gamma0 ')).toBe(true)
    expect(sel.text.length).toBeLessThan(LONG.length)
    await editor.mouse.dblclick(barX, await rowY('gamma0 '))
    sel = await selection(editor)
    expect(sel.text).toBe(LONG)

    await gallery().locator('.style-card[data-style-id="Strong"]').click()
    expect((await textAttrs()).styleId).toBe('Strong')
    await expect(gallery().locator('.style-card[data-style-id="Strong"]')).toHaveClass(/active/)

    await editor.mouse.click(barX, await rowY('gamma0 '), { clickCount: 3 })
    sel = await selection(editor)
    expect(sel.from).toBe(0)
    expect(sel.to).toBe(sel.docSize)

    await editor.keyboard.press('ControlOrMeta+b')
    await expect.poll(async () => (await boldText(editor)).length, POLL).toBeGreaterThan(0)
    await editor.keyboard.press('ControlOrMeta+b')
    await expect.poll(() => boldText(editor), POLL).toEqual([])
  })

  test('3 format painter: locked Heading 1 brush, Esc, undo across steps', async () => {
    const painter = editor.getByRole('button', { name: /^Format Painter/ })
    await clickWord(editor, 'Overview')
    await painter.dblclick()
    await expect(painter).toHaveAttribute('aria-pressed', 'true')
    await clickWord(editor, 'alpha')
    await expect(docPage().locator('h1')).toHaveCount(2)
    await expect(painter).toHaveAttribute('aria-pressed', 'true')
    await clickWord(editor, 'beta')
    await expect(docPage().locator('h1')).toHaveCount(3)
    await editor.keyboard.press('Escape')
    await expect(painter).toHaveAttribute('aria-pressed', 'false')
    await clickWord(editor, 'gamma3 ')
    await editor.waitForTimeout(300)
    await expect(docPage().locator('h1')).toHaveCount(3)

    await undo()
    await expect(docPage().locator('h1')).toHaveCount(2)
    await undo()
    await expect(docPage().locator('h1')).toHaveCount(1)
    // the third undo reaches step 2's second bold toggle
    await undo()
    await expect.poll(async () => (await boldText(editor)).length, POLL).toBeGreaterThan(0)
    await redo()
    await expect.poll(() => boldText(editor), POLL).toEqual([])
    expect(await blocks(editor)).toEqual([
      'docParagraph/Title:Smoke Title',
      'docHeading/Heading1:Overview',
      'docParagraph:alpha one two three',
      'docParagraph:beta four five six',
      `docParagraph:${LONG}`,
      'docListItem:item one',
      'docListItem:item two',
      'docListItem:item three',
    ])
  })

  test('4 list keys: staged Backspace, Tab inside text, Enter on an empty item', async () => {
    const items = docPage().locator('.doc-li')
    const third = para('item three')
    const before = await blocks(editor)
    await caretAt(editor, 'item three')
    await editor.keyboard.press('Backspace')
    await expect(items).toHaveCount(2)
    const indent = await third.evaluate((el) => parseFloat(getComputedStyle(el).marginInlineStart))
    expect(indent).toBeGreaterThan(0)
    expect((await blocks(editor))[7]).toBe('docParagraph:item three')
    // history merges strokes within 500 ms; a user pauses between the stages
    await editor.waitForTimeout(600)
    await editor.keyboard.press('Backspace')
    await expect(third).toHaveCSS('margin-inline-start', '0px')
    expect((await blocks(editor)).length).toBe(8)
    await undo()
    await expect(items).toHaveCount(2)
    await undo()
    await expect(items).toHaveCount(3)
    expect(await blocks(editor)).toEqual(before)
    expect(await boldText(editor)).toEqual([])

    await caretAt(editor, 'item three', 5)
    await editor.keyboard.press('Tab')
    await expect.poll(async () => (await blocks(editor))[7], POLL).toBe('docListItem:item \tthree')

    await caretAt(editor, 'item \tthree', 'item \tthree'.length)
    await editor.keyboard.press('Enter')
    await expect(items).toHaveCount(4)
    await editor.keyboard.press('Enter')
    await expect(items).toHaveCount(3)
    const attrs = await paraAttrs()
    expect(attrs.styleId ?? null).toBeNull()
    expect(attrs.indentLeft ?? null).toBeNull()
    await editor.keyboard.type('closing paragraph', { delay: 10 })
    await expect(para('closing paragraph')).toHaveCSS('margin-inline-start', '0px')
  })

  test('5 font and size boxes: type-ahead, Esc, stepping past 72, mixed selection', async () => {
    const fontBox = editor.locator('input.rb-font-family')
    const sizeBox = editor.locator('input.rb-font-size')
    const menu = editor.locator('.rb-font-family-menu')
    await selectText(editor, 'alpha')
    const shownFont = await fontBox.inputValue()
    const shownSize = await sizeBox.inputValue()
    expect(shownFont).not.toBe('')
    expect(shownSize).not.toBe('')

    await fontBox.click()
    await editor.keyboard.type('ge')
    await expect(menu).toBeVisible()
    await expect(fontBox).toHaveValue('Georgia')
    await expect(menu.locator('button.kbd-focus')).toHaveText('Georgia')
    await editor.keyboard.press('Escape')
    await expect(menu).toBeHidden()
    await expect(fontBox).toHaveValue(shownFont)
    expect((await textAttrs()).fontAscii ?? null).not.toBe('Georgia')

    await selectText(editor, 'alpha')
    await sizeBox.click()
    await editor.keyboard.type('72')
    await editor.keyboard.press('Enter')
    await expect(sizeBox).toHaveValue('72')
    await selectText(editor, 'alpha')
    await editor.keyboard.press('ControlOrMeta+Shift+Period')
    await expect(sizeBox).toHaveValue('80')
    expect((await textAttrs()).sizeHalfPoints).toBe(160)
    await editor.keyboard.press('ControlOrMeta+Shift+Period')
    await expect(sizeBox).toHaveValue('90')

    await selectText(editor, 'alpha one')
    await expect(sizeBox).toHaveValue('')
    await expect(fontBox).toHaveValue(shownFont)

    await undo()
    await undo()
    await undo()
    await selectText(editor, 'alpha')
    await expect(sizeBox).toHaveValue(shownSize)
  })

  test('6 table: grid insert activates Table Design, Select Row, Bottom Border, tab reverts', async () => {
    await caretAt(editor, 'closing paragraph', 'closing paragraph'.length)
    await showTab('Insert')
    await editor.getByRole('button', { name: 'Table', exact: true }).click()
    await editor.locator('.table-picker-grid .table-cell').nth(12).click()
    await expect(activeTab()).toHaveText('Table Design')
    const table = docPage().locator('.doc-table').first()
    await expect(table).toHaveCount(1)
    const cols = await table.locator('tr').first().locator('td, th').count()
    const rowCount = await table.locator('tr').count()
    expect(cols).toBeGreaterThan(1)
    const cell = (r: number, c: number) => table.locator('tr').nth(r).locator('td, th').nth(c)

    await cell(1, 0).click()
    await editor.keyboard.type('cell', { delay: 10 })
    await expect(activeTab()).toHaveText('Table Design')

    await showTab('Table Layout')
    await editor.getByRole('button', { name: 'Select', exact: true }).click()
    await editor.locator('.layout-menu button', { hasText: 'Select Row' }).click()
    await expect(table.locator('.selectedCell')).toHaveCount(cols)

    await cell(1, 0).click()
    await showTab('Table Design')
    const bordersMain = editor.locator('.table-split-main')
    await expect(bordersMain).toHaveText('Borders')
    await bordersMain.click()
    await expect(cell(rowCount - 1, 0)).toHaveCSS('border-bottom-style', 'solid')
    await expect(cell(rowCount - 1, cols - 1)).toHaveCSS('border-bottom-style', 'solid')

    await para('closing paragraph').click()
    await expect(tab('Table Design')).toBeHidden()
    await expect(activeTab()).toHaveText('Insert')
  })

  test('7 ruler: first-line marker, inches in Preferences, ruler and margins dialog follow', async () => {
    await showTab('View')
    await editor.getByRole('button', { name: 'Ruler', exact: true }).click()
    const ruler = editor.locator('.ruler')
    await expect(ruler).toBeVisible()
    // the ruler scrolls with the page: bring the top of the document back
    await editor.evaluate(() => document.querySelector('.editor-scroll')?.scrollTo(0, 0))
    await expect.poll(async () => (await ruler.boundingBox())!.y, POLL).toBeGreaterThan(117)
    const beta = para('beta four')
    await beta.click({ position: { x: 20, y: 8 } })
    await expect.poll(async () => (await selection(editor)).text, POLL).toBe('')
    const marker = editor.getByRole('slider', { name: 'First Line Indent', exact: true })
    await expect(marker).toHaveAttribute('aria-valuenow', '1440')
    await dragBy(editor, marker, 96)
    await expect(beta).toHaveCSS('text-indent', '96px')
    await expect(beta).toHaveCSS('margin-inline-start', '0px')
    await expect(para('alpha one')).toHaveCSS('text-indent', '0px')

    const rulerLabels = () =>
      editor.locator('.ruler-num').evaluateAll((els) => els.map((el) => el.textContent))
    const pickUnit = async (unit: string) => {
      await app.evaluate(({ webContents }) => {
        for (const wc of webContents.getAllWebContents())
          if (wc.getURL().includes('://docs/')) wc.send('menu:command', 'preferences')
      })
      const dialog = editor.getByRole('dialog', { name: 'Preferences' })
      await expect(dialog).toBeVisible()
      await dialog.getByRole('combobox').selectOption(unit)
      await dialog.getByRole('button', { name: 'Close' }).click()
      await expect(dialog).toBeHidden()
    }
    const marginsTop = async () => {
      await showTab('Layout')
      await editor.getByRole('button', { name: 'Margins', exact: true }).click()
      await editor.getByRole('button', { name: 'Custom Margins…' }).click()
      const top = editor.getByRole('textbox', { name: 'Top', exact: true })
      const value = await top.inputValue()
      await editor.getByRole('button', { name: 'Cancel', exact: true }).click()
      await expect(top).toBeHidden()
      return value
    }
    // the English UI starts in inches: one label per inch, numbered from the margin
    const inchLabels = await rulerLabels()
    expect(inchLabels.slice(0, 3)).toEqual(['1', '2', '3'])
    expect(inchLabels.length).toBeLessThan(12)
    expect(await marginsTop()).toBe('1"')

    await pickUnit('cm')
    await expect.poll(async () => (await rulerLabels()).length, POLL).toBeGreaterThan(15)
    expect((await rulerLabels()).slice(0, 3)).toEqual(['1', '2', '3'])
    expect(await marginsTop()).toBe('2.54 cm')
    // the paragraph indent written by the ruler drag survives the unit switch
    await expect(beta).toHaveCSS('text-indent', '96px')

    await pickUnit('in')
    await expect.poll(rulerLabels, POLL).toEqual(inchLabels)
    expect(await marginsTop()).toBe('1"')
    await pickUnit('cm')
    await expect.poll(async () => (await rulerLabels()).length, POLL).toBeGreaterThan(15)
  })

  test('8 layout galleries: Next Page section break, Mirrored margins, A5', async () => {
    const status = editor.locator('.status-page')
    await expect(status).toContainText('of 1')
    const width0 = (await docPage().boundingBox())!.width
    const padLeft0 = await docPage().evaluate((el) => getComputedStyle(el).paddingLeft)

    await caretAt(editor, 'closing paragraph', 'closing paragraph'.length)
    await showTab('Layout')
    await editor.getByRole('button', { name: 'Breaks', exact: true }).click()
    await expect(editor.locator('.layout-menu-head')).toHaveText(['Page Breaks', 'Section Breaks'])
    await layoutMenuItem('Next Page').click()
    // an unsaved document keeps the break as a pending sectPr block until the first save
    await expect(editor.locator('.status-msg')).toContainText('takes effect after saving')
    expect((await blocks(editor)).findIndex((b) => b.startsWith('docProtected'))).toBeGreaterThan(
      (await blocks(editor)).indexOf('docParagraph:closing paragraph'),
    )
    await expect(status).toContainText('of 1')

    await editor.getByRole('button', { name: 'Margins', exact: true }).click()
    await layoutMenuItem('Mirrored').click()
    await expect(docPage()).not.toHaveCSS('padding-left', padLeft0)

    await editor.getByRole('button', { name: 'Size', exact: true }).click()
    await layoutMenuItem('A5').click()
    await expect
      .poll(async () => (await docPage().boundingBox())!.width, POLL)
      .toBeLessThan(width0 * 0.8)
  })

  test('8b undo reverts the page setup changes', async () => {
    test.fixme(
      true,
      'Margins and paper size live outside the editor history: Cmd+Z after Mirrored / A5 undoes the section break and earlier text edits instead (App.tsx onSection / onMirrorMargins setState); Word undoes page setup like any edit',
    )
  })

  test('9 paste: Keep Text Only drops bold, the context menu Paste keeps it', async () => {
    await selectText(editor, 'beta four five six')
    await editor.keyboard.press('ControlOrMeta+b')
    await expect.poll(() => boldText(editor), POLL).toEqual(['beta four five six'])
    await editor.keyboard.press('ControlOrMeta+c')
    await expect
      .poll(() => app.evaluate(({ clipboard }) => clipboard.readText()), POLL)
      .toContain('beta four five six')

    await caretAt(editor, 'closing paragraph', 'closing paragraph'.length)
    await editor.keyboard.press('Enter')
    await showTab('Home')
    await editor.getByTestId('ribbon-paste-menu').click()
    await editor.getByRole('menuitem', { name: 'Keep Text Only' }).click()
    await expect
      .poll(
        async () => (await runs(editor)).filter((r) => r.text.includes('beta four')).length,
        POLL,
      )
      .toBe(2)
    expect(await boldText(editor)).toEqual(['beta four five six'])

    await caretAtEnd(editor)
    await editor.keyboard.press('Enter')
    const sel = await selection(editor)
    const caret = await editor.evaluate((pos) => {
      const ed = (window as unknown as AidocsWindow).__aidocs!.editor!
      const c = ed.view.coordsAtPos(pos)
      return { x: c.left + 4, y: (c.top + c.bottom) / 2 }
    }, sel.from)
    const menu = await openContextMenu(editor, caret.x, caret.y)
    await menu
      .locator('.ctx-item')
      .filter({ has: editor.locator('.ctx-label', { hasText: /^Paste$/ }) })
      .click()
    await expect
      .poll(
        async () => (await runs(editor)).filter((r) => r.text.includes('beta four')).length,
        POLL,
      )
      .toBe(3)
    const pasted = (await runs(editor)).filter((r) => r.text.includes('beta four'))
    expect(pasted.map((r) => r.bold)).toEqual([true, false, true])
  })

  test('10 zoom dialog, slider floor, view buttons', async () => {
    const zoomValue = editor.locator('.zoom-value')
    await expect(zoomValue).toHaveText('100%')
    const width100 = (await docPage().boundingBox())!.width
    const widthRatio = async () => (await docPage().boundingBox())!.width / width100
    const dialog = editor.locator('.zoom-dialog')

    await zoomValue.click()
    await dialog.waitFor()
    await dialog.locator('.zoom-dialog-radio', { hasText: '200%' }).click()
    await dialog.locator('button[type="submit"]').click()
    await expect(dialog).toHaveCount(0)
    await expect(zoomValue).toHaveText('200%')
    await expect.poll(widthRatio, POLL).toBeGreaterThan(1.9)

    const slider = editor.locator('.zoom-slider')
    await slider.focus()
    await editor.keyboard.press('Home')
    await expect(zoomValue).toHaveText('10%')
    await expect.poll(widthRatio, POLL).toBeLessThan(0.11)
    await expect(docPage()).toContainText('closing paragraph')

    await zoomValue.click()
    await dialog.waitFor()
    await dialog.locator('.zoom-dialog-radio', { hasText: '100%' }).click()
    await dialog.locator('button[type="submit"]').click()
    await expect(zoomValue).toHaveText('100%')
    await expect.poll(widthRatio, POLL).toBeCloseTo(1, 1)

    const views = editor.locator('.status-view-btn')
    await views.nth(1).click()
    await expect(editor.locator('.doc-zoom')).toHaveClass(/view-web/)
    await views.nth(0).click()
    await expect(editor.locator('.doc-zoom')).toHaveClass(/view-print/)
    await expect(docPage().locator('.doc-table')).toHaveCount(1)
  })

  test('11 navigation pane: heading click moves the caret, search counts hits', async () => {
    await showTab('View')
    await editor.locator('.rb-big', { hasText: /^Navigation Pane$/ }).click()
    const pane = editor.locator('.nav-pane')
    await pane.waitFor()
    const rows = pane.locator('.nav-row')
    await expect(rows.locator('.nav-row-text')).toHaveText(['Overview'])
    await caretAt(editor, 'closing paragraph')
    await rows.nth(0).click()
    await expect
      .poll(
        () =>
          editor.evaluate(() => {
            const ed = (window as unknown as AidocsWindow).__aidocs!.editor!
            const $p = ed.state.doc.resolve(ed.state.selection.from)
            return `${$p.parent.type.name}:${$p.parent.textContent}:${$p.parentOffset}`
          }),
        POLL,
      )
      .toBe('docHeading:Overview:0')
    await pane.locator('.nav-search-input').fill('beta')
    await expect(pane.locator('.nav-tab.on')).toHaveText('Results')
    await expect(pane.locator('.nav-result')).toHaveCount(3, POLL)
    await expect(pane.locator('.nav-search-count')).toHaveText('1/3')
    await pane.locator('.nav-pane-close').click()
    await expect(pane).toHaveCount(0)
    await expect(docPage().locator('.search-hit')).toHaveCount(0)
  })

  test('12 hyperlink: Cmd/Ctrl+K inserts, the context menu offers Word’s items, Remove keeps the text', async () => {
    await selectText(editor, 'item one')
    await editor.keyboard.press('ControlOrMeta+k')
    const modal = editor.locator('.modal')
    await modal.waitFor()
    await modal.locator('input').nth(1).fill('https://example.com/smoke')
    await modal.locator('input').nth(1).press('Enter')
    const link = docPage().locator('a.doc-link')
    await expect(link).toHaveCount(1)
    await expect(link).toHaveText('item one')

    const c = await wordCenter(editor, 'item one')
    const menu = await openContextMenu(editor, c.x, c.y)
    const labels = await menu
      .locator('.ctx-label')
      .evaluateAll((els) => els.map((el) => el.textContent))
    expect(labels).toEqual(
      expect.arrayContaining([
        'Edit Hyperlink…',
        'Open Hyperlink',
        'Copy Hyperlink',
        'Remove Hyperlink',
      ]),
    )
    expect(labels).not.toContain('Hyperlink…')
    await menu.locator('.ctx-item', { hasText: 'Remove Hyperlink' }).click()
    await expect(link).toHaveCount(0)
    await expect(docPage().locator('.doc-li').first()).toHaveText('item one')
  })

  test('13 save and reopen: styles, table, section break and list survive', async () => {
    expect(await editor.evaluate(() => (window as unknown as AidocsWindow).__aidocs!.save!())).toBe(
      true,
    )
    await expect
      .poll(() => readdirSync(dir).filter((f) => f.endsWith('.docx')), POLL)
      .toHaveLength(1)
    savedPath = join(
      dir,
      readdirSync(dir).find((f) => f.endsWith('.docx'))!,
    )
    const xml = await documentXml(savedPath)
    expect(xml).toContain('<w:pStyle w:val="Title"/>')
    expect(xml).toContain('<w:pStyle w:val="Heading1"/>')
    expect(xml).toContain('<w:rStyle w:val="Strong"/>')
    expect(xml).toContain('<w:tbl>')
    expect(xml.match(/<w:sectPr>/g)).toHaveLength(2)
    expect(xml.match(/<w:numPr>/g)).toHaveLength(3)
    expect(xml).toContain('<w:tab/>')
    expect(xml).toMatch(/<w:tcBorders>[^]*?<w:bottom w:val="single"/)
    expect(xml).not.toContain('<w:hyperlink')
    expect(xml).toContain('<w:pgSz w:w="8391" w:h="11906"/>')
    const zip = await JSZip.loadAsync(readFileSync(savedPath))
    expect(await zip.file('word/settings.xml')!.async('string')).toContain('<w:mirrorMargins/>')

    await editor.evaluate(
      (p) => (window as unknown as AidocsWindow).__aidocs!.openPath!(p),
      savedPath,
    )
    await expect(docPage()).toContainText('closing paragraph')
    // the writer files bulleted items under List Paragraph, as Word's Bullets button does
    await expect
      .poll(() => blocks(editor), POLL)
      .toEqual(
        expect.arrayContaining([
          'docParagraph/Title:Smoke Title',
          'docHeading/Heading1:Overview',
          'docListItem/ListParagraph:item one',
          'docListItem/ListParagraph:item \tthree',
          'docParagraph:closing paragraph',
        ]),
      )
    await expect(docPage().locator('.doc-li')).toHaveCount(3)
    await expect(docPage().locator('.doc-table')).toHaveCount(1)
    await expect(editor.locator('.status-page')).toContainText('of 2')
    await showTab('Home')
    await expect(gallery().locator('.style-card-label')).toHaveText(WORD_GALLERY)
    await selectText(editor, 'gamma0')
    expect((await textAttrs()).styleId).toBe('Strong')
  })
})
