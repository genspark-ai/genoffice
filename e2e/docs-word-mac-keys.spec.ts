/**
 * Word for Mac's keyboard on a real document: the chords that used to mean
 * something else here (indent, extend by paragraph, Preferences, Subscript,
 * Normal style, Paste Text Only, Replace, Spelling), a spread of the ones
 * Docs lacked (caret and selection moves, word deletion, the Control twins,
 * character effects, fields, special characters, views), a text field keeping
 * its own keys, and a Customize Keyboard assignment winning over all of it.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import { mkdtempSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, type LaunchedApp } from './helpers'
import { docxWith, para } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-word-mac-keys'
const IS_MAC = process.platform === 'darwin'
const SKIP = 'these chords are Word for Mac’s; other platforms keep their own'

const cell = (text: string) => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`
const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
  `<w:tr>${cell('A1')}${cell('B1')}</w:tr><w:tr>${cell('A2')}${cell('B2')}</w:tr></w:tbl>`

const FIRST = 'alpha one beta'
// long enough to wrap, so a line start is not the paragraph start
const LONG = Array.from({ length: 60 }, (_, i) => `word${String(i).padStart(2, '0')}`).join(' ')
const THIRD = 'delta three'
const BODY = para(FIRST) + para(LONG) + para(THIRD) + '<w:p/>' + TABLE + para('After')

interface Snap {
  /** the text of every paragraph and cell, in order */
  blocks: string[]
  /** [block index, offset] of the two ends of the selection */
  anchor: [number, number]
  head: [number, number]
  selected: string
  style: Record<string, unknown>
  underline: string | null | false
  bold: boolean
  italic: boolean
  strike: boolean
  bullet: boolean
  heading: boolean
  para: Record<string, unknown>
  fields: string[]
}

const snap = (page: Page): Promise<Snap> =>
  page.evaluate(() => {
    interface Node {
      isTextblock: boolean
      textContent: string
      nodeSize: number
      attrs: Record<string, unknown>
      marks: { type: { name: string }; attrs: Record<string, unknown> }[]
    }
    const { editor } = (
      window as never as {
        __aidocs: {
          editor: {
            state: {
              doc: {
                descendants(visit: (node: Node, pos: number) => void): void
                textBetween(from: number, to: number, sep: string): string
              }
              selection: { anchor: number; head: number; from: number; to: number }
            }
            getAttributes(name: string): Record<string, unknown>
            isActive(name: string, attrs?: Record<string, unknown>): boolean
          }
        }
      }
    ).__aidocs
    const { doc, selection } = editor.state
    const starts: number[] = []
    const blocks: string[] = []
    const fields: string[] = []
    let para: Record<string, unknown> = {}
    doc.descendants((node, pos) => {
      if (node.isTextblock) {
        starts.push(pos + 1)
        blocks.push(node.textContent)
        if (selection.head > pos && selection.head < pos + node.nodeSize) para = node.attrs
      }
      for (const mark of node.marks) {
        const name = mark.type.name === 'instrField' ? String(mark.attrs.instr).trim() : ''
        if (name && fields.at(-1) !== name.split(/\s+/)[0]) fields.push(name.split(/\s+/)[0])
      }
    })
    const place = (pos: number): [number, number] => {
      let index = 0
      for (let i = 0; i < starts.length; i++) if (starts[i] <= pos) index = i
      return [index, pos - starts[index]]
    }
    return {
      blocks,
      anchor: place(selection.anchor),
      head: place(selection.head),
      selected: doc.textBetween(selection.from, selection.to, '\n'),
      style: editor.getAttributes('docTextStyle'),
      underline: editor.isActive('underline')
        ? ((editor.getAttributes('underline').line as string | null) ?? null)
        : false,
      bold: editor.isActive('bold'),
      italic: editor.isActive('italic'),
      strike: editor.isActive('strike'),
      bullet: editor.isActive('docListItem', { kind: 'bullet' }),
      heading: /^(heading|docHeading)/i.test(String(para.styleId ?? '')) || !!para.headingLevel,
      para,
      fields,
    }
  })

/** selects `length` characters of the block that reads `text`, from `offset`; a bare caret with no length */
const selectIn = (page: Page, text: string, offset = 0, length = 0) =>
  page.evaluate(
    ({ text, offset, length }) => {
      interface Node {
        isTextblock: boolean
        textContent: string
      }
      const { editor } = (
        window as never as {
          __aidocs: {
            editor: {
              state: { doc: { descendants(visit: (node: Node, pos: number) => void): void } }
              chain(): {
                focus(): {
                  setTextSelection(range: { from: number; to: number }): { run(): boolean }
                }
              }
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
        .setTextSelection({ from: at + offset, to: at + offset + length })
        .run()
    },
    { text, offset, length },
  )

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

async function openDoc(userDataDir: string) {
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-word-mac-keys-')))
  const docPath = join(dir, 'keys.docx')
  writeFileSync(docPath, await docxWith(BODY))
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
  await expect(page.locator('.doc-page p', { hasText: THIRD })).toBeVisible({ timeout: 30_000 })
  return { launched, page, dir }
}

const userData = () => realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-')))

async function close(launched: LaunchedApp, ...dirs: string[]) {
  await closeAndSaveVideo(launched, VIDEO_DIR)
  for (const dir of dirs) rmSync(dir, { recursive: true, force: true })
}

const indentOf = async (page: Page) => Number((await snap(page)).para.indentLeft ?? 0)

test('the chords Word for Mac uses differently do what Word does', async () => {
  test.skip(!IS_MAC, SKIP)
  test.setTimeout(240_000)
  const userDataDir = userData()
  const { launched, page, dir } = await openDoc(userDataDir)
  const { app } = launched
  try {
    // Replace has Word's chord; New Window and Back to Home gave theirs up
    await expect
      .poll(async () => (await menuItem(app, ['Edit', 'Find', 'Replace…'])).accelerator)
      .toBe('Shift+Cmd+H')
    expect((await menuItem(app, ['Window', 'New Window'])).accelerator).toBe('')
    expect((await menuItem(app, ['File', 'Back to Home'])).accelerator).toBe('')

    // Control+Shift+M indents, Shift+Cmd+M takes it back, Control+M still indents
    await selectIn(page, FIRST)
    expect(await indentOf(page)).toBe(0)
    await page.keyboard.press('Control+Shift+KeyM')
    const step = await indentOf(page)
    expect(step).toBeGreaterThan(0)
    await page.keyboard.press('Control+Shift+KeyM')
    expect(await indentOf(page)).toBe(step * 2)
    await page.keyboard.press('Shift+Meta+KeyM')
    expect(await indentOf(page)).toBe(step)
    await page.keyboard.press('Shift+Meta+KeyM')
    expect(await indentOf(page)).toBe(0)
    await page.keyboard.press('Control+KeyM')
    expect(await indentOf(page)).toBe(step)
    await page.keyboard.press('Shift+Meta+KeyM')

    // Option+Shift+Down / Up extend the selection by paragraph and move nothing
    await selectIn(page, THIRD, 6)
    await page.keyboard.press('Alt+Shift+ArrowUp')
    let now = await snap(page)
    expect(now.selected).toBe('delta ')
    expect(now.anchor).toEqual([2, 6])
    await page.keyboard.press('Alt+Shift+ArrowUp')
    now = await snap(page)
    expect(now.head).toEqual([1, 0])
    expect(now.selected).toBe(`${LONG}\ndelta `)
    await selectIn(page, FIRST, 6)
    await page.keyboard.press('Alt+Shift+ArrowDown')
    now = await snap(page)
    expect(now.selected).toBe('one beta\n')
    expect(now.head).toEqual([1, 0])
    expect(now.blocks.slice(0, 3)).toEqual([FIRST, LONG, THIRD])
    // moving a paragraph is on Control+Shift+Down / Up
    await selectIn(page, FIRST, 2)
    await page.keyboard.press('Control+Shift+ArrowDown')
    expect((await snap(page)).blocks.slice(0, 3)).toEqual([LONG, FIRST, THIRD])
    await page.keyboard.press('Control+Shift+ArrowUp')
    expect((await snap(page)).blocks.slice(0, 3)).toEqual([FIRST, LONG, THIRD])

    // Cmd+, and Cmd+. no longer set the text low or high; Shift+Cmd+- is Subscript
    await selectIn(page, FIRST, 6, 3)
    await page.keyboard.press('Meta+Comma')
    await page.keyboard.press('Meta+Period')
    expect((await snap(page)).style.vertAlign ?? null).toBeNull()
    await page.keyboard.press('Shift+Meta+Minus')
    now = await snap(page)
    expect(now.style.vertAlign).toBe('subscript')
    expect(now.blocks[0]).toBe(FIRST)
    await page.keyboard.press('Shift+Meta+Equal')
    expect((await snap(page)).style.vertAlign).toBe('superscript')
    await page.keyboard.press('Shift+Meta+Equal')
    expect((await snap(page)).style.vertAlign ?? null).toBeNull()
    // the non-breaking hyphen moved to Option+Shift+Cmd+-
    await selectIn(page, THIRD, 5)
    await page.keyboard.press('Alt+Shift+Meta+Minus')
    expect((await snap(page)).blocks[2]).toBe('delta‑ three')
    await page.keyboard.press('Meta+KeyZ')
    expect((await snap(page)).blocks[2]).toBe(THIRD)

    // Shift+Cmd+N is the Normal style
    await selectIn(page, THIRD, 2)
    await page.keyboard.press('Alt+Meta+Digit1')
    const heading = page.locator('.doc-page h1', { hasText: THIRD })
    await expect(heading).toHaveCount(1)
    await page.keyboard.press('Shift+Meta+KeyN')
    await expect(heading).toHaveCount(0)
    await expect(page.locator('.doc-page p', { hasText: THIRD })).toHaveCount(1)

    // Shift+Cmd+V pastes the text alone; Option+Cmd+C / V carry the formatting
    await app.evaluate(({ clipboard }) =>
      clipboard.write({ text: 'pasted', html: '<b><i>pasted</i></b>' }),
    )
    await selectIn(page, THIRD, 5)
    await page.keyboard.press('Shift+Meta+KeyV')
    await expect.poll(async () => (await snap(page)).blocks[2]).toBe('deltapasted three')
    await selectIn(page, 'deltapasted three', 6, 4)
    now = await snap(page)
    expect([now.bold, now.italic]).toEqual([false, false])
    await page.keyboard.press('Meta+KeyB')
    await page.keyboard.press('Alt+Meta+KeyC')
    await selectIn(page, FIRST, 0, 5)
    expect((await snap(page)).bold).toBe(false)
    await page.keyboard.press('Alt+Meta+KeyV')
    await selectIn(page, FIRST, 1, 2)
    expect((await snap(page)).bold).toBe(true)

    // Shift+Cmd+H opens Replace
    await selectIn(page, FIRST)
    await page.keyboard.press('Shift+Meta+KeyH')
    const find = page.locator('.find-panel')
    await expect(find.locator('input[placeholder^="Replace"]')).toBeVisible()
    await page.keyboard.press('Escape')
    await expect(find).toBeHidden()

    // F7 is Spelling and Grammar
    await selectIn(page, FIRST)
    await page.keyboard.press('F7')
    const spelling = page.locator('.spelling-dialog')
    await expect(spelling).toBeVisible()
    await page.keyboard.press('Escape')
    await expect(spelling).toBeHidden()
  } finally {
    await close(launched, dir, userDataDir)
  }
})

test('the Word for Mac chords Docs lacked work on a document', async () => {
  test.skip(!IS_MAC, SKIP)
  test.setTimeout(300_000)
  const userDataDir = userData()
  const { launched, page, dir } = await openDoc(userDataDir)
  const { app } = launched
  const head = async () => (await snap(page)).head
  try {
    // ---- navigation ----
    await selectIn(page, LONG, LONG.length - 3)
    // a line move is the browser's, which the editor reads back a moment later
    await page.keyboard.press('Home')
    await expect.poll(async () => (await head())[1]).toBeLessThan(LONG.length - 3)
    expect((await head())[1]).toBeGreaterThan(0)
    await page.keyboard.press('End')
    await expect.poll(head).toEqual([1, LONG.length])
    await page.keyboard.press('Meta+ArrowUp')
    expect(await head()).toEqual([1, 0])
    await page.keyboard.press('Meta+ArrowUp')
    expect(await head()).toEqual([0, 0])
    await page.keyboard.press('Meta+ArrowDown')
    expect(await head()).toEqual([1, 0])
    await page.keyboard.press('Alt+ArrowDown')
    expect(await head()).toEqual([2, 0])
    await page.keyboard.press('Alt+ArrowUp')
    expect(await head()).toEqual([1, 0])
    await page.keyboard.press('Meta+End')
    let now = await snap(page)
    expect(now.head).toEqual([now.blocks.length - 1, 'After'.length])
    await page.keyboard.press('Meta+Home')
    expect(await head()).toEqual([0, 0])
    await page.keyboard.press('Alt+End')
    expect((await head())[0]).toBe(now.blocks.length - 1)
    await page.keyboard.press('Alt+Home')
    expect(await head()).toEqual([0, 0])

    // ---- selection ----
    await selectIn(page, FIRST, 6)
    await page.keyboard.press('Shift+End')
    await expect.poll(async () => (await snap(page)).selected).toBe('one beta')
    await selectIn(page, FIRST, 6)
    await page.keyboard.press('Shift+Home')
    await expect.poll(async () => (await snap(page)).selected).toBe('alpha ')
    await selectIn(page, FIRST, 6)
    await page.keyboard.press('Shift+Meta+ArrowDown')
    expect((await snap(page)).selected).toBe('one beta\n')
    await selectIn(page, THIRD, 5)
    await page.keyboard.press('Shift+Meta+ArrowUp')
    expect((await snap(page)).selected).toBe('delta')
    await page.keyboard.press('Shift+Meta+Home')
    now = await snap(page)
    expect(now.head).toEqual([0, 0])
    expect(now.anchor).toEqual([2, 5])
    await selectIn(page, THIRD, 5)
    await page.keyboard.press('Shift+Meta+End')
    now = await snap(page)
    expect(now.anchor).toEqual([2, 5])
    expect(now.head).toEqual([now.blocks.length - 1, 'After'.length])
    await page.keyboard.press('Control+KeyA')
    now = await snap(page)
    expect(now.selected.startsWith(FIRST)).toBe(true)
    expect(now.selected.endsWith('After')).toBe(true)

    // ---- editing ----
    await selectIn(page, FIRST, FIRST.length)
    await page.keyboard.press('Meta+Backspace')
    expect((await snap(page)).blocks[0]).toBe('alpha one ')
    await page.keyboard.press('Control+KeyZ')
    expect((await snap(page)).blocks[0]).toBe(FIRST)
    await page.keyboard.press('Control+KeyY')
    expect((await snap(page)).blocks[0]).toBe('alpha one ')
    await page.keyboard.press('Meta+KeyZ')
    await selectIn(page, FIRST, 0)
    await page.keyboard.press('Meta+Delete')
    await expect.poll(async () => (await snap(page)).blocks[0].trimStart()).toBe('one beta')
    await page.keyboard.press('Meta+KeyZ')
    expect((await snap(page)).blocks[0]).toBe(FIRST)
    // at the start of a paragraph the word deletion joins it to the one before
    await selectIn(page, THIRD, 0)
    await page.keyboard.press('Meta+Backspace')
    expect((await snap(page)).blocks[1]).toBe(LONG + THIRD)
    await page.keyboard.press('Meta+KeyZ')
    expect((await snap(page)).blocks.slice(0, 3)).toEqual([FIRST, LONG, THIRD])
    // Control+C / Control+V and F3 / F4 are Copy and Paste
    await app.evaluate(({ clipboard }) => clipboard.clear())
    await selectIn(page, FIRST, 0, 5)
    await page.keyboard.press('Control+KeyC')
    await expect.poll(() => app.evaluate(({ clipboard }) => clipboard.readText())).toBe('alpha')
    await selectIn(page, THIRD, THIRD.length)
    await page.keyboard.press('Control+KeyV')
    await expect.poll(async () => (await snap(page)).blocks[2]).toBe(`${THIRD}alpha`)
    await selectIn(page, FIRST, 6, 3)
    await page.keyboard.press('F3')
    await expect.poll(() => app.evaluate(({ clipboard }) => clipboard.readText())).toBe('one')
    await selectIn(page, `${THIRD}alpha`, THIRD.length + 5)
    await page.keyboard.press('F4')
    await expect.poll(async () => (await snap(page)).blocks[2]).toBe(`${THIRD}alphaone`)
    await selectIn(page, `${THIRD}alphaone`, THIRD.length, 8)
    await page.keyboard.press('F2')
    await expect.poll(async () => (await snap(page)).blocks[2]).toBe(THIRD)
    await expect.poll(() => app.evaluate(({ clipboard }) => clipboard.readText())).toBe('alphaone')

    // F5 is Go To, Shift+F4 repeats the find
    const find = page.locator('.find-panel')
    await selectIn(page, FIRST)
    await page.keyboard.press('F5')
    await expect(find.locator('.find-goto-input').first()).toBeVisible()
    await page.keyboard.press('Escape')
    await expect(find).toBeHidden()
    await selectIn(page, FIRST)
    await page.keyboard.press('Meta+KeyF')
    await find.locator('input').first().fill('word0')
    const count = find.locator('.find-count')
    await expect(count).toHaveText('1/10')
    await page.keyboard.press('Shift+F4')
    await expect(count).toHaveText('2/10')
    await page.keyboard.press('Shift+Meta+F4')
    await expect(count).toHaveText('3/10')
    // the Find box keeps its keys: Home, Control+A and Cmd+Delete are its own
    const before = await snap(page)
    await find.locator('input').first().click()
    await expect(find.locator('input').first()).toBeFocused()
    await page.keyboard.press('Home')
    await page.keyboard.press('Control+KeyA')
    await page.keyboard.press('Shift+Meta+KeyX')
    await page.keyboard.press('Meta+Backspace')
    const after = await snap(page)
    expect(after.blocks).toEqual(before.blocks)
    expect(after.strike).toBe(false)
    await page.keyboard.press('Escape')
    await expect(find).toBeHidden()

    // ---- character formatting ----
    await selectIn(page, FIRST, 6, 3)
    await page.keyboard.press('Shift+Meta+KeyA')
    expect((await snap(page)).style.caps).toBe('all')
    await page.keyboard.press('Shift+Meta+KeyK')
    expect((await snap(page)).style.caps).toBe('small')
    await page.keyboard.press('Shift+Meta+KeyK')
    expect((await snap(page)).style.caps).toBe('none')
    await page.keyboard.press('Shift+Meta+KeyD')
    expect((await snap(page)).underline).toBe('double')
    await page.keyboard.press('Shift+Meta+KeyW')
    expect((await snap(page)).underline).toBe('words')
    await page.keyboard.press('Shift+Meta+KeyW')
    expect((await snap(page)).underline).toBe(false)
    await page.keyboard.press('Shift+Meta+KeyX')
    expect((await snap(page)).strike).toBe(true)
    await page.keyboard.press('Shift+Meta+KeyX')
    expect((await snap(page)).strike).toBe(false)
    await page.keyboard.press('Control+KeyB')
    expect((await snap(page)).bold).toBe(true)
    await page.keyboard.press('Control+Shift+KeyB')
    expect((await snap(page)).bold).toBe(false)
    await page.keyboard.press('Shift+Meta+KeyI')
    expect((await snap(page)).italic).toBe(true)
    await page.keyboard.press('Control+KeyI')
    expect((await snap(page)).italic).toBe(false)
    await page.keyboard.press('Control+Shift+KeyU')
    expect((await snap(page)).underline).toBe(null)
    await page.keyboard.press('Shift+Meta+KeyU')
    expect((await snap(page)).underline).toBe(false)
    await page.keyboard.press('Control+Shift+KeyH')
    expect((await snap(page)).style.vanish).toBe(true)
    await page.keyboard.press('Control+Shift+KeyH')
    expect((await snap(page)).style.vanish).toBe(false)

    // ---- paragraph ----
    await selectIn(page, THIRD, 2)
    await page.keyboard.press('Shift+Meta+KeyL')
    expect((await snap(page)).bullet).toBe(true)
    await page.keyboard.press('Shift+Meta+KeyL')
    expect((await snap(page)).bullet).toBe(true)
    await page.keyboard.press('Meta+KeyZ')
    expect((await snap(page)).bullet).toBe(false)
    await selectIn(page, FIRST, 2)
    await page.keyboard.press('Meta+KeyE')
    await expect.poll(async () => (await snap(page)).para.align).toBe('center')
    await page.keyboard.press('Alt+Meta+KeyQ')
    await expect.poll(async () => (await snap(page)).para.align ?? null).not.toBe('center')

    // ---- special characters ----
    await selectIn(page, THIRD, 5)
    await page.keyboard.press('Control+Shift+Space')
    expect((await snap(page)).blocks[2]).toBe('delta  three')
    await page.keyboard.press('Meta+KeyZ')
    await selectIn(page, THIRD, 5)
    await page.keyboard.press('Alt+Meta+Minus')
    expect((await snap(page)).blocks[2]).toBe('delta­ three')
    await page.keyboard.press('Meta+KeyZ')
    // Option+Tab puts a tab in a cell, where Tab moves to the next one
    await selectIn(page, 'A1', 2)
    await page.keyboard.press('Alt+Tab')
    now = await snap(page)
    expect(now.blocks).toContain('A1\t')
    await page.keyboard.press('Control+Tab')
    expect((await snap(page)).blocks).toContain('A1\t\t')
    await page.keyboard.press('Meta+KeyZ')
    await page.keyboard.press('Meta+KeyZ')

    // ---- tables ----
    const rows = page.locator('.doc-page table tr')
    await expect(rows).toHaveCount(2)
    await selectIn(page, 'A2')
    await page.keyboard.press('Control+Meta+KeyI')
    await expect(rows).toHaveCount(3)
    await expect(rows.nth(2)).toContainText('A2')
    await page.keyboard.press('Meta+KeyZ')
    await expect(rows).toHaveCount(2)

    // the caret is in no word: Thesaurus asks for one
    await selectIn(page, '')
    await page.keyboard.press('Shift+F7')
    await expect(page.locator('.status-msg')).toContainText('Select a word or a short phrase first')

    // ---- fields ----
    await selectIn(page, '')
    await page.keyboard.press('Control+Shift+KeyD')
    await expect.poll(async () => (await snap(page)).fields).toEqual(['DATE'])
    // Shift+F9 shows the code of the field the caret is in, and hides it again
    const code = page.locator('.doc-page .doc-field-code')
    await page.keyboard.press('ArrowLeft')
    await page.keyboard.press('ArrowLeft')
    await page.keyboard.press('Shift+F9')
    await expect(code).toHaveText(/\{ DATE/)
    await page.keyboard.press('Shift+F9')
    await expect(code).toHaveCount(0)
    // back behind the field: the next ones go after it, not into it
    await page.keyboard.press('End')
    await expect
      .poll(async () => {
        const at = await snap(page)
        return at.head[1] === at.blocks[at.head[0]].length
      })
      .toBe(true)
    await page.keyboard.press('Control+Shift+KeyT')
    await page.keyboard.press('Control+Shift+KeyP')
    await expect.poll(async () => (await snap(page)).fields).toEqual(['DATE', 'TIME', 'PAGE'])

    // ---- dialogs and panes ----
    await selectIn(page, FIRST, 2)
    await page.keyboard.press('Shift+Meta+F5')
    const bookmark = page.locator('.bookmark-dialog')
    await expect(bookmark).toBeVisible()
    await page.keyboard.press('Escape')
    await expect(bookmark).toBeHidden()
    await selectIn(page, FIRST, 6, 3)
    await page.keyboard.press('Alt+Shift+Meta+KeyX')
    const markEntry = page.locator('.mark-entry-dialog')
    await expect(markEntry).toBeVisible()
    await page.keyboard.press('Escape')
    await expect(markEntry).toBeHidden()
    await selectIn(page, FIRST, 2)
    const stylesPane = page.locator('.styles-pane')
    await page.keyboard.press('Alt+Shift+Meta+KeyS')
    await expect(stylesPane).toBeVisible()
    await page.keyboard.press('Alt+Shift+Meta+KeyS')
    await expect(stylesPane).toBeHidden()

    // ---- views ----
    const checked = async (label: string) => (await menuItem(app, ['View', label])).checked
    await selectIn(page, FIRST, 2)
    await expect.poll(() => checked('Print Layout')).toBe(true)
    await page.keyboard.press('Alt+Meta+KeyN')
    await expect.poll(() => checked('Draft')).toBe(true)
    await page.keyboard.press('Alt+Meta+KeyO')
    await expect.poll(() => checked('Outline')).toBe(true)
    await page.keyboard.press('Alt+Meta+KeyP')
    await expect.poll(() => checked('Print Layout')).toBe(true)
    const split = page.locator('.split-pane')
    await page.keyboard.press('Alt+Meta+KeyS')
    await expect(split.first()).toBeVisible()
    await page.keyboard.press('Alt+Meta+KeyS')
    await expect(split).toHaveCount(0)
  } finally {
    await close(launched, dir, userDataDir)
  }
})

test('a Customize Keyboard assignment wins over a Word for Mac chord', async () => {
  test.skip(!IS_MAC, SKIP)
  test.setTimeout(180_000)
  const userDataDir = userData()
  // as Customize Keyboard stores it: Word Count takes Control+Shift+M from Increase Indent,
  // Strikethrough takes Cmd+Up from the paragraph move
  writeFileSync(
    join(userDataDir, 'docs-keyboard.json'),
    JSON.stringify({
      'word-count': ['Ctrl+Shift+M'],
      'indent-increase': ['Ctrl+M'],
      strikethrough: ['Shift+Cmd+X', 'Cmd+Up'],
      'caret-move:para-up': ['Alt+Up'],
    }),
  )
  const { launched, page, dir } = await openDoc(userDataDir)
  try {
    await selectIn(page, FIRST, 2)
    await page.keyboard.press('Control+Shift+KeyM')
    const wordCount = page.locator('.word-count-dialog')
    await expect(wordCount).toBeVisible()
    await page.keyboard.press('Escape')
    await expect(wordCount).toBeHidden()
    await selectIn(page, FIRST, 2)
    expect(await indentOf(page)).toBe(0)
    // the chords the user left alone keep Word's meaning
    await page.keyboard.press('Control+KeyM')
    expect(await indentOf(page)).toBeGreaterThan(0)
    await page.keyboard.press('Shift+Meta+KeyM')
    expect(await indentOf(page)).toBe(0)

    await selectIn(page, THIRD, 0, 5)
    await page.keyboard.press('Meta+ArrowUp')
    let now = await snap(page)
    expect(now.strike).toBe(true)
    expect(now.selected).toBe('delta')
    await page.keyboard.press('Meta+ArrowUp')
    await selectIn(page, THIRD, 4)
    await page.keyboard.press('Alt+ArrowUp')
    now = await snap(page)
    expect(now.head).toEqual([2, 0])
    await page.keyboard.press('Meta+ArrowDown')
    expect((await snap(page)).head).toEqual([3, 0])
  } finally {
    await close(launched, dir, userDataDir)
  }
})
