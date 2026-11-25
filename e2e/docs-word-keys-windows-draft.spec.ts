/**
 * More chords of Word for Mac pressed on a real document: field braces typed
 * in place, the screen and window moves, Browse, sections, Outline levels and
 * the keypad's own keys; then Word's Window menu, the way back to Draft after
 * a header or footer, and the style area.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, type LaunchedApp } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-word-keys-windows-draft'
const IS_MAC = process.platform === 'darwin'
const SKIP = 'these chords are Word for Mac’s; other platforms keep their own'

const styled = (style: string, text: string) =>
  `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`
const SECTION_END =
  '<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '</w:sectPr></w:pPr><w:r><w:t>End of section one</w:t></w:r></w:p>'
const lines = (prefix: string, n: number) =>
  Array.from({ length: n }, (_, i) => para(`${prefix} line ${i + 1} needle`)).join('')

const LONG_BODY =
  para('First paragraph') +
  lines('One', 70) +
  SECTION_END +
  para('Section two start') +
  lines('Two', 20) +
  para('Last paragraph') +
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'

const OUTLINE_BODY =
  styled('Heading1', 'Alpha') +
  para('alpha body') +
  styled('Heading2', 'Alpha one') +
  styled('Heading3', 'Alpha one deep') +
  styled('Heading1', 'Beta') +
  styled('Heading2', 'Beta one') +
  para('Last paragraph')

interface Snap {
  /** text of the block the selection's head is in */
  block: string
  /** index of that block among the text blocks */
  index: number
  selected: string
  texts: string[]
}

const snap = (page: Page): Promise<Snap> =>
  page.evaluate(() => {
    interface Node {
      isTextblock: boolean
      textContent: string
      nodeSize: number
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
              selection: { head: number; from: number; to: number }
            }
          }
        }
      }
    ).__aidocs
    const { doc, selection } = editor.state
    const texts: string[] = []
    let index = -1
    doc.descendants((node, pos) => {
      if (!node.isTextblock) return
      if (selection.head > pos && selection.head < pos + node.nodeSize) index = texts.length
      texts.push(node.textContent)
    })
    return {
      block: texts[index] ?? '',
      index,
      selected: doc.textBetween(selection.from, selection.to, '\n'),
      texts,
    }
  })

/** the caret into the paragraph that holds `needle`, `offset` characters in */
const caretIn = (page: Page, needle: string, offset = 0, length = 0) =>
  page.evaluate(
    ({ needle, offset, length }) => {
      const { editor } = (
        window as never as {
          __aidocs: {
            editor: {
              state: {
                doc: {
                  descendants(
                    visit: (
                      node: { isTextblock: boolean; textContent: string },
                      pos: number,
                    ) => void,
                  ): void
                }
              }
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
        const found = node.isTextblock ? node.textContent.indexOf(needle) : -1
        if (at < 0 && found >= 0) at = pos + 1 + found
      })
      if (at < 0) throw new Error(`no paragraph contains "${needle}"`)
      editor
        .chain()
        .focus()
        .setTextSelection({ from: at + offset, to: at + offset + length })
        .run()
    },
    { needle, offset, length },
  )

/** the text blocks whose first line is inside the window, by index */
const blocksInWindow = (page: Page): Promise<number[]> =>
  page.evaluate(() => {
    const frame = document.querySelector('.editor-scroll')!.getBoundingClientRect()
    const blocks = [...document.querySelectorAll('.doc-page > :is(p, h1, h2, h3)')]
    return blocks.flatMap((el, i) => {
      const box = el.getBoundingClientRect()
      return box.top >= frame.top && box.bottom <= frame.bottom ? [i] : []
    })
  })

const scrollTop = (page: Page) => page.locator('.editor-scroll').evaluate((el) => el.scrollTop)

const userData = () => realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-')))

async function open(docPath: string, userDataDir: string) {
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
  await expect(page.locator('.doc-page p', { hasText: 'Last paragraph' })).toBeAttached({
    timeout: 30_000,
  })
  return { launched, page }
}

async function newDoc(body: string) {
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-word-keys-')))
  const docPath = join(dir, 'keys.docx')
  writeFileSync(docPath, await docxWith(body))
  return { dir, docPath }
}

async function close(launched: LaunchedApp, ...dirs: string[]) {
  await closeAndSaveVideo(launched, VIDEO_DIR)
  for (const dir of dirs) rmSync(dir, { recursive: true, force: true })
}

const documentXml = async (docPath: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file('word/document.xml')!.async('string')

interface Item {
  label: string
  type: string
  enabled: boolean
  checked: boolean
}

const menuItems = (app: ElectronApplication, menu: string): Promise<Item[]> =>
  app.evaluate(
    ({ Menu }, menu) =>
      (
        Menu.getApplicationMenu()?.items.find((item) => item.label === menu)?.submenu?.items ?? []
      ).map(({ label, type, enabled, checked }) => ({ label, type, enabled, checked })),
    menu,
  )

const clickMenu = (app: ElectronApplication, menu: string, label: string) =>
  app.evaluate(
    ({ Menu }, [menu, label]) => {
      const top = Menu.getApplicationMenu()?.items.find((item) => item.label === menu)
      top?.submenu?.items.find((candidate) => candidate.label === label)?.click()
    },
    [menu, label],
  )

const menuItem = async (app: ElectronApplication, menu: string, label: string) =>
  (await menuItems(app, menu)).find((item) => item.label === label)

test('Cmd+F9 inserts field braces whose code is typed in place, updated by F9 and saved as a field', async () => {
  test.skip(!IS_MAC, SKIP)
  test.setTimeout(180_000)
  const userDataDir = userData()
  const { dir, docPath } = await newDoc(
    para('Year here') + para('Pages here') + para('Wrap SAVEDATE') + para('Last paragraph'),
  )
  const { launched, page } = await open(docPath, userDataDir)
  try {
    const open = page.locator('.doc-page .doc-field-code-open')
    await caretIn(page, 'Year here', 'Year '.length)
    await page.keyboard.press('Meta+F9')
    await expect(open).toHaveCount(1)
    await expect(open).toHaveText('  ')
    await page.keyboard.type('DATE \\@ "yyyy"')
    // AutoCorrect curls the quotes as they are typed; the field reads them as straight ones
    await expect(open).toHaveText(' DATE \\@ \u201cyyyy\u201d ')
    // Shift+F9 on the code being typed: the field takes its result
    await page.keyboard.press('Shift+F9')
    await expect(open).toHaveCount(0)
    const year = String(new Date().getFullYear())
    await expect.poll(async () => (await snap(page)).texts[0]).toBe(`Year ${year}here`)

    // a selection becomes the code; F9 closes every open code, computed here or not
    await caretIn(page, 'Wrap SAVEDATE', 'Wrap '.length, 'SAVEDATE'.length)
    await page.keyboard.press('Meta+F9')
    await expect(open).toHaveText(' SAVEDATE ')
    await caretIn(page, 'Pages here', 'Pages '.length)
    await page.keyboard.press('Meta+F9')
    await page.keyboard.type('NUMPAGES')
    await expect(open).toHaveCount(2)
    await page.keyboard.press('F9')
    await expect(open).toHaveCount(0)
    await expect.poll(async () => (await snap(page)).texts[1]).toBe('Pages 1here')

    // left open at a save: written as the field it spells, for Word to compute; empty braces as nothing
    await caretIn(page, 'Last paragraph', 'Last '.length)
    await page.keyboard.press('Meta+F9')
    await page.keyboard.type('AUTHOR')
    await caretIn(page, 'Last ', 0)
    await page.keyboard.press('Meta+F9')
    await expect(open).toHaveCount(2)
    await saveFromMenu(launched.app)
    await expect
      .poll(async () => (await documentXml(docPath)).includes('AUTHOR'), { timeout: 20_000 })
      .toBe(true)
    const xml = await documentXml(docPath)
    const codes = [...xml.matchAll(/<w:instrText[^>]*>([^<]*)<\/w:instrText>/g)].map((m) =>
      m[1].trim(),
    )
    expect(codes).toEqual(['DATE \\@ "yyyy"', 'NUMPAGES', 'SAVEDATE', 'AUTHOR'])
    expect(xml).toMatch(new RegExp(`<w:t[^>]*>${year}</w:t>`))
  } finally {
    await close(launched, dir, userDataDir)
  }
})

test('Page Up and Page Down move the caret by a screen, and the window, section and Browse chords go where Word goes', async () => {
  test.skip(!IS_MAC, SKIP)
  test.setTimeout(240_000)
  const userDataDir = userData()
  const { dir, docPath } = await newDoc(LONG_BODY)
  const { launched, page } = await open(docPath, userDataDir)
  try {
    const status = page.locator('.status-page')
    await caretIn(page, 'First paragraph', 2)
    expect(await scrollTop(page)).toBe(0)

    // ---- a screen down and back: the window scrolls and the caret goes with it ----
    await page.keyboard.press('PageDown')
    const down = await snap(page)
    expect(down.index).toBeGreaterThan(5)
    expect(await scrollTop(page)).toBeGreaterThan(100)
    expect(await blocksInWindow(page)).toContain(down.index)
    await page.keyboard.press('PageDown')
    expect((await snap(page)).index).toBeGreaterThan(down.index + 5)
    await page.keyboard.press('PageUp')
    await page.keyboard.press('PageUp')
    expect(await scrollTop(page)).toBe(0)
    expect((await snap(page)).index).toBeLessThan(3)
    // at the top there is nowhere to scroll: the caret goes to the start
    await page.keyboard.press('PageUp')
    expect((await snap(page)).index).toBe(0)

    // ---- with Shift the selection grows by a screen ----
    await caretIn(page, 'One line 2 ', 0)
    await page.keyboard.press('Shift+PageDown')
    const grown = await snap(page)
    expect(grown.selected.startsWith('One line 2 needle\nOne line 3 needle')).toBe(true)
    expect(grown.selected.split('\n').length).toBeGreaterThan(8)
    await page.keyboard.press('Shift+PageUp')
    expect((await snap(page)).selected.split('\n').length).toBeLessThan(3)

    // ---- top and bottom of the window: the caret moves, the window does not ----
    await caretIn(page, 'One line 30 ', 3)
    await page
      .locator('.doc-page p', { hasText: 'One line 30 needle' })
      .evaluate((el) => el.scrollIntoView({ block: 'center' }))
    const held = await scrollTop(page)
    const shown = await blocksInWindow(page)
    await page.keyboard.press('Alt+Meta+PageUp')
    const top = await snap(page)
    expect(await scrollTop(page)).toBe(held)
    expect(Math.abs(top.index - shown[0])).toBeLessThanOrEqual(1)
    await page.keyboard.press('Alt+Meta+PageDown')
    const bottom = await snap(page)
    expect(await scrollTop(page)).toBe(held)
    expect(Math.abs(bottom.index - shown.at(-1)!)).toBeLessThanOrEqual(1)
    expect(bottom.index).toBeGreaterThan(top.index + 5)
    await page.keyboard.press('Alt+Shift+Meta+PageUp')
    const span = (await snap(page)).selected.split('\n').length
    expect(span).toBeGreaterThan(bottom.index - top.index - 2)

    // ---- sections ----
    await caretIn(page, 'One line 5 ', 0)
    await page.keyboard.press('Alt+PageDown')
    expect((await snap(page)).block).toBe('Section two start')
    await page.keyboard.press('Alt+PageDown')
    expect((await snap(page)).block).toBe('Section two start')
    await caretIn(page, 'Two line 9 ', 4)
    await page.keyboard.press('Alt+PageUp')
    expect((await snap(page)).block).toBe('First paragraph')

    // ---- Browse: pages until a Find or a Go To says otherwise ----
    await expect(status).toContainText('1')
    await page.keyboard.press('Meta+PageDown')
    await expect(status).toContainText('2')
    const second = await snap(page)
    expect(second.index).toBeGreaterThan(10)
    await page.keyboard.press('Meta+PageDown')
    expect((await snap(page)).index).toBeGreaterThan(second.index)
    await page.keyboard.press('Meta+PageUp')
    expect((await snap(page)).index).toBe(second.index)
    await page.keyboard.press('Meta+PageUp')
    expect((await snap(page)).index).toBe(0)

    // after a Find, Browse Next and Previous step through its matches
    await page.keyboard.press('Meta+F')
    const find = page.locator('.find-panel input').first()
    await expect(find).toBeFocused()
    await find.fill('line 4')
    const count = page.locator('.find-count')
    await expect(count).toHaveText(/^1\/\d+$/)
    await find.press('Enter')
    await expect(count).toHaveText(/^2\//)
    await page.locator('.doc-page p', { hasText: 'First paragraph' }).click()
    await page.keyboard.press('Meta+PageDown')
    await expect(count).toHaveText(/^3\//)
    await page.keyboard.press('Meta+PageUp')
    await expect(count).toHaveText(/^2\//)
  } finally {
    await close(launched, dir, userDataDir)
  }
})

test('Outline view shows levels from the keyboard, and the keypad keys do what Word gives them', async () => {
  test.skip(!IS_MAC, SKIP)
  test.setTimeout(180_000)
  const userDataDir = userData()
  const { dir, docPath } = await newDoc(OUTLINE_BODY)
  const { launched, page } = await open(docPath, userDataDir)
  try {
    const shownHeadings = () =>
      page.evaluate(() =>
        [...document.querySelectorAll<HTMLElement>('.doc-page > :is(h1, h2, h3)')]
          .filter((el) => el.offsetParent !== null)
          .map((el) => el.textContent?.replace(/[⊕⊖]/g, '').trim()),
      )
    await caretIn(page, 'Alpha', 1)
    // outside Outline view the level chords change nothing
    await page.keyboard.press('Control+Shift+1')
    expect(await shownHeadings()).toHaveLength(5)

    await page.keyboard.press('Alt+Meta+O')
    await expect(page.locator('.outline-handle').first()).toBeVisible()
    await page.keyboard.press('Control+Shift+1')
    await expect.poll(shownHeadings).toEqual(['Alpha', 'Beta'])
    await page.keyboard.press('Control+Shift+2')
    await expect.poll(shownHeadings).toEqual(['Alpha', 'Alpha one', 'Beta', 'Beta one'])
    await page.keyboard.press('Control+Shift+3')
    await expect.poll(shownHeadings).toHaveLength(5)
    await page.keyboard.press('Control+Shift+1')
    await page.keyboard.press('Control+Shift+A')
    await expect.poll(shownHeadings).toHaveLength(5)
    // the keypad's minus and plus fold and open the heading at the caret
    await page.keyboard.press('Control+Shift+NumpadSubtract')
    await expect.poll(shownHeadings).toEqual(['Alpha', 'Beta', 'Beta one'])
    await page.keyboard.press('Control+Shift+NumpadAdd')
    await expect.poll(shownHeadings).toHaveLength(5)
    await page.keyboard.press('Alt+Meta+P')
    await expect(page.locator('.outline-handle')).toHaveCount(0)

    // ---- the keypad's dashes, Enter and Clear ----
    await caretIn(page, 'alpha body', 'alpha'.length)
    await page.keyboard.press('Meta+NumpadSubtract')
    await page.keyboard.press('Alt+Meta+NumpadSubtract')
    expect((await snap(page)).texts[1]).toBe('alpha–— body')
    await caretIn(page, 'alpha', 0, 'alpha'.length)
    // the driver's Numpad5 is the Mac keypad's Clear
    await page.keyboard.press('Numpad5')
    expect((await snap(page)).texts[1]).toBe('–— body')
    await caretIn(page, 'Beta one', 0)
    await page.keyboard.press('Meta+NumpadEnter')
    await saveFromMenu(launched.app)
    await expect
      .poll(async () => (await documentXml(docPath)).includes('–— body'), {
        timeout: 20_000,
      })
      .toBe(true)
    const xml = await documentXml(docPath)
    // a section break of its own before the heading, not a page break
    expect(xml).toMatch(/<w:pPr>(?:(?!<\/w:pPr>).)*<w:sectPr\b/s)
    expect(xml).not.toContain('w:val="continuous"')
    expect(xml).not.toContain('w:type="page"')
  } finally {
    await close(launched, dir, userDataDir)
  }
})

test('the Window menu is Word’s, Split and the window chords work, and a second window is reached from the keyboard', async () => {
  test.setTimeout(180_000)
  const userDataDir = userData()
  const { dir, docPath } = await newDoc(para('Window one') + para('Last paragraph'))
  const { launched, page } = await open(docPath, userDataDir)
  const { app } = launched
  try {
    // between the system's own items: New Window, Arrange All, Merge All Windows, Split
    const ours = (await menuItems(app, 'Window')).map((item) =>
      item.type === 'separator' ? '---' : item.label,
    )
    const from = ours.indexOf('New Window')
    expect(ours.slice(from - 1, from + 5)).toEqual([
      '---',
      'New Window',
      'Arrange All',
      'Merge All Windows',
      'Split',
      '---',
    ])
    expect((await menuItems(app, 'File')).map((item) => item.label)).not.toContain('New Window')
    expect((await menuItem(app, 'Window', 'Arrange All'))?.enabled).toBe(true)
    expect((await menuItem(app, 'Window', 'Merge All Windows'))?.enabled).toBe(false)

    // Split is a tick; Control+Shift+C takes the split off, as Word's Close Pane does
    const pane = page.locator('.split-pane')
    await clickMenu(app, 'Window', 'Split')
    await expect(pane).toBeVisible()
    await expect.poll(async () => (await menuItem(app, 'Window', 'Split'))?.checked).toBe(true)
    if (IS_MAC) {
      await page.locator('.doc-page p', { hasText: 'Window one' }).first().click()
      await page.keyboard.press('Control+Shift+C')
    } else await clickMenu(app, 'Window', 'Split')
    await expect(pane).toHaveCount(0)
    await expect.poll(async () => (await menuItem(app, 'Window', 'Split'))?.checked).toBe(false)

    // New Window opens the document again; the window chords go round the two
    const focused = () =>
      page.evaluate(async () => {
        const tabs = await (
          window as never as {
            desktop: { listDocsTabs(): Promise<{ id: string; focused: boolean }[]> }
          }
        ).desktop.listDocsTabs()
        return { count: tabs.length, at: tabs.findIndex((tab) => tab.focused) }
      })
    await clickMenu(app, 'Window', 'New Window')
    await expect.poll(async () => (await focused()).count, { timeout: 30_000 }).toBe(2)
    await expect.poll(async () => (await focused()).at).toBe(1)
    if (IS_MAC) {
      const second = app
        .windows()
        .filter((candidate) => candidate.url().includes('://docs/'))
        .at(-1)!
      await second.locator('.doc-page').first().click()
      await second.keyboard.press('Meta+F6')
      await expect.poll(async () => (await focused()).at).toBe(0)
      await page.locator('.doc-page p', { hasText: 'Window one' }).first().click()
      await page.keyboard.press('Shift+Meta+F6')
      await expect.poll(async () => (await focused()).at).toBe(1)
    }
  } finally {
    await close(launched, dir, userDataDir)
  }
})

test('Draft comes back after a header is edited from it, and the style area names each paragraph’s style', async () => {
  test.setTimeout(180_000)
  const userDataDir = userData()
  const { dir, docPath } = await newDoc(OUTLINE_BODY)
  const { launched, page } = await open(docPath, userDataDir)
  const { app } = launched
  try {
    const draft = page.locator('.doc-zoom.view-draft')
    const print = page.locator('.doc-zoom.view-print')
    const original = await documentXml(docPath)

    // ---- from Draft into a header and back ----
    await clickMenu(app, 'View', 'Draft')
    await expect(draft).toBeVisible()
    await clickMenu(app, 'View', 'Header and Footer')
    await expect(print).toBeVisible()
    await expect(page.locator('.page-hf-header .page-hf-editor')).toBeFocused()
    await page.keyboard.type('Draft header')
    await page.locator('.hf-close').click()
    await expect(page.locator('.page-hf-editing')).toHaveCount(0)
    await expect(draft).toBeVisible()
    await expect.poll(async () => (await menuItem(app, 'View', 'Draft'))?.checked).toBe(true)
    // from Print Layout the view stays where it was
    await clickMenu(app, 'View', 'Print Layout')
    await clickMenu(app, 'View', 'Header and Footer')
    await expect(page.locator('.page-hf-editing')).toHaveCount(1)
    await page.locator('.hf-close').click()
    await expect(page.locator('.page-hf-editing')).toHaveCount(0)
    await expect(print).toBeVisible()
    await expect(page.locator('.page-hf', { hasText: 'Draft header' }).first()).toBeVisible()

    // ---- the style area: off by default, set in Preferences, Draft and Outline only ----
    const names = page.locator('.style-area-name')
    await clickMenu(app, 'View', 'Draft')
    await expect(draft).toBeVisible()
    await expect(names).toHaveCount(0)
    const textLeft = () =>
      page
        .locator('.doc-page p', { hasText: 'alpha body' })
        .evaluate((el) => el.getBoundingClientRect().left)
    const wrapWidth = () =>
      page
        .locator('.doc-page p', { hasText: 'alpha body' })
        .evaluate((el) => el.getBoundingClientRect().width)
    const before = { left: await textLeft(), width: await wrapWidth() }
    await app.evaluate(({ webContents }) => {
      for (const wc of webContents.getAllWebContents())
        if (wc.getURL().includes('://docs/')) wc.send('menu:command', 'preferences')
    })
    const width = page.getByRole('spinbutton', { name: /Style area width/ })
    await expect(width).toHaveValue('0')
    await width.fill('1')
    await page.getByRole('button', { name: 'Close', exact: true }).click()
    await expect(names).toHaveText([
      'Heading1',
      'Normal',
      'Heading2',
      'Heading3',
      'Heading1',
      'Heading2',
      'Normal',
    ])
    // each name sits on the first line of its paragraph, left of the text, which keeps its width
    const rows = await page.evaluate(() => {
      const tops = (selector: string) =>
        [...document.querySelectorAll(selector)].map((el) => el.getBoundingClientRect())
      const labels = tops('.style-area-name')
      const blocks = tops('.doc-page > :is(p, h1, h2, h3)')
      return labels.map((box, i) => ({
        dy: box.top - blocks[i].top,
        right: box.right,
        left: blocks[i].left,
      }))
    })
    for (const row of rows) {
      expect(Math.abs(row.dy)).toBeLessThan(3)
      expect(row.right).toBeLessThanOrEqual(row.left)
    }
    expect(await wrapWidth()).toBe(before.width)
    expect(await textLeft()).toBeGreaterThan(before.left)

    await clickMenu(app, 'View', 'Outline')
    await expect(names).toHaveText(['Heading1', 'Heading2', 'Heading3', 'Heading1', 'Heading2'])
    await clickMenu(app, 'View', 'Print Layout')
    await expect(print).toBeVisible()
    await expect(names).toHaveCount(0)

    // the setting is this machine's: only the header typed above reaches the file
    await saveFromMenu(app)
    const parts = async () => {
      const zip = await JSZip.loadAsync(readFileSync(docPath))
      const names = Object.keys(zip.files).filter((name) => /^word\/header\d*\.xml$/.test(name))
      const texts = await Promise.all(names.map((name) => zip.file(name)!.async('string')))
      return texts.join('')
    }
    await expect.poll(parts, { timeout: 20_000 }).toContain('Draft header')
    const saved = await documentXml(docPath)
    expect(saved).not.toContain('style-area')
    expect(saved).toContain('Alpha one deep')
    expect(original).toContain('Alpha one deep')
  } finally {
    await close(launched, dir, userDataDir)
  }
})
