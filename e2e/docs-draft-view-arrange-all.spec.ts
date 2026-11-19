/**
 * View > Draft and View > Arrange All, the last two commands that were shown
 * greyed out. Draft is a view: the saved file and the print layout it returns
 * to are checked to be what they were. Arrange All is checked on the windows
 * themselves.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, pictureRun, saveFromMenu } from './docs-objects-fixture'
import { DRAFT_HEADER, draftDocx } from './docs-draft-fixture'

const VIDEO_DIR = 'docs-draft-view-arrange-all'

const viewItem = (app: ElectronApplication, label: string) =>
  app.evaluate(({ Menu }, label) => {
    const view = Menu.getApplicationMenu()?.items.find((item) => item.label === 'View')
    const item = view?.submenu?.items.find((candidate) => candidate.label === label)
    return item ? { checked: item.checked, enabled: item.enabled } : null
  }, label)

const clickViewItem = (app: ElectronApplication, label: string) =>
  app.evaluate(({ Menu }, label) => {
    const view = Menu.getApplicationMenu()?.items.find((item) => item.label === 'View')
    view?.submenu?.items.find((candidate) => candidate.label === label)?.click()
  }, label)

const part = async (docPath: string, name: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file(name)?.async('string') ?? ''

/** the text of each page of a PDF */
async function pdfPages(path: string): Promise<string[]> {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs')
  const doc = await pdfjs.getDocument({ data: new Uint8Array(readFileSync(path)) }).promise
  const out: string[] = []
  for (let i = 1; i <= doc.numPages; i++) {
    const content = await (await doc.getPage(i)).getTextContent()
    out.push(content.items.map((it) => ('str' in it ? it.str : '')).join(' '))
  }
  return out
}

const ribbonButton = (page: Page, name: string) =>
  page.locator('.ribbon :is(.rb-big, .rb-small)', { hasText: new RegExp(`^${name}$`) })

const openViewTab = (page: Page) =>
  page
    .locator('.ribbon-tab', { hasText: /^View$/ })
    .first()
    .click()

/** what the print layout looks like, to compare before Draft and after it */
const printLayout = (page: Page) =>
  page.evaluate(() => {
    const gaps = Array.from(document.querySelectorAll('.doc-page .page-gap'))
    return {
      pages: document.querySelector('.status-page')?.textContent?.match(/of \d+/)?.[0],
      // the paragraph each page opens with
      leads: gaps.map((gap) => gap.nextElementSibling?.textContent ?? ''),
    }
  })

/** how far the editor can be scrolled sideways, and where the Draft rules sit in its visible box */
const scrollerBox = (page: Page) =>
  page.evaluate(() => {
    const scroller = document.querySelector('.editor-scroll')!
    const box = scroller.getBoundingClientRect()
    const left = box.left + scroller.clientLeft
    return {
      scrollWidth: scroller.scrollWidth,
      clientWidth: scroller.clientWidth,
      // each rule's edges, measured from the edges of what the scroller shows
      rules: Array.from(scroller.querySelectorAll('.draft-break')).map((rule) => {
        const r = rule.getBoundingClientRect()
        return [Math.round(r.left - left), Math.round(left + scroller.clientWidth - r.right)]
      }),
    }
  })

async function setZoom(page: Page, percent: string) {
  await page.locator('.zoom-value').click()
  const dialog = page.locator('.zoom-dialog')
  await dialog.locator('.zoom-dialog-radio', { hasText: percent }).click()
  await dialog.locator('button[type="submit"]').click()
  await expect(dialog).toHaveCount(0)
  await expect(page.locator('.zoom-value')).toHaveText(percent)
}

/**
 * Draft must not widen what can be scrolled: its rules span the visible box
 * and nothing more. Checked against Print Layout at the same window and zoom.
 */
async function expectRulesFitTheWindow(app: ElectronApplication, page: Page) {
  const canvas = page.locator('.doc-zoom')
  await clickViewItem(app, 'Print Layout')
  await expect(canvas).toHaveClass(/\bview-print\b/)
  await expect.poll(async () => (await scrollerBox(page)).rules.length).toBe(0)
  // a resized window may refit the zoom a moment later: read once the layout holds still
  let last = ''
  await expect
    .poll(
      async () => {
        const now = JSON.stringify([
          await scrollerBox(page),
          await page.locator('.zoom-value').textContent(),
        ])
        const still = now === last
        last = now
        return still
      },
      { timeout: 30_000, intervals: [1000] },
    )
    .toBe(true)
  const print = await scrollerBox(page)
  const fits = print.scrollWidth === print.clientWidth
  await clickViewItem(app, 'Draft')
  await expect(canvas).toHaveClass(/\bview-draft\b/)
  await expect.poll(async () => (await scrollerBox(page)).rules.length).toBeGreaterThan(0)
  const edges = (n: number) => Array.from({ length: n }, () => [0, 0])
  await expect
    .poll(async () => {
      const draft = await scrollerBox(page)
      return { scrollWidth: draft.scrollWidth, rules: draft.rules }
    })
    .toEqual({
      scrollWidth: print.scrollWidth,
      rules: edges((await scrollerBox(page)).rules.length),
    })
  if (fits) return true
  // scrolled sideways as far as it goes, the rules still span what is shown
  await page.locator('.editor-scroll').evaluate((el) => (el.scrollLeft = el.scrollWidth))
  await expect
    .poll(async () => (await scrollerBox(page)).rules)
    .toEqual(edges((await scrollerBox(page)).rules.length))
  await page.locator('.editor-scroll').evaluate((el) => (el.scrollLeft = 0))
  return false
}

test('Draft shows the text as one flow with its page breaks, and changes nothing in the file', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-draft-')))
  const docPath = join(dir, 'report.docx')
  writeFileSync(docPath, await draftDocx())
  const original = {
    header: await part(docPath, 'word/header1.xml'),
    footer: await part(docPath, 'word/footer1.xml'),
  }
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const canvas = page.locator('.doc-zoom')
    const gaps = page.locator('.doc-page .page-gap')
    const shape = page.locator('.doc-page .doc-textbox').first()
    const headers = page.locator('.page-hf', { hasText: DRAFT_HEADER })
    await expect(page.locator('.doc-page p', { hasText: 'Opening paragraph' })).toBeVisible({
      timeout: 30_000,
    })
    // the first automatic break falls inside the run of "First" lines
    await expect
      .poll(async () => (await printLayout(page)).leads[0] ?? '', { timeout: 30_000 })
      .toMatch(/^First line \d+$/)
    const before = await printLayout(page)
    expect(before.leads.length).toBeGreaterThanOrEqual(3)
    await expect(headers.first()).toBeVisible()
    await expect(shape).toBeVisible()

    // the menu item is live and is the layout's tick
    await expect.poll(async () => (await viewItem(app, 'Draft'))?.enabled).toBe(true)
    await clickViewItem(app, 'Draft')
    await expect(canvas).toHaveClass(/\bview-draft\b/)
    await expect
      .poll(async () => [
        (await viewItem(app, 'Draft'))?.checked,
        (await viewItem(app, 'Print Layout'))?.checked,
      ])
      .toEqual([true, false])

    // no pages: no gaps between them, no header or footer, no floating shape
    await expect(gaps).toHaveCount(0)
    await expect(page.locator('.page-hf:visible')).toHaveCount(0)
    await expect(shape).toBeHidden()
    // the page count is still the print layout's
    await expect(page.locator('.status-page')).toContainText(before.pages!)

    // each automatic break is a rule above the paragraph that opened that page in Print Layout
    const rules = page.locator('.editor-scroll > .draft-breaks > .draft-break')
    // the manual page break and the section break open pages too, and get no rule
    const automatic = before.leads.filter((text) => /^First line \d+$/.test(text))
    expect(automatic.length).toBeLessThan(before.leads.length)
    await expect(rules).toHaveCount(automatic.length)
    for (const [i, text] of automatic.entries()) {
      const lead = page.locator('.doc-page p', { hasText: new RegExp(`^${text}$`) })
      await expect
        .poll(async () => {
          const rule = (await rules.nth(i).boundingBox())!
          const below = (await lead.boundingBox())!
          const above = (await lead.locator('xpath=preceding-sibling::p[1]').boundingBox())!
          return rule.y >= above.y + above.height - 1 && rule.y <= below.y + 1
        })
        .toBe(true)
    }
    // the breaks the document asks for carry their names, formatting marks off
    await expect(canvas).not.toHaveClass(/\bshow-marks\b/)
    await expect(page.locator('.doc-page .doc-fmt-pagebr')).toHaveText('Page Break')
    await expect(page.locator('.doc-page .doc-fmt-sectbr')).toHaveText('Section Break (Next Page)')
    await page.screenshot({ path: test.info().outputPath('draft.png') })

    // the text is edited here as anywhere
    await page.locator('.doc-page p', { hasText: 'Opening paragraph' }).click()
    await page.keyboard.press('End')
    await page.keyboard.type(' edited in draft')
    await saveFromMenu(app)
    await expect
      .poll(() => part(docPath, 'word/document.xml'), { timeout: 30_000 })
      .toContain('edited in draft')
    const saved = await part(docPath, 'word/document.xml')
    // what Draft does not show is still in the file
    expect(saved).toContain('<wp:anchor')
    expect(saved).toMatch(/<w:br w:type="page"\s*\/>/)
    expect(saved.match(/<w:sectPr/g)).toHaveLength(2)
    expect(saved.match(/<w:headerReference /g)).toHaveLength(2)
    expect(await part(docPath, 'word/header1.xml')).toBe(original.header)
    expect(await part(docPath, 'word/footer1.xml')).toBe(original.footer)

    // back in Print Layout the pages are the ones it had
    await openViewTab(page)
    await expect(ribbonButton(page, 'Draft')).toHaveClass(/\bactive\b/)
    await ribbonButton(page, 'Print Layout').click()
    await expect(canvas).toHaveClass(/\bview-print\b/)
    await expect.poll(() => printLayout(page), { timeout: 30_000 }).toEqual(before)
    await expect(rules).toHaveCount(0)
    await expect(page.locator('.doc-page .doc-fmt-chip')).toHaveCount(0)
    await expect(shape).toBeVisible()
    await expect(headers.first()).toBeVisible()

    // the ribbon button runs the same command, and a second press returns to Print Layout
    await ribbonButton(page, 'Draft').click()
    await expect(canvas).toHaveClass(/\bview-draft\b/)
    await expect.poll(async () => (await viewItem(app, 'Draft'))?.checked).toBe(true)
    await ribbonButton(page, 'Draft').click()
    await expect(canvas).toHaveClass(/\bview-print\b/)
    await expect.poll(() => printLayout(page), { timeout: 30_000 }).toEqual(before)

    // the rules span the visible width and add nothing to scroll: normal window, narrow window, 200%
    expect(await expectRulesFitTheWindow(app, page)).toBe(true)
    await app.evaluate(({ BrowserWindow }) => {
      const win = BrowserWindow.getAllWindows()[0]
      const b = win.getBounds()
      win.setBounds({ ...b, width: 760 })
    })
    await expectRulesFitTheWindow(app, page)
    await app.evaluate(({ BrowserWindow }) => {
      const win = BrowserWindow.getAllWindows()[0]
      const b = win.getBounds()
      win.setBounds({ ...b, width: 1280 })
    })
    // at 200% the page is wider than the window: the text scrolls sideways, the rules do not grow
    await setZoom(page, '200%')
    expect(await expectRulesFitTheWindow(app, page)).toBe(false)
    await setZoom(page, '100%')
    await clickViewItem(app, 'Print Layout')
    await expect.poll(() => printLayout(page), { timeout: 30_000 }).toEqual(before)

    // the page paints its own paper in Draft, in the main canvas and in the split pane
    const fill = (selector: string) =>
      page.locator(selector).evaluate((el) => getComputedStyle(el).backgroundColor)
    const paper = await fill('.editor-scroll .doc-page')
    await clickViewItem(app, 'Draft')
    await expect(canvas).toHaveClass(/\bview-draft\b/)
    expect(await fill('.editor-scroll .doc-page')).toBe(paper)
    await ribbonButton(page, 'Split').scrollIntoViewIfNeeded()
    await ribbonButton(page, 'Split').click()
    await expect(page.locator('.split-pane .doc-zoom')).toHaveClass(/\bview-draft\b/)
    expect(await fill('.split-pane .doc-page')).toBe(paper)
    await ribbonButton(page, 'Split').click()
    await expect(page.locator('.split-pane')).toHaveCount(0)

    // exported from Draft, the PDF is the print layout's pages with their headers, and Draft stays on
    const pdfPath = join(dir, 'report.pdf')
    await app.evaluate(({ dialog, shell }, out) => {
      dialog.showSaveDialog = async () => ({ canceled: false, filePath: out })
      shell.showItemInFolder = () => {}
    }, pdfPath)
    expect(
      await page.evaluate(() =>
        (window as unknown as { __exportPdf: () => Promise<boolean> }).__exportPdf(),
      ),
    ).toBe(true)
    const pdf = await pdfPages(pdfPath)
    expect(`of ${pdf.length}`).toBe(before.pages)
    for (const text of pdf) expect(text).toContain(DRAFT_HEADER)
    await expect(canvas).toHaveClass(/\bview-draft\b/)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test("a PDF exported from Draft has the print layout's pages, columns included", async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-draft-pdf-')))
  const docPath = join(dir, 'columns.docx')
  const long =
    'Lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor. '.repeat(6)
  const section = (cols: string) =>
    '<w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
    `${cols}</w:sectPr>`
  // a wrapped picture, then a two-column section: as one plain column it would take fewer pages
  writeFileSync(
    docPath,
    await docxWith(
      `<w:p>${pictureRun(1, 'left', 0, 200, 150)}<w:r><w:t>${long}</w:t></w:r></w:p>` +
        Array.from({ length: 12 }, () => para(long)).join('') +
        `<w:p><w:pPr>${section('')}</w:pPr></w:p>` +
        Array.from({ length: 40 }, (_, i) => para(`Column text ${i} ${long}`)).join('') +
        section('<w:cols w:num="2" w:space="720"/>'),
    ),
  )
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const canvas = page.locator('.doc-zoom')
    const pageCount = async () =>
      Number((await page.locator('.status-page').textContent())?.match(/of (\d+)/)?.[1] ?? 0)
    // the count settles once the columns are laid out
    let printed = 0
    await expect
      .poll(
        async () => {
          const now = await pageCount()
          const settled = now > 1 && now === printed
          printed = now
          return settled
        },
        { timeout: 60_000, intervals: [1500] },
      )
      .toBe(true)

    await clickViewItem(app, 'Draft')
    await expect(canvas).toHaveClass(/\bview-draft\b/)
    // this document's Draft flow is not the print layout's: no rules are drawn over it
    await expect(page.locator('.draft-break')).toHaveCount(0)
    await expect(page.locator('.doc-page .doc-inline-img').first()).toBeHidden()

    const pdfPath = join(dir, 'columns.pdf')
    await app.evaluate(({ dialog, shell }, out) => {
      dialog.showSaveDialog = async () => ({ canceled: false, filePath: out })
      shell.showItemInFolder = () => {}
    }, pdfPath)
    expect(
      await page.evaluate(() =>
        (window as unknown as { __exportPdf: () => Promise<boolean> }).__exportPdf(),
      ),
    ).toBe(true)
    expect(await pdfPages(pdfPath)).toHaveLength(printed)
    await expect(canvas).toHaveClass(/\bview-draft\b/)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

interface Tile {
  title: string
  x: number
  y: number
  width: number
  height: number
}

const windowTiles = (app: ElectronApplication) =>
  app.evaluate(({ BrowserWindow, screen }) => {
    const wins = BrowserWindow.getAllWindows().filter((w) => w.isVisible() && !w.isMinimized())
    return {
      area: screen.getDisplayMatching(wins[0].getBounds()).workArea,
      tiles: wins.map((w) => ({ title: w.getTitle(), ...w.getBounds() })),
    }
  })

const overlaps = (a: Tile, b: Tile) =>
  a.x < b.x + b.width && b.x < a.x + a.width && a.y < b.y + b.height && b.y < a.y + a.height

test('Arrange All gives every open document a window and tiles them over the screen', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-arrange-')))
  const docPath = join(dir, 'report.docx')
  writeFileSync(docPath, await draftDocx())
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const first = await waitForPageWithUrl(app, '://docs/')
    const opening = first.locator('.doc-page p', { hasText: 'Opening paragraph' })
    await expect(opening).toBeVisible({ timeout: 30_000 })
    // an unsaved edit has to survive the move into its own window
    await opening.click()
    await first.keyboard.press('End')
    await first.keyboard.type(' unsaved words')

    // a second document, in a second tab of the same window
    await expect
      .poll(
        async () => {
          for (const candidate of app.windows()) {
            const opened = await candidate
              .evaluate(async () => {
                const api = (window as never as { aiOffice?: { newDoc(): Promise<void> } }).aiOffice
                if (!api) return false
                await api.newDoc()
                return true
              })
              .catch(() => false)
            if (opened) return true
          }
          return false
        },
        { timeout: 30_000 },
      )
      .toBe(true)
    const docsPages = () => app.windows().filter((p) => p.url().includes('://docs/'))
    await expect.poll(() => docsPages().length, { timeout: 30_000 }).toBe(2)
    const second = docsPages().find((p) => p !== first)!
    await expect(second.locator('.status-page')).toContainText(/\d/, { timeout: 30_000 })
    expect((await windowTiles(app)).tiles).toHaveLength(1)

    await openViewTab(second)
    const arrange = ribbonButton(second, 'Arrange All')
    await arrange.scrollIntoViewIfNeeded()
    await expect(arrange).toBeEnabled()
    await arrange.click()

    // two windows, one per document, neither over the other, together filling the work area
    await expect
      .poll(async () => (await windowTiles(app)).tiles.length, { timeout: 30_000 })
      .toBe(2)
    await expect
      .poll(
        async () => {
          const { area, tiles } = await windowTiles(app)
          const covered = tiles.reduce((sum, t) => sum + t.width * t.height, 0)
          return (
            !overlaps(tiles[0], tiles[1]) &&
            tiles.every(
              (t) =>
                t.x >= area.x - 1 &&
                t.y >= area.y - 1 &&
                t.x + t.width <= area.x + area.width + 1 &&
                t.y + t.height <= area.y + area.height + 1,
            ) &&
            covered >= area.width * area.height * 0.9
          )
        },
        { timeout: 30_000 },
      )
      .toBe(true)
    const { tiles } = await windowTiles(app)
    expect(tiles.map((t) => t.title).sort()).toContain('report.docx')
    // stacked top to bottom at full width
    expect(Math.abs(tiles[0].width - tiles[1].width)).toBeLessThanOrEqual(2)
    expect(tiles[0].y).not.toBe(tiles[1].y)

    // the document that moved out is the same live document, unsaved words and all
    await expect(first.locator('.doc-page p', { hasText: 'unsaved words' })).toBeVisible()
    await first.locator('.doc-page p', { hasText: 'unsaved words' }).click()
    await first.keyboard.press('End')
    await first.keyboard.type(' and more')
    await expect(first.locator('.doc-page p', { hasText: 'unsaved words and more' })).toBeVisible()
    await first.screenshot({ path: test.info().outputPath('arranged-first.png') })

    // asked again from the other window, the same two windows stay tiled
    await openViewTab(first)
    await ribbonButton(first, 'Arrange All').scrollIntoViewIfNeeded()
    await ribbonButton(first, 'Arrange All').click()
    await expect.poll(async () => (await windowTiles(app)).tiles.length).toBe(2)
    const again = (await windowTiles(app)).tiles
    expect(overlaps(again[0], again[1])).toBe(false)
  } finally {
    // the moved document holds unsaved words: close its window without the prompt
    await launched.app
      .evaluate(({ BrowserWindow }) => {
        for (const w of BrowserWindow.getAllWindows()) w.destroy()
      })
      .catch(() => {})
    await closeAndSaveVideo(launched, VIDEO_DIR).catch(() => {})
    rmSync(dir, { recursive: true, force: true })
  }
})
