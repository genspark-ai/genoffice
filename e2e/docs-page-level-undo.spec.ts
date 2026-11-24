/**
 * Settings that live beside the editor document are steps of the undo history,
 * in order with text edits: page colour, watermark, page number format, page
 * setup made by the AI tools, the Footnote dialog's Insert, a list style's
 * levels, the Typography options, and a style set picked and undone before it
 * has landed.
 */
import { test, expect, type ElectronApplication, type Locator, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { createServer, type IncomingMessage, type Server } from 'node:http'
import type { AddressInfo } from 'node:net'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { defaultAiSettings, resolveCustomEndpoints } from '@genoffice/ai-provider'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-page-level-undo'
const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/></w:sectPr>'

const menuItem = (app: ElectronApplication, path: string[]) =>
  app.evaluate(({ Menu }, path) => {
    let items = Menu.getApplicationMenu()?.items
    let item: Electron.MenuItem | undefined
    for (const label of path) {
      item = items?.find((candidate) => candidate.label === label)
      items = item?.submenu?.items
    }
    return item ? { enabled: item.enabled } : null
  }, path)

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

/** saves and waits until the document on screen is the one that was written */
async function saveAndSettle(app: ElectronApplication, page: Page) {
  await saveFromMenu(app)
  await expect
    .poll(
      () =>
        page.evaluate(() =>
          (window as never as { __aidocs: { edited(): boolean } }).__aidocs.edited(),
        ),
      { timeout: 30_000 },
    )
    .toBe(false)
}

const undo = (app: ElectronApplication) => clickMenu(app, ['Edit', 'Undo'])
const redo = (app: ElectronApplication) => clickMenu(app, ['Edit', 'Redo'])

const part = async (docPath: string, name: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file(name)?.async('string') ?? ''

const parts = async (docPath: string, names: RegExp) => {
  const zip = await JSZip.loadAsync(readFileSync(docPath))
  const found = Object.keys(zip.files).filter((name) => names.test(name))
  return (await Promise.all(found.map((name) => zip.file(name)!.async('string')))).join('')
}

async function caretIn(paragraph: Locator) {
  await paragraph.click()
  await paragraph.page().keyboard.press('End')
}

async function openDoc(docPath: string, shows: string, userDataDir?: string) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
    ...(userDataDir ? { userDataDir } : {}),
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  const paragraph = (text: string) => page.locator('.doc-page p', { hasText: text }).first()
  await expect(page.locator('.doc-page').first()).toContainText(shows, { timeout: 30_000 })
  return { launched, app: launched.app, page, paragraph }
}

async function designPick(page: Page, control: string) {
  const button = page.locator('[data-ribbon-body] button.rb-big', { hasText: control }).first()
  // clicking the tab that is already open folds the ribbon away
  if (!(await button.isVisible()))
    await page.locator('.ribbon-tab', { hasText: /^Design$/ }).click()
  await button.click()
}

const pageColor = (page: Page) =>
  page
    .locator('.doc-page')
    .first()
    .evaluate((el) => getComputedStyle(el).backgroundColor)

/** the text watermarks drawn behind the first page */
const watermarks = (page: Page) =>
  page.locator('.page-hf-float-img[data-hf-behind="true"] svg text').allTextContents()

async function setWatermark(page: Page, text: string) {
  await designPick(page, 'Watermark')
  await page.locator('[data-rb-panel] button', { hasText: 'Custom Watermark' }).click()
  const dialog = page.locator('.watermark-dialog')
  await dialog.locator('.wm-mode', { hasText: 'Text' }).locator('input').check()
  await dialog.locator('.wm-row input[type=text]').first().fill(text)
  await dialog.locator('.btn-primary').click()
  await expect(dialog).toBeHidden()
}

/** the Page Number Format dialog, opened from Insert > Page Numbers… */
async function openPageNumberFormat(app: ElectronApplication, page: Page) {
  await clickMenu(app, ['Insert', 'Page Numbers…'])
  await page
    .locator('.page-numbers-dialog')
    .getByRole('button', { name: 'Format Page Numbers…' })
    .click()
  const format = page.locator('.pgnum-dialog')
  await expect(format).toBeVisible()
  return format
}

async function closePageNumbers(page: Page) {
  const format = page.locator('.pgnum-dialog')
  if (await format.isVisible()) await format.getByRole('button', { name: 'Cancel' }).click()
  const numbers = page.locator('.page-numbers-dialog')
  if (await numbers.isVisible()) await numbers.getByRole('button', { name: 'Cancel' }).click()
  await expect(numbers).toBeHidden()
}

const pageNumberFormat = async (app: ElectronApplication, page: Page) => {
  const format = await openPageNumberFormat(app, page)
  const shown = (await format.locator('.gs-dd-btn').textContent())?.trim()
  await closePageNumbers(page)
  return shown
}

test('page colour, watermark and page number format undo and redo in order with typing, before a save and after one', async () => {
  test.setTimeout(360_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-page-level-undo-')))
  const docPath = join(dir, 'page.docx')
  writeFileSync(docPath, await docxWith(['apple', 'banana'].map(para).join('') + SECT_PR))
  const { launched, app, page, paragraph } = await openDoc(docPath, 'apple')
  const saved = async () => {
    const xml = await part(docPath, 'word/document.xml')
    return {
      color: /<w:background\b[^>]* w:color="([^"]*)"/.exec(xml)?.[1] ?? null,
      numbers: /<w:pgNumType\b[^>]* w:fmt="([^"]*)"/.exec(xml)?.[1] ?? null,
      watermark: /string="([^"]*)"/.exec(await parts(docPath, /^word\/header\d*\.xml$/))?.[1],
    }
  }
  try {
    const white = await pageColor(page)
    const numbered = await pageNumberFormat(app, page)
    expect((await menuItem(app, ['Edit', 'Undo']))?.enabled).toBe(false)

    await caretIn(paragraph('apple'))
    await page.keyboard.type(' one')
    await designPick(page, 'Page Color')
    await page.locator('.docs-color-pop .gcp-swatch[title="Orange"]').first().click()
    await expect.poll(() => pageColor(page)).not.toBe(white)
    const orange = await pageColor(page)
    await caretIn(paragraph('apple one'))
    await page.keyboard.type(' two')
    await setWatermark(page, 'INTERNAL')
    await expect.poll(() => watermarks(page)).toEqual(['INTERNAL'])
    const format = await openPageNumberFormat(app, page)
    await format.locator('.gs-dd-btn').click()
    await format.getByRole('option', { name: 'I, II, III, …', exact: true }).click()
    await format.locator('button.btn-primary').click()
    await expect(format).toBeHidden()
    await closePageNumbers(page)
    await page.screenshot({ path: test.info().outputPath('page-level-changed.png') })

    // each is its own step, the newest first, with the typing between them in its place
    await undo(app)
    expect(await pageNumberFormat(app, page)).toBe(numbered)
    expect(await watermarks(page)).toEqual(['INTERNAL'])
    await undo(app)
    await expect.poll(() => watermarks(page)).toEqual([])
    await expect(paragraph('apple one two')).toBeVisible()
    await undo(app)
    await expect(paragraph('apple one two')).toHaveCount(0)
    expect(await pageColor(page)).toBe(orange)
    await undo(app)
    await expect.poll(() => pageColor(page)).toBe(white)
    await expect(paragraph('apple one')).toBeVisible()
    await undo(app)
    await expect(paragraph('apple one')).toHaveCount(0)
    await expect.poll(async () => (await menuItem(app, ['Edit', 'Undo']))?.enabled).toBe(false)

    for (let step = 0; step < 5; step++) await redo(app)
    await expect.poll(() => pageColor(page)).toBe(orange)
    await expect.poll(() => watermarks(page)).toEqual(['INTERNAL'])
    await expect(paragraph('apple one two')).toBeVisible()
    expect(await pageNumberFormat(app, page)).toBe('I, II, III, …')
    await saveAndSettle(app, page)
    await expect
      .poll(saved, { timeout: 30_000 })
      .toEqual({ color: 'E97132', numbers: 'upperRoman', watermark: 'INTERNAL' })

    // the file now holds all three: undone after the save, they are written out of it again
    await undo(app)
    await undo(app)
    await expect.poll(() => watermarks(page)).toEqual([])
    await undo(app)
    await undo(app)
    await expect.poll(() => pageColor(page)).toBe(white)
    await saveAndSettle(app, page)
    await expect
      .poll(saved, { timeout: 30_000 })
      .toEqual({ color: null, numbers: null, watermark: undefined })
    expect(await part(docPath, 'word/document.xml')).toContain('apple one</w:t>')
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test("replacing the file's own watermark undoes to it without touching its part, and after a save writes it again", async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-page-level-undo-')))
  const docPath = join(dir, 'marked.docx')
  writeFileSync(docPath, await docxWith(['apple', 'banana'].map(para).join('') + SECT_PR))
  const headers = () => parts(docPath, /^word\/header\d*\.xml$/)
  const first = await openDoc(docPath, 'apple')
  try {
    await first.paragraph('apple').click()
    await setWatermark(first.page, 'INTERNAL')
    await saveAndSettle(first.app, first.page)
    await expect.poll(headers, { timeout: 30_000 }).toContain('string="INTERNAL"')
  } finally {
    await closeAndSaveVideo(first.launched, VIDEO_DIR)
  }
  const own = await headers()

  const { launched, app, page, paragraph } = await openDoc(docPath, 'apple')
  try {
    await expect.poll(() => watermarks(page)).toEqual(['INTERNAL'])
    await caretIn(paragraph('banana'))
    await page.keyboard.type(' split')
    await setWatermark(page, 'DRAFT')
    await expect.poll(() => watermarks(page)).toEqual(['DRAFT'])
    await undo(app)
    await expect.poll(() => watermarks(page)).toEqual(['INTERNAL'])
    await saveAndSettle(app, page)
    await expect
      .poll(() => part(docPath, 'word/document.xml'), { timeout: 30_000 })
      .toContain('banana split')
    // the undone change left the watermark to the file: its part is the one that was read
    expect(await headers()).toBe(own)

    // a save came in between: the file holds DRAFT, and the undo says INTERNAL in full
    await setWatermark(page, 'DRAFT')
    await saveAndSettle(app, page)
    await expect.poll(headers, { timeout: 30_000 }).toContain('string="DRAFT"')
    await undo(app)
    await expect.poll(() => watermarks(page)).toEqual(['INTERNAL'])
    await saveAndSettle(app, page)
    await expect.poll(headers, { timeout: 30_000 }).toContain('string="INTERNAL"')
    expect(await headers()).not.toContain('string="DRAFT"')
    await redo(app)
    await expect.poll(() => watermarks(page)).toEqual(['DRAFT'])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

interface ChatRequest {
  messages: Array<{ role: string }>
}

async function readBody(req: IncomingMessage): Promise<string> {
  const chunks: Buffer[] = []
  for await (const chunk of req) chunks.push(chunk as Buffer)
  return Buffer.concat(chunks).toString('utf8')
}

/** an OpenAI-compatible model that answers each prompt with one scripted tool call, then with a word */
function startModelStub(calls: Array<{ name: string; input: object }>): Promise<Server> {
  let turn = 0
  const server = createServer(async (req, res) => {
    if (req.method !== 'POST') {
      res.writeHead(404).end()
      return
    }
    const body = JSON.parse(await readBody(req)) as ChatRequest
    const answered = body.messages.at(-1)?.role === 'tool'
    const call = answered ? undefined : calls[turn++]
    const message = call
      ? {
          role: 'assistant',
          content: null,
          tool_calls: [
            {
              id: `call_${turn}`,
              type: 'function',
              function: { name: call.name, arguments: JSON.stringify(call.input) },
            },
          ],
        }
      : { role: 'assistant', content: 'Done.' }
    res.writeHead(200, { 'content-type': 'application/json' })
    res.end(JSON.stringify({ choices: [{ message, finish_reason: call ? 'tool_calls' : 'stop' }] }))
  })
  return new Promise((resolve) => server.listen(0, '127.0.0.1', () => resolve(server)))
}

/** sends a prompt in the AI panel and waits for the run to end */
async function ask(page: Page, prompt: string) {
  const rail = page.locator('.ai-dock .ai-rail')
  const box = page.locator('.ai-dock textarea')
  await expect(rail.or(box).first()).toBeVisible({ timeout: 60_000 })
  if (await rail.isVisible()) await rail.click()
  await box.fill(prompt)
  const stop = page.locator('button.ai-send-btn.ai-stop-btn')
  await page.locator('button.ai-send-btn:not(.ai-stop-btn)').click()
  await stop.waitFor({ timeout: 15_000 }).catch(() => {})
  await expect(stop).toHaveCount(0, { timeout: 120_000 })
}

test('page setup and a watermark made by the AI tools are undo steps in order with typing', async () => {
  test.setTimeout(300_000)
  const server = await startModelStub([
    { name: 'set_page_setup', input: { orientation: 'landscape' } },
    { name: 'set_watermark', input: { text: 'DRAFT' } },
  ])
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-page-level-undo-')))
  const userDataDir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-')))
  const settings = defaultAiSettings()
  settings.provider = 'custom'
  settings.providers.custom = {
    apiKey: 'stub',
    model: 'stub-model',
    baseUrl: `http://127.0.0.1:${(server.address() as AddressInfo).port}/v1`,
  }
  writeFileSync(
    join(userDataDir, 'ai-settings.json'),
    JSON.stringify(resolveCustomEndpoints(settings)),
  )
  const docPath = join(dir, 'ai.docx')
  writeFileSync(docPath, await docxWith(['apple', 'banana'].map(para).join('') + SECT_PR))
  const { launched, app, page, paragraph } = await openDoc(docPath, 'apple', userDataDir)
  const landscape = () =>
    page
      .locator('.doc-page')
      .first()
      .evaluate((el) => {
        const sheet = el.getBoundingClientRect()
        return sheet.width > sheet.height
      })
  try {
    await caretIn(paragraph('apple'))
    await page.keyboard.type(' one')
    await ask(page, 'Turn the page sideways.')
    await expect.poll(landscape).toBe(true)
    await caretIn(paragraph('apple one'))
    await page.keyboard.type(' two')
    await ask(page, 'Mark it as a draft.')
    await expect.poll(() => watermarks(page)).toEqual(['DRAFT'])
    await page.screenshot({ path: test.info().outputPath('ai-page-setup.png') })

    await undo(app)
    await expect.poll(() => watermarks(page)).toEqual([])
    expect(await landscape()).toBe(true)
    await expect(paragraph('apple one two')).toBeVisible()
    await undo(app)
    await expect(paragraph('apple one two')).toHaveCount(0)
    expect(await landscape()).toBe(true)
    await undo(app)
    await expect.poll(landscape).toBe(false)
    await expect(paragraph('apple one')).toBeVisible()
    await redo(app)
    await expect.poll(landscape).toBe(true)
    await saveAndSettle(app, page)
    await expect
      .poll(
        async () =>
          /<w:pgSz\b[^>]* w:orient="([^"]*)"/.exec(await part(docPath, 'word/document.xml'))?.[1],
        { timeout: 30_000 },
      )
      .toBe('landscape')
    expect(await parts(docPath, /^word\/header\d*\.xml$/)).not.toContain('DRAFT')
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    server.close()
    rmSync(dir, { recursive: true, force: true })
    rmSync(userDataDir, { recursive: true, force: true })
  }
})

const noteMarks = (page: Page) =>
  page.locator('.doc-page sup.doc-note-ref[data-note-kind="footnote"]').allTextContents()

test('settings changed in the Footnote dialog and the note its Insert puts in are one undo step', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-page-level-undo-')))
  const docPath = join(dir, 'notes.docx')
  writeFileSync(docPath, await docxWith(['apple', 'banana'].map(para).join('') + SECT_PR))
  const { launched, app, page, paragraph } = await openDoc(docPath, 'apple')
  const dialog = page.locator('.footnote-dialog')
  const format = dialog.locator('select[name="note-format"]')
  const shownFormat = async () => {
    await clickMenu(app, ['Insert', 'Footnote…'])
    const value = await format.inputValue()
    await dialog.getByRole('button', { name: 'Cancel' }).click()
    await expect(dialog).toBeHidden()
    return value
  }
  try {
    await caretIn(paragraph('banana'))
    await page.keyboard.type(' split')
    await caretIn(paragraph('apple'))
    await clickMenu(app, ['Insert', 'Footnote…'])
    await format.selectOption('lowerLetter')
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    const prompt = page.locator('.modal textarea.prompt-textarea')
    await prompt.fill('a note')
    await page.keyboard.press('ControlOrMeta+Enter')
    await expect(prompt).toBeHidden()
    await expect.poll(() => noteMarks(page)).toEqual(['a'])
    await page.screenshot({ path: test.info().outputPath('footnote-insert.png') })

    // one Undo takes the mark and the number format; the typing before it stays
    await undo(app)
    await expect.poll(() => noteMarks(page)).toEqual([])
    expect(await shownFormat()).toBe('decimal')
    await expect(paragraph('banana split')).toBeVisible()
    await redo(app)
    await expect.poll(() => noteMarks(page)).toEqual(['a'])
    expect(await shownFormat()).toBe('lowerLetter')
    await undo(app)
    await undo(app)
    await expect(paragraph('banana split')).toHaveCount(0)
    await expect.poll(async () => (await menuItem(app, ['Edit', 'Undo']))?.enabled).toBe(false)
    await redo(app)
    await redo(app)
    await saveAndSettle(app, page)
    await expect
      .poll(() => part(docPath, 'word/settings.xml'), { timeout: 30_000 })
      .toMatch(/<w:footnotePr>[\s\S]*<w:numFmt w:val="lowerLetter"\/>[\s\S]*<\/w:footnotePr>/)
    expect((await part(docPath, 'word/document.xml')).match(/<w:footnoteReference /g)).toHaveLength(
      1,
    )
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

type EditorWindow = {
  __aidocs: {
    editor: {
      state: {
        doc: {
          descendants(
            f: (
              n: { textContent: string; isTextblock: boolean; nodeSize: number },
              pos: number,
            ) => boolean,
          ): void
        }
      }
      storage: { cjkPunctShrink: { enabled: boolean; kana: boolean; strictKinsoku: boolean } }
      commands: { undo(): boolean; redo(): boolean }
      chain(): { focus(): { setTextSelection(r: object | number): { run(): boolean } } }
    }
  }
}

/** selects the paragraphs from the one starting with `first` through `last`, or without `last` puts the caret at the end of the first */
const select = (page: Page, first: string, last: string | null) =>
  page.evaluate(
    ([first, last]) => {
      const { editor } = (window as never as EditorWindow).__aidocs
      let from = -1
      let to = -1
      editor.state.doc.descendants((node, pos) => {
        if (!node.isTextblock) return true
        if (node.textContent.startsWith(first)) {
          from = last !== null ? pos + 1 : pos + node.nodeSize - 1
          if (last === null) to = from
        }
        if (last !== null && node.textContent === last) to = pos + 1 + last.length
        return false
      })
      if (from < 0 || to < 0) throw new Error(`no paragraph reads ${first}`)
      editor.chain().focus().setTextSelection({ from, to }).run()
    },
    [first, last] as const,
  )

async function typeAt(page: Page, first: string, text: string) {
  await select(page, first, null)
  await expect
    .poll(() => page.evaluate(() => !!document.activeElement?.closest('.ProseMirror')))
    .toBe(true)
  await page.keyboard.type(text)
}

const markers = (page: Page) =>
  page
    .locator('.doc-page .doc-li')
    .evaluateAll((items) => items.map((el) => el.getAttribute('data-marker')))

async function openListStyles(app: ElectronApplication, page: Page) {
  await clickMenu(app, ['Format', 'Bullets and Numbering…'])
  const dialog = page.locator('.bn-dialog')
  await dialog.getByRole('tab', { name: 'List Styles' }).click()
  return dialog
}

/** List Styles > Modify… on the style at the caret: another number format for its first level */
async function modifyListStyle(app: ElectronApplication, page: Page, numberFormat: string) {
  const dialog = await openListStyles(app, page)
  await dialog.locator('button[name="list-style-modify"]').click()
  const define = page.locator('.modal.list-define-modal')
  await define.getByLabel('Number format').fill(numberFormat)
  await define.locator('button.primary').click()
  await expect(define).toBeHidden()
  await dialog.getByRole('button', { name: 'Cancel' }).click()
  await expect(dialog).toBeHidden()
}

test('List Styles > Modify… is one undo step, in order with typing, before a save and after one', async () => {
  test.setTimeout(360_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-page-level-undo-')))
  const docPath = join(dir, 'list.docx')
  writeFileSync(docPath, await docxWith(['alpha', 'beta', 'gamma', 'delta'].map(para).join('')))
  const { launched, app, page } = await openDoc(docPath, 'alpha')
  const levelTexts = async () =>
    Array.from(
      (await part(docPath, 'word/numbering.xml')).matchAll(/<w:lvlText w:val="([^"]*)"\/>/g),
      (m) => m[1],
    )
  try {
    await select(page, 'alpha', 'gamma')
    const dialog = await openListStyles(app, page)
    await dialog.locator('.bn-styles-list [role=option]').nth(2).click()
    await dialog.locator('button.primary').click()
    await expect.poll(() => markers(page)).toEqual(['1)', '2)', '3)'])
    // the style and its list are in the file before the style is modified
    await saveAndSettle(app, page)
    await expect
      .poll(() => part(docPath, 'word/numbering.xml'), { timeout: 30_000 })
      .toContain('<w:styleLink w:val="1ai"/>')
    const numbering = await part(docPath, 'word/numbering.xml')

    await typeAt(page, 'delta', ' one')
    await select(page, 'beta', null)
    await modifyListStyle(app, page, 'Item %1:')
    await expect.poll(() => markers(page)).toEqual(['Item 1:', 'Item 2:', 'Item 3:'])
    await typeAt(page, 'delta', ' two')
    await select(page, 'beta', null)
    await modifyListStyle(app, page, 'Step %1.')
    await expect.poll(() => markers(page)).toEqual(['Step 1.', 'Step 2.', 'Step 3.'])
    await page.screenshot({ path: test.info().outputPath('list-style-modified.png') })

    await undo(app)
    await expect.poll(() => markers(page)).toEqual(['Item 1:', 'Item 2:', 'Item 3:'])
    await expect(page.locator('.doc-page p', { hasText: 'delta one two' })).toBeVisible()
    await undo(app)
    await expect(page.locator('.doc-page p', { hasText: 'delta one two' })).toHaveCount(0)
    expect(await markers(page)).toEqual(['Item 1:', 'Item 2:', 'Item 3:'])
    await undo(app)
    await expect.poll(() => markers(page)).toEqual(['1)', '2)', '3)'])
    // undone before any save, the style is left to the file as it was read
    await saveAndSettle(app, page)
    await expect
      .poll(() => part(docPath, 'word/document.xml'), { timeout: 30_000 })
      .toContain('delta one</w:t>')
    expect(await part(docPath, 'word/numbering.xml')).toBe(numbering)

    await redo(app)
    await redo(app)
    await redo(app)
    await expect.poll(() => markers(page)).toEqual(['Step 1.', 'Step 2.', 'Step 3.'])
    await saveAndSettle(app, page)
    await expect.poll(levelTexts, { timeout: 30_000 }).toContain('Step %1.')

    // the file holds the second modification: undone after the save, the first is written again
    await undo(app)
    await expect.poll(() => markers(page)).toEqual(['Item 1:', 'Item 2:', 'Item 3:'])
    await saveAndSettle(app, page)
    await expect.poll(levelTexts, { timeout: 30_000 }).toContain('Item %1:')
    expect(await levelTexts()).not.toContain('Step %1.')
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

async function openTypography(app: ElectronApplication, page: Page) {
  await clickMenu(app, ['Format', 'Paragraph…'])
  const dialog = page.locator('.modal', { hasText: 'Paragraph' }).first()
  await dialog.getByRole('tab', { name: 'Asian Typography' }).click()
  await dialog.getByRole('button', { name: 'Options…' }).click()
  const typography = page.getByRole('dialog', { name: 'Typography' })
  await expect(typography).toBeVisible()
  return { dialog, typography }
}

async function closeTypography(dialog: Locator, typography: Locator, button: 'OK' | 'Cancel') {
  await typography.getByRole('button', { name: button, exact: true }).click()
  await expect(typography).toBeHidden()
  await dialog.getByRole('button', { name: button, exact: true }).click()
  await expect(dialog).toBeHidden()
}

test('the Typography options are one undo step, in order with typing', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-page-level-undo-')))
  const docPath = join(dir, 'typography.docx')
  writeFileSync(docPath, await docxWith(['apple', 'banana'].map(para).join('') + SECT_PR))
  const { launched, app, page, paragraph } = await openDoc(docPath, 'apple')
  // the switches the page compresses and breaks lines by
  const drawn = () =>
    page.evaluate(() => {
      const { enabled, kana } = (window as never as EditorWindow).__aidocs.editor.storage
        .cjkPunctShrink
      return { enabled, kana }
    })
  const chosen = async () => {
    const { dialog, typography } = await openTypography(app, page)
    const kana = await typography.getByLabel('Compress punctuation and Japanese kana').isChecked()
    const none = await typography.getByLabel('Do not compress').isChecked()
    await closeTypography(dialog, typography, 'Cancel')
    return kana ? 'kana' : none ? 'none' : 'punct'
  }
  const control = async () =>
    /<w:characterSpacingControl w:val="([^"]*)"/.exec(await part(docPath, 'word/settings.xml'))?.[1]
  try {
    const started = await chosen()
    const drawnAtStart = await drawn()
    await caretIn(paragraph('apple'))
    await page.keyboard.type(' one')
    // a dialog opened by hand comes later than the typing, which is then a step of its own
    await page.waitForTimeout(600)
    const { dialog, typography } = await openTypography(app, page)
    await typography.getByLabel('Compress punctuation and Japanese kana').check()
    await page.screenshot({ path: test.info().outputPath('typography-options.png') })
    await closeTypography(dialog, typography, 'OK')
    await expect.poll(drawn).toEqual({ enabled: true, kana: true })
    await caretIn(paragraph('banana'))
    await page.keyboard.type(' two')

    await undo(app)
    await expect(paragraph('banana two')).toHaveCount(0)
    expect(await drawn()).toEqual({ enabled: true, kana: true })
    // the option and the OK of the Paragraph dialog it was set from are one step
    await undo(app)
    await expect.poll(drawn).toEqual(drawnAtStart)
    expect(await chosen()).toBe(started)
    await expect(paragraph('apple one')).toBeVisible()
    await redo(app)
    await expect.poll(drawn).toEqual({ enabled: true, kana: true })
    expect(await chosen()).toBe('kana')
    await saveAndSettle(app, page)
    await expect.poll(control, { timeout: 30_000 }).toBe('compressPunctuationAndJapaneseKana')

    // after the save the file holds the option: the undo writes the earlier one over it
    await undo(app)
    await expect.poll(drawn).toEqual(drawnAtStart)
    await saveAndSettle(app, page)
    await expect.poll(control, { timeout: 30_000 }).not.toBe('compressPunctuationAndJapaneseKana')
    expect(await chosen()).toBe(started)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

const styled = (style: string, text: string) =>
  `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`

test('a style set undone before it has landed stays off the page, and Redo puts it there', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-page-level-undo-')))
  const docPath = join(dir, 'report.docx')
  writeFileSync(
    docPath,
    await docxWith(styled('Heading1', 'Overview') + ['apple', 'banana'].map(para).join('')),
  )
  const { launched, page, paragraph } = await openDoc(docPath, 'apple')
  const heading = page.locator('.doc-page :is(h1, p)', { hasText: 'Overview' }).first()
  const look = () =>
    heading.evaluate((el) => {
      const style = getComputedStyle(el)
      return `${style.textTransform} ${style.color}`
    })
  const history = (step: 'undo' | 'redo') =>
    page.evaluate(
      (step) => (window as never as EditorWindow).__aidocs.editor.commands[step](),
      step,
    )
  try {
    await paragraph('apple').click()
    await page.locator('.ribbon-tab', { hasText: /^Design$/ }).click()
    const strip = page.locator('.style-set-strip')
    await expect(strip.getByTestId('style-set-elegant')).toBeVisible()
    const plain = await look()
    await page.getByTestId('style-set-more').click()
    await page.locator('.style-set-menu').getByTestId('style-set-modern').click()
    await expect.poll(look).toContain('uppercase')
    const modern = await look()
    await history('undo')
    await expect.poll(look).toBe(plain)
    // the styles the undo left are read again before the next pick
    await page.waitForTimeout(500)

    // Modern lands at once; Elegant has to read the styles Modern left, and is undone before it has
    await page.getByTestId('style-set-more').click()
    await page
      .locator('.style-set-menu')
      .getByTestId('style-set-modern')
      .evaluate((card) => {
        const { editor } = (window as never as EditorWindow).__aidocs
        ;(card as HTMLElement).click()
        ;(
          card.ownerDocument.querySelector(
            '.style-set-strip [data-testid="style-set-elegant"]',
          ) as HTMLElement
        ).click()
        editor.commands.undo()
      })
    await page.waitForTimeout(1500)
    expect(await look()).toBe(modern)
    await history('redo')
    await expect.poll(look).not.toBe(modern)
    const elegant = await look()
    expect(elegant).not.toBe(plain)
    await page.screenshot({ path: test.info().outputPath('style-set-redone.png') })
    await history('undo')
    await expect.poll(look).toBe(modern)
    await history('undo')
    await expect.poll(look).toBe(plain)
    await history('redo')
    await history('redo')
    await expect.poll(look).toBe(elegant)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
