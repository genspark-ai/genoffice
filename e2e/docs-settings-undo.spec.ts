/**
 * The Footnote and Endnote dialog's settings and the Hyphenation dialog's
 * settings are steps of the undo history, in order with text edits, and a
 * save after an undo writes the earlier values.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu, withLanguage } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-settings-undo'
const PAGE =
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
const SECT_PR = `<w:sectPr>${PAGE}</w:sectPr>`
const sectionEnd = (text: string) =>
  `<w:p><w:pPr><w:sectPr>${PAGE}</w:sectPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`

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

const part = async (docPath: string, name: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file(name)?.async('string') ?? ''

const sectPrs = async (docPath: string) =>
  (await part(docPath, 'word/document.xml')).match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g) ?? []

const val = (xml: string, tag: string) =>
  new RegExp(`<w:${tag}\\b[^>]* w:val="([^"]*)"`).exec(xml)?.[1] ?? null

/** format, start and restart a w:footnotePr states, inside `xml` */
const footnotePr = (xml: string) => {
  const own = /<w:footnotePr>([\s\S]*?)<\/w:footnotePr>/.exec(xml)?.[1] ?? ''
  return [val(own, 'numFmt'), val(own, 'numStart'), val(own, 'numRestart')]
}

async function openDoc(docPath: string) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await expect(page.locator('.doc-page p').first()).toBeVisible({ timeout: 30_000 })
  return { launched, app: launched.app, page }
}

const paragraphOf = (page: Page, text: string) =>
  page.locator('.doc-page p', { hasText: text }).first()

/** the caret at the end of a paragraph; a click can lose to the focus a closing dialog hands back */
async function caretIn(page: Page, paragraph: string) {
  await expect(async () => {
    await paragraphOf(page, paragraph).click()
    await page.keyboard.press('End')
    const at = await page.evaluate(() => {
      const { editor } = (
        window as never as {
          __aidocs: {
            editor: { state: { selection: { $from: { parent: { textContent: string } } } } }
          }
        }
      ).__aidocs
      return editor.state.selection.$from.parent.textContent
    })
    expect(at).toContain(paragraph)
  }).toPass({ timeout: 15_000 })
}

async function typeIn(page: Page, paragraph: string, text: string) {
  await caretIn(page, paragraph)
  // a Home pressed while the page is still settling leaves the caret at the end
  await expect(async () => {
    await page.keyboard.press('Home')
    const offset = await page.evaluate(
      () =>
        (
          window as never as {
            __aidocs: { editor: { state: { selection: { $from: { parentOffset: number } } } } }
          }
        ).__aidocs.editor.state.selection.$from.parentOffset,
    )
    expect(offset).toBe(0)
  }).toPass({ timeout: 15_000 })
  await page.keyboard.type(text)
  await expect(paragraphOf(page, text.trim())).toBeVisible()
}

async function noteDialog(app: ElectronApplication, page: Page, paragraph: string) {
  await caretIn(page, paragraph)
  await clickMenu(app, ['Insert', 'Footnote…'])
  const dialog = page.locator('.footnote-dialog')
  await expect(dialog).toBeVisible()
  return dialog
}

async function insertNote(app: ElectronApplication, page: Page, paragraph: string, text: string) {
  const dialog = await noteDialog(app, page, paragraph)
  await dialog.locator('button.primary').click()
  await expect(dialog).toBeHidden()
  const prompt = page.locator('.modal textarea.prompt-textarea')
  await prompt.fill(text)
  await page.keyboard.press('ControlOrMeta+Enter')
  await expect(prompt).toBeHidden()
}

const marks = (page: Page) =>
  page.locator('.doc-page sup.doc-note-ref[data-note-kind="footnote"]').allTextContents()

const undo = (app: ElectronApplication) => clickMenu(app, ['Edit', 'Undo'])
const redo = (app: ElectronApplication) => clickMenu(app, ['Edit', 'Redo'])

test('note numbering for the whole document and for one section undoes in order with typing, and saves what is on the page', async () => {
  test.setTimeout(360_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-settings-undo-')))
  const docPath = join(dir, 'notes.docx')
  writeFileSync(
    docPath,
    await docxWith(
      // the plain paragraphs are typed in: text typed next to a note mark takes the mark's
      // run properties, and a save that reads them back differently starts a new history
      para('Alpha first') +
        para('Delta plain') +
        sectionEnd('Bravo second') +
        para('Charlie third') +
        para('Echo plain') +
        SECT_PR,
    ),
  )
  let session = await openDoc(docPath)
  try {
    let { app, page } = session
    await insertNote(app, page, 'Alpha first', 'note alpha')
    await insertNote(app, page, 'Bravo second', 'note bravo')
    await insertNote(app, page, 'Charlie third', 'note charlie')
    await expect.poll(() => marks(page)).toEqual(['1', '2', '3'])
    await saveFromMenu(app)
    await expect
      .poll(async () => (await part(docPath, 'word/document.xml')).match(/<w:footnoteReference /g))
      .toHaveLength(3)

    // a fresh history: the notes are the file's
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    session = await openDoc(docPath)
    ;({ app, page } = session)
    await expect.poll(() => marks(page)).toEqual(['1', '2', '3'])
    expect((await menuItem(app, ['Edit', 'Undo']))?.enabled).toBe(false)

    await typeIn(page, 'Delta plain', 'uno ')
    let dialog = await noteDialog(app, page, 'Alpha first')
    await dialog.locator('select[name="note-format"]').selectOption('lowerLetter')
    await dialog.locator('input[name="note-start"]').fill('3')
    await dialog.locator('button[name="apply"]').click()
    await expect(dialog).toBeHidden()
    await expect.poll(() => marks(page)).toEqual(['c', 'd', 'e'])

    await typeIn(page, 'Echo plain', 'duo ')
    dialog = await noteDialog(app, page, 'Charlie third')
    await dialog.locator('select[name="note-apply-to"]').selectOption('section')
    await dialog.locator('select[name="note-format"]').selectOption('upperRoman')
    await dialog.locator('input[name="note-start"]').fill('5')
    await dialog.locator('select[name="note-restart"]').selectOption('eachSect')
    await dialog.locator('button[name="apply"]').click()
    await expect(dialog).toBeHidden()
    await expect.poll(() => marks(page)).toEqual(['c', 'd', 'V'])
    await page.screenshot({ path: test.info().outputPath('notes-changed.png') })

    await saveFromMenu(app)
    await expect
      .poll(async () => footnotePr((await sectPrs(docPath))[1] ?? ''), { timeout: 30_000 })
      .toEqual(['upperRoman', '5', 'eachSect'])
    expect(footnotePr((await sectPrs(docPath))[0]).slice(0, 2)).toEqual(['lowerLetter', '3'])

    // This section, one step: the section counts with the document again, and the save says so
    await undo(app)
    await expect.poll(() => marks(page)).toEqual(['c', 'd', 'e'])
    await expect(paragraphOf(page, 'duo')).toBeVisible()
    await saveFromMenu(app)
    await expect
      .poll(async () => footnotePr((await sectPrs(docPath))[1] ?? ''), { timeout: 30_000 })
      .toEqual(['lowerLetter', '3', 'continuous'])
    expect(await part(docPath, 'word/document.xml')).not.toContain('upperRoman')
    dialog = await noteDialog(app, page, 'Charlie third')
    await expect(dialog.locator('select[name="note-format"]')).toHaveValue('lowerLetter')
    await expect(dialog.locator('input[name="note-start"]')).toHaveValue('3')
    await expect(dialog.locator('select[name="note-restart"]')).toHaveValue('continuous')
    await dialog.getByRole('button', { name: 'Cancel' }).click()
    await expect(dialog).toBeHidden()

    // the typing between the two settings, then the whole-document setting
    await undo(app)
    await expect(paragraphOf(page, 'duo')).toHaveCount(0)
    await expect.poll(() => marks(page)).toEqual(['c', 'd', 'e'])
    await undo(app)
    await expect.poll(() => marks(page)).toEqual(['1', '2', '3'])
    await expect(paragraphOf(page, 'uno')).toBeVisible()
    await undo(app)
    await expect(paragraphOf(page, 'uno')).toHaveCount(0)
    await expect.poll(async () => (await menuItem(app, ['Edit', 'Undo']))?.enabled).toBe(false)
    await page.screenshot({ path: test.info().outputPath('notes-undone.png') })

    await saveFromMenu(app)
    await expect
      .poll(async () => footnotePr(await part(docPath, 'word/settings.xml')).slice(0, 2), {
        timeout: 30_000,
      })
      .toEqual(['decimal', '1'])
    const undone = await sectPrs(docPath)
    expect(undone.map((xml) => footnotePr(xml).slice(0, 2))).toEqual([
      ['decimal', '1'],
      ['decimal', '1'],
    ])
    expect(await part(docPath, 'word/document.xml')).not.toContain('lowerLetter')

    // redo walks the same steps forward
    await redo(app)
    await expect(paragraphOf(page, 'uno')).toBeVisible()
    await expect.poll(() => marks(page)).toEqual(['1', '2', '3'])
    await redo(app)
    await expect.poll(() => marks(page)).toEqual(['c', 'd', 'e'])
    await redo(app)
    await expect(paragraphOf(page, 'duo')).toBeVisible()
    await redo(app)
    await expect.poll(() => marks(page)).toEqual(['c', 'd', 'V'])
    await saveFromMenu(app)
    await expect
      .poll(async () => footnotePr((await sectPrs(docPath))[1] ?? ''), { timeout: 30_000 })
      .toEqual(['upperRoman', '5', 'eachSect'])
    expect(footnotePr(await part(docPath, 'word/settings.xml')).slice(0, 2)).toEqual([
      'lowerLetter',
      '3',
    ])

    // read back
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    session = await openDoc(docPath)
    ;({ app, page } = session)
    await expect.poll(() => marks(page)).toEqual(['c', 'd', 'V'])
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

const LOWER =
  'lower internationalization of organizational responsibilities demonstrates characteristically ' +
  'disproportionate institutional representation throughout intercontinental telecommunication ' +
  'infrastructure notwithstanding considerable administrative misunderstanding concerning ' +
  'environmental sustainability recommendations and multidisciplinary investigations into ' +
  'uncharacteristically complicated organizational communication problems considering ' +
  'international transportation regulations and extraordinary governmental accountability'
const CAPS =
  'UPPER INTERNATIONALIZATION OF ORGANIZATIONAL RESPONSIBILITIES DEMONSTRATES CHARACTERISTICALLY ' +
  'DISPROPORTIONATE INSTITUTIONAL REPRESENTATION THROUGHOUT INTERCONTINENTAL TELECOMMUNICATION'

/** line ends of a paragraph that end in a hyphen of ours, and the longest run of such lines */
const hyphenLines = (page: Page, startsWith: string) =>
  page.evaluate((startsWith) => {
    const el = Array.from(document.querySelectorAll<HTMLElement>('.ProseMirror.doc-page p')).find(
      (p) => p.textContent?.startsWith(startsWith),
    )
    if (!el) return null
    const range = document.createRange()
    range.selectNodeContents(el)
    const tops = Array.from(
      new Set(Array.from(range.getClientRects()).map((r) => Math.round(r.top))),
    )
      .sort((a, b) => a - b)
      .filter((top, i, all) => i === 0 || top - all[i - 1] > 4)
    const lineOf = (top: number) => tops.findIndex((t) => Math.abs(t - top) <= 4)
    const broken = Array.from(el.querySelectorAll('.doc-shy'))
      .map((s) => s.getBoundingClientRect())
      .filter((r) => r.width > 0.01)
      .map((r) => lineOf(Math.round(r.top)))
    let longest = 0
    let run = 0
    broken.forEach((line, i) => {
      run = i > 0 && line === broken[i - 1] + 1 ? run + 1 : 1
      longest = Math.max(longest, run)
    })
    return { hyphens: broken.length, longest }
  }, startsWith)

/** what the saved settings say about hyphenation: on, capitals excluded, the limit */
const hyphenationIn = async (docPath: string) => {
  const settings = await part(docPath, 'word/settings.xml')
  return [
    /<w:autoHyphenation(?:\s+w:val="(?:1|true|on)")?\s*\/>/.test(settings),
    /<w:doNotHyphenateCaps(?:\s+w:val="(?:1|true|on)")?\s*\/>/.test(settings),
    val(settings, 'consecutiveHyphenLimit') ?? '0',
  ]
}

test('one OK in the Hyphenation dialog is one undo step, in order with typing, and saves what is on the page', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-settings-undo-')))
  const docPath = join(dir, 'hyphenation.docx')
  writeFileSync(
    docPath,
    await withLanguage(await docxWith(para('apple') + para(LOWER) + para(CAPS) + SECT_PR)),
  )
  const session = await openDoc(docPath)
  try {
    const { app, page } = session
    const dialog = () => page.locator('.hyphenation-dialog')
    const openDialog = async () => {
      await clickMenu(app, ['Tools', 'Hyphenation…'])
      await expect(dialog()).toBeVisible()
    }
    const ok = async () => {
      await dialog().locator('button.primary').click()
      await expect(dialog()).toBeHidden()
    }
    const auto = () => dialog().getByLabel('Automatically hyphenate document')
    const caps = () => dialog().getByLabel('Hyphenate words in CAPS')
    const limit = () => dialog().getByLabel('Limit consecutive hyphens to')
    const lower = async () => (await hyphenLines(page, 'lower'))?.hyphens
    const upper = async () => (await hyphenLines(page, 'UPPER'))?.hyphens

    await expect(paragraphOf(page, 'lower')).toBeVisible()
    expect(await lower()).toBe(0)
    expect((await menuItem(app, ['Edit', 'Undo']))?.enabled).toBe(false)

    await typeIn(page, 'apple', 'one ')
    await openDialog()
    await expect(auto()).not.toBeChecked()
    await auto().check()
    await ok()
    await expect.poll(lower, { timeout: 30_000 }).toBeGreaterThan(3)
    await expect.poll(upper).toBeGreaterThan(0)
    const on = await hyphenLines(page, 'lower')
    const onUpper = await upper()

    await typeIn(page, 'apple', 'two ')
    // two rules in one OK
    await openDialog()
    await caps().uncheck()
    await limit().selectOption('1')
    await ok()
    await expect.poll(upper).toBe(0)
    await expect.poll(async () => (await hyphenLines(page, 'lower'))?.longest).toBe(1)
    await page.screenshot({ path: test.info().outputPath('hyphenation-changed.png') })
    await saveFromMenu(app)
    await expect.poll(() => hyphenationIn(docPath), { timeout: 30_000 }).toEqual([true, true, '1'])

    // both rules come back in one step; the save after it takes them out of the file
    await undo(app)
    await expect.poll(upper).toBe(onUpper)
    await expect.poll(lower).toBe(on!.hyphens)
    await expect(paragraphOf(page, 'two one apple')).toBeVisible()
    await openDialog()
    await expect(auto()).toBeChecked()
    await expect(caps()).toBeChecked()
    await expect(limit()).toHaveValue('0')
    await ok()
    await saveFromMenu(app)
    await expect.poll(() => hyphenationIn(docPath), { timeout: 30_000 }).toEqual([true, false, '0'])

    await undo(app)
    await expect(paragraphOf(page, 'two one apple')).toHaveCount(0)
    await expect.poll(lower).toBe(on!.hyphens)
    await undo(app)
    await expect.poll(lower).toBe(0)
    await expect(paragraphOf(page, 'one apple')).toBeVisible()
    await undo(app)
    await expect(paragraphOf(page, 'one apple')).toHaveCount(0)
    await expect.poll(async () => (await menuItem(app, ['Edit', 'Undo']))?.enabled).toBe(false)
    await page.screenshot({ path: test.info().outputPath('hyphenation-undone.png') })
    await saveFromMenu(app)
    await expect
      .poll(() => hyphenationIn(docPath), { timeout: 30_000 })
      .toEqual([false, false, '0'])

    await redo(app)
    await expect(paragraphOf(page, 'one apple')).toBeVisible()
    expect(await lower()).toBe(0)
    await redo(app)
    await expect.poll(lower, { timeout: 30_000 }).toBe(on!.hyphens)
    await redo(app)
    await expect(paragraphOf(page, 'two one apple')).toBeVisible()
    await redo(app)
    await expect.poll(upper).toBe(0)
    await expect.poll(async () => (await hyphenLines(page, 'lower'))?.longest).toBe(1)
    await saveFromMenu(app)
    await expect.poll(() => hyphenationIn(docPath), { timeout: 30_000 }).toEqual([true, true, '1'])
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
