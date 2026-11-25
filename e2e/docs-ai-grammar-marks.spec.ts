/**
 * Grammar checking is the AI's: Tools > Spelling and Grammar > Spelling and
 * Grammar… asks the model once, the findings show as marks on the page, the
 * right-click menu of a mark offers its suggestions, and Hide Grammar Errors
 * hides the marks. Nothing of it reaches the saved file.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { createServer, type Server } from 'node:http'
import { mkdirSync, mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import type { AddressInfo } from 'node:net'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { defaultAiSettings, resolveCustomEndpoints } from '@genoffice/ai-provider'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu, settingsXml, withWordPart } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-ai-grammar-marks'
const POLL = { timeout: 20_000, intervals: [250, 500, 1000] }
const SPELLING = ['Tools', 'Spelling and Grammar']
const CHECK = [...SPELLING, 'Spelling and Grammar…']
const HIDE = [...SPELLING, 'Hide Grammar Errors']
const AGREEMENT = 'The subject and the verb do not agree.'

interface Sent {
  p: number
  text: string
}

/** a model that knows three mistakes, and also reports one that is not in the text */
function startModelStub(asked: Sent[][]): Promise<Server> {
  const server = createServer((req, res) => {
    let body = ''
    req.on('data', (chunk) => (body += chunk))
    req.on('end', () => {
      const messages = (JSON.parse(body) as { messages: Array<{ role: string; content: string }> })
        .messages
      const sent = JSON.parse(messages.filter((m) => m.role === 'user').at(-1)!.content) as Sent[]
      asked.push(sent)
      const found: unknown[] = [{ p: sent[0].p, quote: 'not in the text', message: 'Made up.' }]
      for (const { p, text } of sent) {
        if (text.includes('They is'))
          found.push({
            p,
            quote: 'They is',
            occurrence: 0,
            message: AGREEMENT,
            suggestions: ['They are'],
          })
        if (text.split('report are').length > 2)
          found.push({
            p,
            quote: 'report are',
            occurrence: 1,
            message: AGREEMENT,
            suggestions: ['report is'],
          })
        if (text.includes('more better'))
          found.push({
            p,
            quote: 'more better',
            message: 'A double comparative.',
            suggestions: ['better'],
          })
      }
      res.writeHead(200, { 'content-type': 'application/json' })
      res.end(
        JSON.stringify({
          choices: [
            {
              message: {
                role: 'assistant',
                content: '```json\n' + JSON.stringify(found) + '\n```',
              },
              finish_reason: 'stop',
            },
          ],
        }),
      )
    })
  })
  return new Promise((resolve) => server.listen(0, '127.0.0.1', () => resolve(server)))
}

interface Item {
  label: string
  type: string
  enabled: boolean
  checked: boolean
}

const menuItem = (app: ElectronApplication, path: string[]): Promise<Item | null> =>
  app.evaluate(({ Menu }, path) => {
    let items = Menu.getApplicationMenu()?.items
    let item: Electron.MenuItem | undefined
    for (const label of path) {
      item = items?.find((candidate) => candidate.label === label)
      items = item?.submenu?.items
    }
    return item
      ? { label: item.label, type: item.type, enabled: item.enabled, checked: item.checked }
      : null
  }, path)

async function clickMenu(app: ElectronApplication, path: string[]) {
  await expect.poll(async () => (await menuItem(app, path))?.enabled, POLL).toBe(true)
  await app.evaluate(({ Menu }, path) => {
    let items = Menu.getApplicationMenu()?.items
    let item: Electron.MenuItem | undefined
    for (const label of path) {
      item = items?.find((candidate) => candidate.label === label)
      items = item?.submenu?.items
    }
    item?.click()
  }, path)
}

/** what the page draws: the ranges of the grammar highlight, and the findings behind them */
const marks = (page: Page) =>
  page.evaluate(() => {
    const drawn = (CSS as unknown as { highlights: Map<string, Set<Range>> }).highlights.get(
      'docs-grammar',
    )
    return [...(drawn ?? [])].map((range) => range.toString())
  })

const wordCenter = (page: Page, words: string, occurrence = 0) =>
  page.evaluate(
    ([w, n]) => {
      const root = document.querySelector('.doc-page')!
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
      let seen = 0
      for (let node = walker.nextNode(); node; node = walker.nextNode()) {
        let idx = node.textContent?.indexOf(w as string) ?? -1
        while (idx !== -1) {
          if (seen++ === n) {
            const range = document.createRange()
            range.setStart(node, idx)
            range.setEnd(node, idx + (w as string).length)
            const r = range.getBoundingClientRect()
            return { x: r.left + r.width / 2, y: r.top + r.height / 2 }
          }
          idx = node.textContent!.indexOf(w as string, idx + 1)
        }
      }
      throw new Error(`no ${w}`)
    },
    [words, occurrence] as const,
  )

async function openMenuOn(page: Page, words: string, occurrence = 0) {
  const at = await wordCenter(page, words, occurrence)
  await page.mouse.click(at.x, at.y, { button: 'right' })
  const menu = page.locator('.ctx-menu')
  await menu.waitFor()
  return menu
}

async function closeMenu(page: Page) {
  const box = (await page.locator('.doc-page').boundingBox())!
  await page.mouse.click(box.x + 10, box.y + 10)
  await expect(page.locator('.ctx-menu')).toHaveCount(0)
}

const shot = async (page: Page, name: string) => {
  if (process.env.GRAMMAR_SHOTS)
    await page.screenshot({ path: join(process.env.GRAMMAR_SHOTS, `${name}.png`) })
}

const BODY = [
  'They is ready for the meeting.',
  'The report are finished and the report are late.',
  'This plan is more better than the old one.',
]
  .map(para)
  .join('')

async function savedBody(docPath: string): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(docPath))
  return zip.file('word/document.xml')!.async('string')
}

test('a grammar check marks the page; suggestions, Ignore Once, Grammar… and Hide Grammar Errors work', async () => {
  test.setTimeout(240_000)
  const asked: Sent[][] = []
  const server = await startModelStub(asked)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-grammar-')))
  const userDataDir = join(dir, 'user-data')
  const docPath = join(dir, 'grammar.docx')
  writeFileSync(docPath, await docxWith(BODY))
  const settings = defaultAiSettings()
  settings.provider = 'custom'
  settings.providers.custom = {
    apiKey: 'stub',
    model: 'stub-model',
    baseUrl: `http://127.0.0.1:${(server.address() as AddressInfo).port}/v1`,
  }
  mkdirSync(userDataDir)
  writeFileSync(
    join(userDataDir, 'ai-settings.json'),
    JSON.stringify(resolveCustomEndpoints(settings)),
  )
  const launched = await launchShell({
    userDataDir,
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'more better' })).toBeVisible({
      timeout: 30_000,
    })
    // the check walks on from the caret: put it at the start
    await page.locator('.doc-page p', { hasText: 'They is' }).click()
    await page.keyboard.press('Home')
    // the toggle is live before any check, and opening the document asked the model nothing
    await expect
      .poll(() => menuItem(app, HIDE), POLL)
      .toMatchObject({ type: 'checkbox', enabled: true, checked: false })
    expect(asked).toHaveLength(0)

    await clickMenu(app, CHECK)
    const dialog = page.locator('.spelling-dialog')
    // the document's text goes to the model only once Check grammar is ticked
    await expect(dialog.getByLabel('Check grammar')).not.toBeChecked()
    expect(asked).toHaveLength(0)
    await dialog.getByLabel('Check grammar').check()
    await expect(dialog.locator('.spelling-context b')).toHaveText('They is', POLL)
    await expect(dialog.locator('.spelling-problem')).toHaveText(AGREEMENT)
    await expect(dialog.getByRole('option')).toHaveText(['They are'])
    await expect(dialog.getByLabel('Check grammar')).toBeChecked()
    await expect(dialog.getByRole('button', { name: 'Ignore All', exact: true })).toBeDisabled()
    expect(asked).toHaveLength(1)
    expect(asked[0].map((s) => s.text)).toEqual([
      'They is ready for the meeting.',
      'The report are finished and the report are late.',
      'This plan is more better than the old one.',
    ])
    // the finding that quotes words the paragraph does not have is dropped
    expect(await marks(page)).toEqual(['They is', 'report are', 'more better'])
    await shot(page, 'dialog')
    await dialog.getByRole('button', { name: 'Close', exact: true }).click()
    await expect(dialog).toBeHidden()
    // the second "report are" is the marked one
    const second = await wordCenter(page, 'report are', 1)
    const markedAt = await page.evaluate(() => {
      const drawn = (CSS as unknown as { highlights: Map<string, Set<Range>> }).highlights.get(
        'docs-grammar',
      )!
      const r = [...drawn][1].getBoundingClientRect()
      return { x: r.left + r.width / 2, y: r.top + r.height / 2 }
    })
    expect(Math.abs(markedAt.x - second.x)).toBeLessThan(2)

    // right-click on a mark: what is wrong, the suggestion, Ignore Once, Grammar…
    let menu = await openMenuOn(page, 'They is')
    await expect(menu.locator('> .ctx-item .ctx-label').first()).toHaveText(AGREEMENT)
    await expect(menu.locator('> .ctx-item').first()).toBeDisabled()
    await expect(menu.locator('> .ctx-item-strong')).toHaveText(['They are'])
    await expect(menu.locator('> .ctx-item', { hasText: 'Ignore Once' })).toBeVisible()
    await expect(menu.locator('> .ctx-item', { hasText: 'Grammar…' })).toBeVisible()
    await shot(page, 'context-menu')
    await menu.locator('> .ctx-item-strong', { hasText: 'They are' }).click()
    await expect(menu).toHaveCount(0)
    const first = page.locator('.doc-page p').first()
    await expect(first).toHaveText('They are ready for the meeting.')
    expect(await marks(page)).toEqual(['report are', 'more better'])
    // one undo step takes the suggestion back; the edited paragraph has no mark until the next check
    // the suggestion is the only edit so far, so a repeated press has nothing more to take back
    await expect(async () => {
      await page.keyboard.press('ControlOrMeta+z')
      await expect(first).toHaveText('They is ready for the meeting.', { timeout: 1000 })
    }).toPass(POLL)
    expect(await marks(page)).toEqual(['report are', 'more better'])
    await expect(async () => {
      await page.keyboard.press('ControlOrMeta+Shift+z')
      await expect(first).toHaveText('They are ready for the meeting.', { timeout: 1000 })
    }).toPass(POLL)

    // Hide Grammar Errors: no marks and no menu items, and they are back when it is unticked
    await clickMenu(app, HIDE)
    await expect.poll(async () => (await menuItem(app, HIDE))?.checked, POLL).toBe(true)
    await expect.poll(() => marks(page), POLL).toEqual([])
    menu = await openMenuOn(page, 'more better')
    await expect(menu.locator('.ctx-item', { hasText: 'Cut' })).toBeVisible()
    await expect(menu.locator('.ctx-item', { hasText: 'Ignore Once' })).toHaveCount(0)
    await closeMenu(page)
    await clickMenu(app, HIDE)
    await expect.poll(async () => (await menuItem(app, HIDE))?.checked, POLL).toBe(false)
    await expect.poll(() => marks(page), POLL).toEqual(['report are', 'more better'])

    // Grammar… opens the dialog at that finding without asking the model again
    menu = await openMenuOn(page, 'report are', 1)
    await menu.locator('> .ctx-item', { hasText: 'Grammar…' }).click()
    await expect(dialog.locator('.spelling-context b')).toHaveText('report are')
    await dialog.getByRole('button', { name: 'Change', exact: true }).click()
    await expect(page.locator('.doc-page p').nth(1)).toHaveText(
      'The report are finished and the report is late.',
    )
    await expect(dialog.locator('.spelling-context b')).toHaveText('more better')
    await dialog.getByRole('button', { name: 'Close', exact: true }).click()
    expect(asked).toHaveLength(1)
    expect(await marks(page)).toEqual(['more better'])

    // Ignore Once: the mark goes and the next check does not bring it back
    menu = await openMenuOn(page, 'more better')
    await menu.locator('> .ctx-item', { hasText: 'Ignore Once' }).click()
    await expect.poll(() => marks(page), POLL).toEqual([])
    await clickMenu(app, CHECK)
    await expect(dialog.locator('.spelling-done')).toHaveText(
      'The spelling and grammar check is complete.',
      POLL,
    )
    expect(asked).toHaveLength(2)
    expect(await marks(page)).toEqual([])
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    // until Reset Ignored Words and Grammar
    await clickMenu(app, [...SPELLING, 'Reset Ignored Words and Grammar'])
    await clickMenu(app, CHECK)
    await expect(dialog.locator('.spelling-context b')).toHaveText('more better', POLL)
    await dialog.getByRole('button', { name: 'Close', exact: true }).click()
    expect(await marks(page)).toEqual(['more better'])
    await shot(page, 'marked-paragraph')

    // typing in the paragraph drops its mark, and asks the model nothing
    const third = page.locator('.doc-page p').nth(2)
    await third.click()
    await page.keyboard.press('End')
    await page.keyboard.type(' Yes.')
    await expect(third).toHaveText('This plan is more better than the old one. Yes.')
    await expect.poll(() => marks(page), POLL).toEqual([])
    expect(asked).toHaveLength(3)

    // the saved file has the accepted words and nothing of the marks
    await saveFromMenu(app)
    await expect.poll(() => savedBody(docPath), { timeout: 30_000 }).toContain('They are ready')
    const xml = await savedBody(docPath)
    expect(xml).toContain('the report is late.')
    expect(xml).toContain('more better than the old one. Yes.')
    expect(xml).not.toMatch(/<w:u\b|proofErr|grammar/i)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    server.close()
    rmSync(dir, { recursive: true, force: true })
  }
})

test('without an AI provider the dialog says grammar was not checked and spelling still works', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-grammar-')))
  const userDataDir = join(dir, 'user-data')
  const docPath = join(dir, 'spelling.docx')
  writeFileSync(docPath, await docxWith(para('They is going to recieve the parcel.')))
  mkdirSync(userDataDir)
  // a home of its own: the machine's Genspark login must not stand in for "no provider"
  const launched = await launchShell({
    userDataDir,
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
    env: { HOME: dir, USERPROFILE: dir, GSK_API_KEY: '' },
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'recieve' })).toBeVisible({
      timeout: 30_000,
    })
    await expect
      .poll(
        () =>
          page.evaluate(() =>
            (
              window as unknown as { desktop: { spellIsMisspelled(word: string): boolean } }
            ).desktop.spellIsMisspelled('recieve'),
          ),
        { timeout: 30_000 },
      )
      .toBe(true)
    await clickMenu(app, CHECK)
    const dialog = page.locator('.spelling-dialog')
    await expect(dialog.locator('.spelling-grammar-note')).toHaveCount(0)
    await dialog.getByLabel('Check grammar').check()
    await expect(dialog.locator('.spelling-grammar-note')).toContainText(
      'Grammar was not checked. It uses the AI provider set up in Settings.',
      POLL,
    )
    expect(await marks(page)).toEqual([])
    await shot(page, 'dialog-no-provider')
    // the spelling walk is the one the dialog always had
    await expect(dialog.locator('.spelling-context b')).toHaveText('recieve')
    await expect(dialog.locator('.spelling-problem')).toHaveText('Not in Dictionary:')
    await dialog.getByRole('option', { name: 'receive', exact: true }).click()
    await dialog.getByRole('button', { name: 'Change', exact: true }).click()
    await expect(dialog.locator('.spelling-done')).toHaveText('The spelling check is complete.')
    await expect(page.locator('.doc-page p').first()).toHaveText(
      'They is going to receive the parcel.',
    )
    // unticking Check grammar is remembered and takes the note away
    await dialog.getByLabel('Check grammar').uncheck()
    await expect(dialog.locator('.spelling-grammar-note')).toHaveCount(0)
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await clickMenu(app, CHECK)
    await expect(dialog.locator('.spelling-done')).toBeVisible()
    await expect(dialog.getByLabel('Check grammar')).not.toBeChecked()
    await expect(dialog.locator('.spelling-grammar-note')).toHaveCount(0)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('an occurrence the dialog was told to Ignore Once is not offered again by the next check', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-spelling-once-')))
  const userDataDir = join(dir, 'user-data')
  const docPath = join(dir, 'once.docx')
  writeFileSync(docPath, await docxWith(para('One plmokn here and one plmokn there.')))
  mkdirSync(userDataDir)
  const launched = await launchShell({
    userDataDir,
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
    env: { HOME: dir, USERPROFILE: dir, GSK_API_KEY: '' },
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'plmokn' })).toBeVisible({
      timeout: 30_000,
    })
    await expect
      .poll(
        () =>
          page.evaluate(() =>
            (
              window as unknown as { desktop: { spellIsMisspelled(word: string): boolean } }
            ).desktop.spellIsMisspelled('plmokn'),
          ),
        { timeout: 30_000 },
      )
      .toBe(true)
    const dialog = page.locator('.spelling-dialog')
    const ignoreOnce = dialog.getByRole('button', { name: 'Ignore Once', exact: true })
    await clickMenu(app, CHECK)
    await expect(dialog.locator('.spelling-context b')).toHaveText('plmokn')
    await ignoreOnce.click()
    await expect(dialog.locator('.spelling-context b')).toHaveText('plmokn')
    await ignoreOnce.click()
    await expect(dialog.locator('.spelling-done')).toBeVisible()
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog).toBeHidden()
    await clickMenu(app, CHECK)
    await expect(dialog.locator('.spelling-done')).toBeVisible()
    await shot(page, 'dialog-ignored-once')
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Hide Grammar Errors belongs to the document: read from the file, and saved with it', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-grammar-hide-')))
  const docPath = join(dir, 'hidden.docx')
  writeFileSync(
    docPath,
    await withWordPart(
      await docxWith(para('A plain line.')),
      'settings',
      settingsXml('<w:hideGrammaticalErrors/>'),
    ),
  )
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'A plain line.' })).toBeVisible({
      timeout: 30_000,
    })
    const settings = async () =>
      (await (
        await JSZip.loadAsync(readFileSync(docPath))
      )
        .file('word/settings.xml')
        ?.async('string')) ?? ''
    // the file says hidden
    await expect.poll(async () => (await menuItem(app, HIDE))?.checked, POLL).toBe(true)
    await clickMenu(app, HIDE)
    await expect.poll(async () => (await menuItem(app, HIDE))?.checked, POLL).toBe(false)
    await saveFromMenu(app)
    await expect.poll(settings, { timeout: 30_000 }).not.toContain('hideGrammaticalErrors')
    await clickMenu(app, HIDE)
    await expect.poll(async () => (await menuItem(app, HIDE))?.checked, POLL).toBe(true)
    await saveFromMenu(app)
    await expect.poll(settings, { timeout: 30_000 }).toContain('<w:hideGrammaticalErrors/>')
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
