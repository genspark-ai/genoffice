/**
 * A document nobody changed has nothing to save. Opening one that is protected
 * (read-only, forms, comments) or going through Read Mode and back switches
 * whether the editor takes edits, and that switch used to count as an edit:
 * quitting asked to save, and AutoSave rewrote the untouched file.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import type { ChildProcess } from 'node:child_process'
import { createHash, randomBytes } from 'node:crypto'
import { mkdtempSync, readFileSync, realpathSync, statSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, waitForPageWithUrl } from './helpers'
import { docxWith, para } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-clean-open-stays-clean'
const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
const BODY = para('Plain paragraph of an untouched document') + para('Second paragraph') + SECT_PR
const PROTECTIONS = ['readOnly', 'forms', 'comments'] as const

async function docxWithSettings(settings: string): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await docxWith(BODY))
  zip.file(
    'word/settings.xml',
    '<?xml version="1.0" encoding="UTF-8"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      `${settings}</w:settings>`,
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

const protectedDocx = (edit: string) =>
  docxWithSettings(`<w:documentProtection w:edit="${edit}" w:enforcement="1"/>`)

/** a password to modify, hashed the way Word stores it */
function writeProtectedDocx(password: string): Promise<Buffer> {
  const spinCount = 1000
  const salt = randomBytes(16)
  let hash = createHash('sha512')
    .update(Buffer.concat([salt, Buffer.from(password, 'utf16le')]))
    .digest()
  for (let i = 0; i < spinCount; i++) {
    const iteration = Buffer.alloc(4)
    iteration.writeUInt32LE(i)
    hash = createHash('sha512')
      .update(Buffer.concat([hash, iteration]))
      .digest()
  }
  return docxWithSettings(
    '<w:writeProtection w:cryptProviderType="rsaAES" w:cryptAlgorithmClass="hash" w:cryptAlgorithmType="typeAny" ' +
      `w:cryptAlgorithmSid="14" w:cryptSpinCount="${spinCount}" w:hash="${hash.toString('base64')}" w:salt="${salt.toString('base64')}"/>`,
  )
}

async function writeDoc(name: string, bytes: Buffer): Promise<string> {
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-clean-open-')))
  const docPath = join(dir, name)
  writeFileSync(docPath, bytes)
  return docPath
}

const fileState = (docPath: string) => ({
  sha: createHash('sha256').update(readFileSync(docPath)).digest('hex'),
  mtimeMs: statSync(docPath).mtimeMs,
})

type Prompt = { message: string; buttons: string[] }

/**
 * Every native message box is recorded and answered with its Cancel button, so
 * a save prompt keeps the app alive where the test can still read what was asked.
 */
async function recordPrompts(app: ElectronApplication) {
  await app.evaluate(({ dialog }) => {
    const prompts: Prompt[] = []
    ;(globalThis as { __prompts?: Prompt[] }).__prompts = prompts
    dialog.showMessageBox = (async (...args: unknown[]) => {
      const options = args.find(
        (arg): arg is Electron.MessageBoxOptions =>
          !!arg && typeof arg === 'object' && 'message' in arg,
      )
      prompts.push({ message: options?.message ?? '', buttons: options?.buttons ?? [] })
      const cancel = options?.cancelId ?? (options?.buttons?.length ?? 1) - 1
      return { response: cancel, checkboxChecked: false }
    }) as typeof dialog.showMessageBox
  })
}

const promptsOf = (app: ElectronApplication) =>
  app.evaluate(() => (globalThis as { __prompts?: Prompt[] }).__prompts ?? [])

async function openDoc(docPath: string, settings?: Record<string, unknown>) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
    settings,
  })
  // taken now: the handle is gone once the app has exited
  const child = launched.app.process()
  await recordPrompts(launched.app)
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await expect(page.locator('.doc-page', { hasText: 'Plain paragraph' }).first()).toBeVisible({
    timeout: 30_000,
  })
  return { launched, page, child }
}

const isEditable = (page: Page) =>
  page.evaluate(
    () =>
      (window as never as { __aidocs: { editor: { isEditable: boolean } } }).__aidocs.editor
        .isEditable,
  )

/**
 * Quits the way the Quit menu item does. The process must exit without a
 * message box having been asked for; returns what was asked when it did not.
 */
async function quitAndCollectPrompts(
  app: ElectronApplication,
  child: ChildProcess,
): Promise<Prompt[] | 'still running'> {
  const exited = new Promise<void>((resolve) => {
    if (child.exitCode !== null || child.signalCode !== null) resolve()
    else child.once('exit', () => resolve())
  })
  await app.evaluate(({ app: electronApp }) => {
    setTimeout(() => electronApp.quit(), 0)
  })
  const deadline = Date.now() + 20_000
  let asked: Prompt[] = []
  let gone = false
  while (!gone && Date.now() < deadline) {
    gone = await Promise.race([
      exited.then(() => true),
      new Promise<boolean>((resolve) => setTimeout(() => resolve(false), 250)),
    ])
    if (gone) break
    asked = await promptsOf(app).catch(() => asked)
    if (asked.length > 0) break
  }
  if (gone) return asked
  child.kill('SIGKILL')
  await exited
  return asked.length > 0 ? asked : 'still running'
}

function forceClose(child: ChildProcess) {
  if (child.exitCode === null && child.signalCode === null) child.kill('SIGKILL')
}

for (const edit of PROTECTIONS) {
  test(`a document protected for ${edit} opens clean: quitting asks nothing and leaves the file alone`, async () => {
    test.setTimeout(120_000)
    const docPath = await writeDoc(`${edit}.docx`, await protectedDocx(edit))
    const before = fileState(docPath)
    const { launched, page, child } = await openDoc(docPath)
    try {
      await expect.poll(() => isEditable(page), { timeout: 15_000 }).toBe(false)
      // the open has settled: fonts, pagination and the protection state are all in
      await page.waitForTimeout(1500)
      expect(await promptsOf(launched.app)).toEqual([])
      expect(await quitAndCollectPrompts(launched.app, child)).toEqual([])
      expect(fileState(docPath)).toEqual(before)
    } finally {
      forceClose(child)
    }
  })
}

for (const answer of ['read-only', 'password'] as const) {
  test(`a document with a password to modify, opened with ${answer}, stays clean`, async () => {
    test.setTimeout(120_000)
    const docPath = await writeDoc(`modify-${answer}.docx`, await writeProtectedDocx('secret'))
    const before = fileState(docPath)
    const { launched, page, child } = await openDoc(docPath)
    try {
      const prompt = page.locator('.modal', { hasText: 'password to modify' })
      await expect(prompt).toBeVisible()
      await expect.poll(() => isEditable(page)).toBe(false)
      if (answer === 'read-only') {
        await prompt.getByRole('button', { name: 'Open Read-Only' }).click()
      } else {
        await prompt.locator('input').first().fill('secret')
        await prompt.getByRole('button', { name: 'OK' }).click()
      }
      await expect(prompt).toHaveCount(0)
      await expect.poll(() => isEditable(page)).toBe(answer === 'password')
      await page.waitForTimeout(1500)
      expect(await quitAndCollectPrompts(launched.app, child)).toEqual([])
      expect(fileState(docPath)).toEqual(before)
    } finally {
      forceClose(child)
    }
  })
}

/** enabled / disabled state of every ribbon control and menu bar item, by position */
async function chromeState(app: ElectronApplication, page: Page) {
  const ribbon = await page.evaluate(() =>
    Array.from(
      document.querySelectorAll<HTMLButtonElement | HTMLInputElement | HTMLSelectElement>(
        '[data-ribbon-body] button, [data-ribbon-body] input, [data-ribbon-body] select',
      ),
    ).map((el) => `${el.getAttribute('data-tip') ?? el.textContent ?? ''}:${el.disabled}`),
  )
  const menu = await app.evaluate(({ Menu }) => {
    const walk = (items: Electron.MenuItem[], prefix: string): string[] =>
      items.flatMap((item) => [
        `${prefix}${item.label}:${item.enabled}`,
        ...(item.submenu ? walk(item.submenu.items, `${prefix}${item.label} > `) : []),
      ])
    return walk(Menu.getApplicationMenu()?.items ?? [], '')
  })
  return { ribbon, menu }
}

test('Read Mode and back with Esc leaves a clean document clean, editable, and its ribbon and menus as they were', async () => {
  test.setTimeout(120_000)
  const docPath = await writeDoc('plain.docx', await docxWith(BODY))
  const before = fileState(docPath)
  const { launched, page, child } = await openDoc(docPath)
  const { app } = launched
  try {
    await page.locator('.doc-page p', { hasText: 'Second paragraph' }).first().click()
    await page
      .locator('.ribbon-tab', { hasText: /^View$/ })
      .first()
      .click()
    await page.waitForTimeout(800)
    const chromeBefore = await chromeState(app, page)
    expect(chromeBefore.ribbon.some((entry) => entry.endsWith(':false'))).toBe(true)

    await page
      .locator('[data-ribbon-body] button', { hasText: /^Read Mode$/ })
      .first()
      .click()
    await expect(page.locator('.app.read-mode')).toHaveCount(1)
    await expect.poll(() => isEditable(page)).toBe(false)

    await page.keyboard.press('Escape')
    await expect(page.locator('.app.read-mode')).toHaveCount(0)
    await expect.poll(() => isEditable(page)).toBe(true)
    await page.locator('.doc-page p', { hasText: 'Second paragraph' }).first().click()
    await expect.poll(() => chromeState(app, page), { timeout: 15_000 }).toEqual(chromeBefore)

    // nothing to undo: the switch is not an edit
    const undo = await app.evaluate(({ Menu }) => {
      const edit = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Edit')
      return edit?.submenu?.items.find((entry) => /^Undo/.test(entry.label))?.enabled
    })
    expect(undo).toBe(false)

    expect(await quitAndCollectPrompts(launched.app, child)).toEqual([])
    expect(fileState(docPath)).toEqual(before)
  } finally {
    forceClose(child)
  }
})

for (const edit of PROTECTIONS) {
  test(`AutoSave does not rewrite an untouched document protected for ${edit}`, async () => {
    test.setTimeout(150_000)
    const docPath = await writeDoc(`${edit}-autosave.docx`, await protectedDocx(edit))
    const before = fileState(docPath)
    const { launched, page, child } = await openDoc(docPath, {
      autoSaveDefault: true,
      autoSaveDefaultUpdatedAt: Date.now(),
    })
    try {
      await expect(page.locator('.autosave-toggle input')).toBeChecked()
      await expect.poll(() => isEditable(page), { timeout: 15_000 }).toBe(false)
      if (edit === 'readOnly') {
        // the 30 s autosave timer itself
        await page.waitForTimeout(33_000)
      } else {
        // the timer's other trigger: the window losing focus
        await page.waitForTimeout(1500)
        await page.evaluate(() => window.dispatchEvent(new Event('blur')))
        await page.waitForTimeout(3000)
      }
      expect(fileState(docPath)).toEqual(before)
      // with AutoSave on a dirty document is saved on the way out instead of prompting
      expect(await quitAndCollectPrompts(launched.app, child)).toEqual([])
      expect(fileState(docPath)).toEqual(before)
    } finally {
      forceClose(child)
    }
  })
}
