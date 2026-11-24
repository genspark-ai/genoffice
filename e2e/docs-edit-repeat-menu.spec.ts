/**
 * Edit > Repeat (Word's second Edit item): the last command again on the
 * current selection, Redo while there is something to redo, greyed out as
 * "Can't Repeat" otherwise. Also the File / Edit / Tools items Word has that
 * sit in the menu at Word's place.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-edit-repeat-menu'
const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '</w:sectPr>'

interface Item {
  label: string
  enabled: boolean
  type: string
  role?: string
  accelerator?: string
}

/** the items of a menu, by the labels leading to it */
const menuItems = (app: ElectronApplication, path: string[]): Promise<Item[]> =>
  app.evaluate(({ Menu }, path) => {
    let items = Menu.getApplicationMenu()?.items
    for (const label of path)
      items = items?.find((candidate) => candidate.label === label)?.submenu?.items
    return (items ?? []).map((item) => ({
      label: item.label,
      enabled: item.enabled,
      type: item.type,
      ...(item.role ? { role: item.role } : {}),
      ...(item.accelerator ? { accelerator: String(item.accelerator) } : {}),
    }))
  }, path)

/** Edit's second item: Repeat under whichever wording it has now */
const repeatItem = async (app: ElectronApplication) => (await menuItems(app, ['Edit']))[1]

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

const paragraph = (page: Page, text: string) =>
  page.locator('.doc-page p', { hasText: text }).first()
/** puts the caret in a paragraph and waits for the editor to have it there */
async function caretIn(page: Page, text: string, double = false) {
  const target = paragraph(page, text)
  if (double) await target.dblclick({ position: { x: 8, y: 8 } })
  else await target.click()
  await expect
    .poll(() =>
      page.evaluate(() => {
        const { editor } = (
          window as never as {
            __aidocs: {
              editor: { state: { selection: { $from: { parent: { textContent: string } } } } }
            }
          }
        ).__aidocs
        return editor.state.selection.$from.parent.textContent
      }),
    )
    .toContain(text)
}

const alignOf = (page: Page, text: string) =>
  paragraph(page, text).evaluate((el) => getComputedStyle(el).textAlign)

const documentXml = async (docPath: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file('word/document.xml')?.async('string') ?? ''
/** the paragraph of the saved document that holds this text */
const savedParagraph = (xml: string, text: string) =>
  xml.split(/<w:p[ >]/).find((p) => p.replace(/<[^>]+>/g, '').includes(text)) ?? ''

test('Edit > Repeat runs the last command on the new selection, is Redo after an undo, and greys out once the document moves on', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-repeat-')))
  const docPath = join(dir, 'repeat.docx')
  writeFileSync(
    docPath,
    await docxWith(
      para('alpha one') + para('beta two') + para('gamma three') + para('delta four') + SECT_PR,
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
    await expect(paragraph(page, 'alpha one')).toBeVisible({ timeout: 30_000 })
    await caretIn(page, 'alpha one')

    // nothing done yet: Word's "Can't Repeat", greyed out, on Word's key
    await expect
      .poll(() => repeatItem(app), { timeout: 30_000 })
      .toEqual({
        label: "Can't Repeat",
        enabled: false,
        type: 'normal',
        accelerator: 'CmdOrCtrl+Y',
      })
    expect((await menuItems(app, ['Edit'])).map((item) => item.label)).not.toContain('Redo')

    // a ribbon command is remembered: Center, then Repeat on another paragraph
    await page.locator('[data-ribbon-body] button[aria-label^="Center"]').first().click()
    await expect.poll(() => alignOf(page, 'alpha one')).toBe('center')
    await expect.poll(async () => (await repeatItem(app)).label).toBe('Repeat')
    await caretIn(page, 'beta two')
    expect(await alignOf(page, 'beta two')).not.toBe('center')
    await clickMenu(app, ['Edit', 'Repeat'])
    await expect.poll(() => alignOf(page, 'beta two')).toBe('center')
    // moving the caret keeps it, and a repeated command can be repeated again
    await caretIn(page, 'gamma three')
    expect((await repeatItem(app)).enabled).toBe(true)
    await page.screenshot({ path: test.info().outputPath('repeat-center.png') })

    // a dialog's pick rides as the command's payload: UPPERCASE, then Repeat on other words
    await caretIn(page, 'gamma three', true)
    await clickMenu(app, ['Format', 'Change Case…'])
    const caseDialog = page.locator('.case-dialog')
    await caseDialog.locator('label', { hasText: 'UPPERCASE' }).locator('input').check()
    await caseDialog.locator('button.primary').click()
    await expect(paragraph(page, 'GAMMA')).toBeVisible()
    await caretIn(page, 'delta four', true)
    await clickMenu(app, ['Edit', 'Repeat'])
    await expect(paragraph(page, 'DELTA four')).toBeVisible()
    expect(await alignOf(page, 'DELTA four')).not.toBe('center')

    // after an undo the item is Redo, and puts the change back
    await clickMenu(app, ['Edit', 'Undo'])
    await expect(paragraph(page, 'delta four')).toBeVisible()
    await expect.poll(async () => (await repeatItem(app)).label).toBe('Redo')
    await clickMenu(app, ['Edit', 'Redo'])
    await expect(paragraph(page, 'DELTA four')).toBeVisible()
    // Shift+Cmd+Z stays Redo without a menu item of its own
    await clickMenu(app, ['Edit', 'Undo'])
    await expect(paragraph(page, 'delta four')).toBeVisible()
    await page.keyboard.press('ControlOrMeta+Shift+z')
    await expect(paragraph(page, 'DELTA four')).toBeVisible()

    // typing is not a command: once the document moves on there is nothing to repeat
    await page.evaluate(() => {
      const { editor } = (
        window as never as {
          __aidocs: {
            editor: {
              state: { selection: { to: number } }
              chain(): { focus(): { setTextSelection(pos: number): { run(): boolean } } }
            }
          }
        }
      ).__aidocs
      editor.chain().focus().setTextSelection(editor.state.selection.to).run()
    })
    await clickMenu(app, ['Format', 'Align', 'Align Right'])
    await expect.poll(async () => (await repeatItem(app)).label).toBe('Repeat')
    await page.keyboard.type('typed ')
    await expect
      .poll(async () => {
        const item = await repeatItem(app)
        return [item.label, item.enabled]
      })
      .toEqual(["Can't Repeat", false])

    await saveFromMenu(app)
    await expect
      .poll(async () => savedParagraph(await documentXml(docPath), 'beta two'), { timeout: 30_000 })
      .toContain('<w:jc w:val="center"/>')
    const xml = await documentXml(docPath)
    expect(savedParagraph(xml, 'alpha one')).toContain('<w:jc w:val="center"/>')
    expect(xml.replace(/<[^>]+>/g, '')).toContain('GAMMA')
    expect(xml.replace(/<[^>]+>/g, '')).toContain('DELTA')
    expect(savedParagraph(xml, 'DELTA')).toContain('<w:jc w:val="right"/>')
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Repeat outlives a header that was opened and left alone, and ends once the header was typed in', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-repeat-header-')))
  const docPath = join(dir, 'repeat-header.docx')
  writeFileSync(docPath, await docxWith(para('alpha one') + para('beta two') + SECT_PR))
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await expect(paragraph(page, 'alpha one')).toBeVisible({ timeout: 30_000 })
    // the open-time respell kick focuses the body and shields keystrokes for a moment
    await page.waitForTimeout(1500)
    await caretIn(page, 'alpha one')
    await page.locator('[data-ribbon-body] button[aria-label^="Center"]').first().click()
    await expect.poll(() => alignOf(page, 'alpha one')).toBe('center')
    await expect.poll(async () => (await repeatItem(app)).label).toBe('Repeat')

    const header = page.locator('.page-hf-header').first()
    const strip = page.locator('.page-hf-editor')
    await clickMenu(app, ['View', 'Header and Footer'])
    await expect(strip).toBeFocused()
    await page.keyboard.press('Escape')
    await expect(strip).toHaveCount(0)
    await caretIn(page, 'beta two')
    expect(await repeatItem(app)).toMatchObject({ label: 'Repeat', enabled: true })

    await clickMenu(app, ['View', 'Header and Footer'])
    await expect(strip).toBeFocused()
    await page.keyboard.type('Head')
    await page.keyboard.press('Escape')
    await expect(strip).toHaveCount(0)
    await expect(header).toContainText('Head')
    await caretIn(page, 'beta two')
    await expect
      .poll(async () => {
        const item = await repeatItem(app)
        return [item.label, item.enabled]
      })
      .toEqual(["Can't Repeat", false])
    // the menu is rebuilt a moment after the state changes: it must not come back
    await page.waitForTimeout(500)
    expect((await repeatItem(app)).enabled).toBe(false)
    expect(await alignOf(page, 'beta two')).not.toBe('center')

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const zip = await JSZip.loadAsync(readFileSync(docPath))
          const headers = Object.keys(zip.files).filter((name) =>
            /^word\/header\d+\.xml$/.test(name),
          )
          const parts = await Promise.all(headers.map((name) => zip.file(name)!.async('string')))
          return parts.some((xml) => xml.includes('Head'))
        },
        { timeout: 30_000 },
      )
      .toBe(true)
    expect(savedParagraph(await documentXml(docPath), 'beta two')).not.toContain('w:val="center"')
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test("File, Edit and Tools carry Word's remaining items at Word's place", async () => {
  test.setTimeout(120_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-menu-gaps-')))
  const docPath = join(dir, 'first.docx')
  const otherPath = join(dir, 'other.docx')
  writeFileSync(docPath, await docxWith(para('first document') + SECT_PR))
  writeFileSync(otherPath, await docxWith(para('opened through More') + SECT_PR))
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await expect(paragraph(page, 'first document')).toBeVisible({ timeout: 30_000 })
    await paragraph(page, 'first document').click()

    const edit = await menuItems(app, ['Edit'])
    const labels = edit.map((item) => (item.type === 'separator' ? '---' : item.label))
    // Word: Paste, Paste Special…, Paste and Match Formatting
    const paste = labels.indexOf('Paste')
    expect(labels.slice(paste, paste + 3)).toEqual([
      'Paste',
      'Paste Special…',
      'Paste and Match Formatting',
    ])
    // its click arms the plain paste, which a role item would never run
    expect(edit[paste + 2].role).toBeUndefined()
    expect(edit[paste + 2].accelerator).toBe(
      process.platform === 'darwin' ? 'CmdOrCtrl+Alt+Shift+V' : 'CmdOrCtrl+Shift+V',
    )
    const find = labels.indexOf('Find')
    expect(labels.slice(find, find + 5)).toEqual(['Find', '---', 'Links…', '---', 'Select Data…'])

    // the grammar marks of an AI check can be hidden: the toggle follows Hide Spelling Errors
    const spelling = await menuItems(app, ['Tools', 'Spelling and Grammar'])
    const hideSpelling = spelling.findIndex((item) => item.label === 'Hide Spelling Errors')
    expect(spelling[hideSpelling + 1]).toMatchObject({
      label: 'Hide Grammar Errors',
      type: 'checkbox',
    })
    expect(spelling[hideSpelling + 2].type).toBe('separator')

    // Open Recent ends in a separator and More…, which opens the Open dialog
    const recent = await menuItems(app, ['File', 'Open Recent'])
    expect(recent.at(-2)?.type).toBe('separator')
    expect(recent.at(-1)).toMatchObject({
      label: 'More…',
      enabled: true,
      accelerator: 'Shift+CmdOrCtrl+O',
    })
    await app.evaluate(({ dialog }, file) => {
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [file] })
    }, otherPath)
    await clickMenu(app, ['File', 'Open Recent', 'More…'])
    await expect
      .poll(
        async () => {
          for (const win of app.windows())
            if ((await win.locator('.doc-page p', { hasText: 'opened through More' }).count()) > 0)
              return true
          return false
        },
        { timeout: 30_000 },
      )
      .toBe(true)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
