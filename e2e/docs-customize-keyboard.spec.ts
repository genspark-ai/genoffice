/**
 * Tools > Customize Keyboard: a chord assigned in the dialog runs its command
 * in the document and shows on the menu item, a chord another command had
 * moves, Remove and Reset All undo it, and it all outlives a restart.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { existsSync, mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-customize-keyboard'
const IS_MAC = process.platform === 'darwin'
const MOD = IS_MAC ? 'Meta' : 'Control'
/** the chords as the dialog and the stored file write them on this platform */
const MOD_B = IS_MAC ? 'Cmd+B' : 'Ctrl+B'
// what the Bold menu item shows once Cmd/Ctrl+B is taken from it: the last chord it has left
const BOLD_LEFT = IS_MAC ? 'Ctrl+Shift+B' : ''
const shown = (mac: string, other: string) => (IS_MAC ? mac : other)
/** the chords Word for Mac alone gives a command, which it keeps beside an assigned one */
const withMac = (mac: string[], rest: string[]) => (IS_MAC ? [...mac, ...rest] : rest)

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

const menuAccelerator = (app: ElectronApplication, path: string[]) =>
  app.evaluate(({ Menu }, path) => {
    let items = Menu.getApplicationMenu()?.items
    let item: Electron.MenuItem | undefined
    for (const label of path) {
      item = items?.find((candidate) => candidate.label === label)
      items = item?.submenu?.items
    }
    return item ? (item.accelerator ?? '') : 'missing'
  }, path)

/** selects the whole text of the paragraph that reads `text` */
const selectLine = (page: Page, text: string) =>
  page.evaluate((text) => {
    interface Node {
      isTextblock: boolean
      textContent: string
    }
    const { editor } = (
      window as never as {
        __aidocs: {
          editor: {
            state: { doc: { descendants(visit: (node: Node, pos: number) => void): void } }
            chain(): { focus(): { setTextSelection(r: object): { run(): boolean } } }
          }
        }
      }
    ).__aidocs
    let from = -1
    editor.state.doc.descendants((node, pos) => {
      if (from < 0 && node.isTextblock && node.textContent === text) from = pos + 1
    })
    if (from < 0) throw new Error(`no paragraph reads "${text}"`)
    editor
      .chain()
      .focus()
      .setTextSelection({ from, to: from + text.length })
      .run()
  }, text)

const documentXml = async (docPath: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file('word/document.xml')?.async('string') ?? ''

/** the run properties of the paragraph holding `text`, as saved */
const savedRunProps = async (docPath: string, text: string) => {
  const paragraphs = (await documentXml(docPath)).split(/<w:p[ >]/)
  const paragraph = paragraphs.find((p) => p.includes(text)) ?? ''
  return (paragraph.match(/<w:rPr>.*?<\/w:rPr>/g) ?? []).join('')
}

async function openDoc(docPath: string, userDataDir?: string) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
    userDataDir,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await expect(page.locator('.doc-page p', { hasText: 'Third line' })).toBeVisible({
    timeout: 30_000,
  })
  return { launched, page }
}

async function openCustomizeKeyboard(app: ElectronApplication, page: Page) {
  await clickMenu(app, ['Tools', 'Customize Keyboard…'])
  const dialog = page.locator('.modal-customize-keyboard')
  await expect(dialog).toBeVisible()
  // the command list fills once the menu's commands arrive
  await expect(dialog.locator('[data-command="font-dialog"]')).toBeAttached()
  return dialog
}

test('a chord assigned in Customize Keyboard runs its command, shows in the menu, and survives a restart', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-customize-keyboard-')))
  const docPath = join(dir, 'keys.docx')
  writeFileSync(
    docPath,
    await docxWith(para('First line') + para('Second line') + para('Third line')),
  )
  const first = await openDoc(docPath)
  const { userDataDir } = first.launched
  const stored = () => {
    const file = join(userDataDir, 'docs-keyboard.json')
    return existsSync(file) ? (JSON.parse(readFileSync(file, 'utf8')) as object) : null
  }
  try {
    const { app } = first.launched
    const page = first.page

    // nothing is stored and the menu carries its built-in chords until an assignment is made
    expect(stored()).toBe(null)
    expect(await menuAccelerator(app, ['Format', 'Bold'])).toBe('CmdOrCtrl+B')
    expect(await menuAccelerator(app, ['File', 'Properties…'])).toBe('')

    const dialog = await openCustomizeKeyboard(app, page)
    const press = dialog.locator('#ck-press')
    const note = dialog.locator('.ck-note')
    const assign = dialog.getByRole('button', { name: 'Assign' })
    const currentKeys = dialog.locator('.ck-current button')
    const pick = async (id: string) => {
      await dialog.locator(`[data-command="${id}"]`).click()
      await expect(dialog.locator(`[data-command="${id}"]`)).toHaveAttribute(
        'aria-selected',
        'true',
      )
    }

    // the categories are the menus; a menu lists its own commands
    await dialog.getByRole('option', { name: 'Format', exact: true }).click()
    await expect(dialog.locator('[data-command="font-dialog"]')).toBeVisible()
    await expect(dialog.locator('[data-command="file-properties"]')).toHaveCount(0)
    await dialog.getByRole('option', { name: 'All Commands' }).click()

    // a free chord goes to a command that had none
    await pick('strikethrough')
    await expect(assign).toBeDisabled()
    await press.press('Control+Shift+y')
    await expect(press).toHaveValue(shown('⌃⇧Y', 'Ctrl+Shift+Y'))
    await expect(note).toHaveText('[unassigned]')
    await assign.click()
    await expect(currentKeys).toHaveText(withMac(['⇧⌘X'], [shown('⌃⇧Y', 'Ctrl+Shift+Y')]))

    // a typing key is refused
    await press.press('Shift+k')
    await expect(assign).toBeDisabled()
    // so is a built-in key that is no command
    await press.press(`${MOD}+w`)
    await expect(note).toContainText('Close')
    await expect(assign).toBeDisabled()

    // a chord another command has names that command, and Assign moves it
    await pick('numbering')
    await press.press('Control+Shift+y')
    await expect(note).toHaveText('Currently assigned to: Strikethrough')
    await assign.click()
    await expect(currentKeys).toHaveText([shown('⌃⇧Y', 'Ctrl+Shift+Y')])
    await pick('strikethrough')
    await expect(currentKeys).toHaveText(withMac(['⇧⌘X'], []))

    // a built-in chord moves the same way
    await press.press(`${MOD}+b`)
    await expect(note).toHaveText('Currently assigned to: Bold')
    await assign.click()
    await expect(currentKeys).toHaveText(withMac(['⇧⌘X'], [shown('⌘B', 'Ctrl+B')]))
    await pick('bold')
    await expect(currentKeys).toHaveText(withMac(['⇧⌘B', '⌃B', '⌃⇧B'], []))

    // a menu command with no chord of its own
    await pick('file-properties')
    await expect(dialog.locator('.ck-description p')).toHaveText('File ▸ Properties…')
    await press.press('Control+Shift+F7')
    await assign.click()

    await expect.poll(stored).toEqual({
      strikethrough: withMac(['Shift+Cmd+X'], [MOD_B]),
      numbering: ['Ctrl+Shift+Y'],
      bold: withMac(['Shift+Cmd+B', 'Ctrl+B', 'Ctrl+Shift+B'], []),
      'file-properties': ['Ctrl+Shift+F7'],
    })
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect(dialog).toBeHidden()

    // the menu shows the new chords and no longer shows the one that moved away
    expect(await menuAccelerator(app, ['File', 'Properties…'])).toBe('Ctrl+Shift+F7')
    expect(await menuAccelerator(app, ['Format', 'Bold'])).toBe(BOLD_LEFT)

    // in the document the moved chord runs its new command, not the old one
    await selectLine(page, 'First line')
    await page.keyboard.press(`${MOD}+b`)
    const firstLine = page.locator('.doc-page p', { hasText: 'First line' })
    await expect(firstLine.locator('s')).toHaveText('First line')
    await expect(firstLine.locator('strong, b')).toHaveCount(0)

    await selectLine(page, 'Second line')
    await page.keyboard.press('Control+Shift+y')
    await expect
      .poll(() =>
        page.evaluate(() =>
          (
            window as never as { __aidocs: { editor: { isActive(name: string): boolean } } }
          ).__aidocs.editor.isActive('docListItem'),
        ),
      )
      .toBe(true)

    await page.keyboard.press('Control+Shift+F7')
    const properties = page.locator('.properties-dialog')
    await expect(properties).toBeVisible()
    await page.keyboard.press('Escape')
    await expect(properties).toBeHidden()

    // the Keyboard Shortcuts panel lists what is in effect
    await clickMenu(app, ['Help', 'Keyboard Shortcuts'])
    const panel = page.locator('.modal-shortcuts')
    await expect(panel.locator('.sc-row', { hasText: 'Strikethrough' }).locator('kbd')).toHaveText(
      shown('⇧⌘X / ⌘B', 'Ctrl+B'),
    )
    await expect(panel.locator('.sc-row', { hasText: 'Properties' }).locator('kbd')).toHaveText(
      shown('⌃⇧F7', 'Ctrl+Shift+F7'),
    )
    await expect(panel.locator('.sc-row .sc-label', { hasText: /^Bold$/ })).toHaveCount(
      IS_MAC ? 1 : 0,
    )
    await page.keyboard.press('Escape')
    await expect(panel).toBeHidden()

    await saveFromMenu(app)
    await expect.poll(() => savedRunProps(docPath, 'First line')).toContain('<w:strike')
    expect(await savedRunProps(docPath, 'First line')).not.toContain('<w:b/>')
    expect(await documentXml(docPath)).toContain('<w:numPr>')
  } finally {
    await closeAndSaveVideo(first.launched, `${VIDEO_DIR}-assign`).catch(() => {})
  }

  const second = await openDoc(docPath, userDataDir)
  try {
    const { app } = second.launched
    const page = second.page

    // after a restart the assignments are still in the menu and in the document
    await expect.poll(() => menuAccelerator(app, ['File', 'Properties…'])).toBe('Ctrl+Shift+F7')
    expect(await menuAccelerator(app, ['Format', 'Bold'])).toBe(BOLD_LEFT)
    const thirdLine = page.locator('.doc-page p', { hasText: 'Third line' })
    await selectLine(page, 'Third line')
    await page.keyboard.press(`${MOD}+b`)
    await expect(thirdLine.locator('s')).toHaveText('Third line')

    const dialog = await openCustomizeKeyboard(app, page)
    const currentKeys = dialog.locator('.ck-current button')
    const remove = dialog.getByRole('button', { name: 'Remove' })

    // Remove unbinds the selected key
    await dialog.locator('[data-command="file-properties"]').click()
    await expect(currentKeys).toHaveText([shown('⌃⇧F7', 'Ctrl+Shift+F7')])
    await expect(remove).toBeDisabled()
    await currentKeys.first().click()
    await remove.click()
    await expect(currentKeys).toHaveCount(0)
    await expect.poll(() => menuAccelerator(app, ['File', 'Properties…'])).toBe('')

    // Reset All asks first, then restores every built-in chord
    await dialog.getByRole('button', { name: 'Reset All…' }).click()
    const confirm = dialog.locator('.ck-confirm')
    await expect(confirm).toContainText('Reset all keyboard shortcuts')
    await confirm.getByRole('button', { name: 'Cancel' }).click()
    await expect(confirm).toBeHidden()
    expect(stored()).not.toEqual({})
    await dialog.getByRole('button', { name: 'Reset All…' }).click()
    await confirm.getByRole('button', { name: 'OK' }).click()
    await expect.poll(stored).toEqual({})
    await expect(dialog.getByRole('button', { name: 'Reset All…' })).toBeDisabled()
    await dialog.locator('[data-command="bold"]').click()
    await expect(currentKeys).toHaveText([
      shown('⌘B', 'Ctrl+B'),
      ...withMac(['⇧⌘B', '⌃B', '⌃⇧B'], []),
    ])
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect(dialog).toBeHidden()

    expect(await menuAccelerator(app, ['Format', 'Bold'])).toBe('CmdOrCtrl+B')
    // the closing dialog hands the focus back a moment later, which can drop a selection made at once
    await expect(async () => {
      await selectLine(page, 'Second line')
      await page.keyboard.press(`${MOD}+b`)
      await expect(page.locator('.doc-page strong')).toHaveText('Second line', { timeout: 1_000 })
    }).toPass({ timeout: 15_000 })
    await saveFromMenu(app)
    await expect.poll(() => savedRunProps(docPath, 'Second line')).toContain('<w:b/>')
    expect(await savedRunProps(docPath, 'Second line')).not.toContain('<w:strike')
    expect(await savedRunProps(docPath, 'Third line')).toContain('<w:strike')
  } finally {
    await closeAndSaveVideo(second.launched, `${VIDEO_DIR}-restart`).catch(() => {})
    rmSync(dir, { recursive: true, force: true })
  }
})
