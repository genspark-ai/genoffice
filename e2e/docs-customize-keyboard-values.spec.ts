/**
 * Tools > Customize Keyboard, commands that take a value: a font from the
 * Fonts category, a style from the Styles category, and a size picked for
 * Font Size each get a chord that applies that value in the document, is
 * written by the save, and outlives a restart.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { existsSync, mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-customize-keyboard-values'
const IS_MAC = process.platform === 'darwin'
const shown = (mac: string, other: string) => (IS_MAC ? mac : other)

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

/** the paragraph holding `text`, as saved */
const savedParagraph = async (docPath: string, text: string) => {
  const xml =
    (await JSZip.loadAsync(readFileSync(docPath))).file('word/document.xml')?.async('string') ?? ''
  return (await xml).split(/<w:p[ >]/).find((p) => p.includes(text)) ?? ''
}

async function fixture(): Promise<Buffer> {
  const source = await JSZip.loadAsync(
    await docxWith(
      para('First line') + para('Second line') + para('Third line') + para('Fourth line'),
    ),
  )
  source.file(
    'word/styles.xml',
    '<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="KeyNote"><w:name w:val="Key Note"/><w:basedOn w:val="Normal"/><w:rPr><w:i/></w:rPr></w:style>' +
      '</w:styles>',
  )
  source.file(
    '[Content_Types].xml',
    (await source.file('[Content_Types].xml')!.async('string')).replace(
      '</Types>',
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
    ),
  )
  source.file(
    'word/_rels/document.xml.rels',
    (await source.file('word/_rels/document.xml.rels')!.async('string')).replace(
      '</Relationships>',
      '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
    ),
  )
  return source.generateAsync({ type: 'nodebuffer' })
}

async function openDoc(docPath: string, userDataDir?: string) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
    userDataDir,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await expect(page.locator('.doc-page p', { hasText: 'Fourth line' })).toBeVisible({
    timeout: 30_000,
  })
  return { launched, page }
}

async function openCustomizeKeyboard(app: ElectronApplication, page: Page) {
  await clickMenu(app, ['Tools', 'Customize Keyboard…'])
  const dialog = page.locator('.modal-customize-keyboard')
  await expect(dialog).toBeVisible()
  await expect(dialog.locator('[data-command="font-dialog"]')).toBeAttached()
  return dialog
}

/** the font the page draws the first text of a paragraph in */
const drawnFont = (page: Page, text: string) =>
  page.evaluate((text) => {
    const block = [...document.querySelectorAll('.doc-page p, .doc-page h1')].find(
      (el) => el.textContent === text,
    )
    const walker = block && document.createTreeWalker(block, NodeFilter.SHOW_TEXT)
    const parent = walker?.nextNode()?.parentElement
    return parent ? getComputedStyle(parent).fontFamily : ''
  }, text)

const drawnSizePx = (page: Page, text: string) =>
  page.evaluate((text) => {
    const block = [...document.querySelectorAll('.doc-page p, .doc-page h1')].find(
      (el) => el.textContent === text,
    )
    const walker = block && document.createTreeWalker(block, NodeFilter.SHOW_TEXT)
    const parent = walker?.nextNode()?.parentElement
    return parent ? parseFloat(getComputedStyle(parent).fontSize) : 0
  }, text)

test('a font, a style and a font size given a chord in Customize Keyboard apply in the document and survive a restart', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-ck-values-')))
  const docPath = join(dir, 'values.docx')
  writeFileSync(docPath, await fixture())
  const first = await openDoc(docPath)
  const { userDataDir } = first.launched
  const stored = () => {
    const file = join(userDataDir, 'docs-keyboard.json')
    return existsSync(file) ? (JSON.parse(readFileSync(file, 'utf8')) as object) : null
  }
  try {
    const { app } = first.launched
    const page = first.page
    const dialog = await openCustomizeKeyboard(app, page)
    const press = dialog.locator('#ck-press')
    const note = dialog.locator('.ck-note')
    const assign = dialog.getByRole('button', { name: 'Assign' })
    const currentKeys = dialog.locator('.ck-current button')
    const command = (id: string) => dialog.locator(`[data-command="${id}"]`)

    // Fonts lists fonts, each one a command
    await dialog.locator('[data-category="fonts"]').click()
    await expect(command('font-dialog')).toHaveCount(0)
    await command('font-family:Georgia').click()
    await expect(command('font-family:Georgia')).toHaveAttribute('aria-selected', 'true')
    await expect(assign).toBeDisabled()
    await press.press('Control+Shift+F6')
    await expect(note).toHaveText('[unassigned]')
    await assign.click()
    await expect(currentKeys).toHaveText([shown('⌃⇧F6', 'Ctrl+Shift+F6')])

    // the chord names the font that has it, and moves to another font on Assign
    await command('font-family:Verdana').click()
    await press.press('Control+Shift+F6')
    await expect(note).toHaveText('Currently assigned to: Georgia')
    await press.press('Control+Shift+F5')
    await assign.click()
    await expect(currentKeys).toHaveText([shown('⌃⇧F5', 'Ctrl+Shift+F5')])

    // Styles lists the styles of the document
    await dialog.locator('[data-category="styles"]').click()
    await expect(command('apply-style:KeyNote')).toHaveText(/Key Note/)
    await command('apply-style:Heading1').click()
    await expect(command('apply-style:Heading1')).toHaveText(/Heading 1/)
    await press.press('Control+Shift+F7')
    await assign.click()
    await expect(currentKeys).toHaveText([shown('⌃⇧F7', 'Ctrl+Shift+F7')])

    // Font Size takes its value from a second list; each value has its own keys
    await dialog.getByRole('option', { name: 'All Commands' }).click()
    await command('font-size').click()
    const value = dialog.locator('.ck-value')
    await expect(value).toBeVisible()
    await value.selectOption('24')
    await press.press('Control+Shift+F8')
    await assign.click()
    await expect(currentKeys).toHaveText([shown('⌃⇧F8', 'Ctrl+Shift+F8')])
    await value.selectOption('12')
    await expect(currentKeys).toHaveCount(0)
    await value.selectOption('24')
    await expect(currentKeys).toHaveText([shown('⌃⇧F8', 'Ctrl+Shift+F8')])

    // a command with no value shows no second list
    await command('strikethrough').click()
    await expect(value).toHaveCount(0)

    await expect.poll(stored).toEqual({
      'font-family:Georgia': ['Ctrl+Shift+F6'],
      'font-family:Verdana': ['Ctrl+Shift+F5'],
      'apply-style:Heading1': ['Ctrl+Shift+F7'],
      'font-size:24': ['Ctrl+Shift+F8'],
    })
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect(dialog).toBeHidden()

    // the closing dialog hands the focus back a moment later, which can drop a selection made at once
    await expect(async () => {
      await selectLine(page, 'First line')
      await page.keyboard.press('Control+Shift+F6')
      expect(await drawnFont(page, 'First line')).toContain('Georgia')
    }).toPass({ timeout: 15_000 })
    expect(await drawnFont(page, 'Fourth line')).not.toContain('Georgia')

    await selectLine(page, 'Second line')
    await page.keyboard.press('Control+Shift+F7')
    await expect(page.locator('.doc-page h1')).toHaveText('Second line')

    const before = await drawnSizePx(page, 'Fourth line')
    await selectLine(page, 'Third line')
    await page.keyboard.press('Control+Shift+F8')
    await expect.poll(() => drawnSizePx(page, 'Third line')).toBeGreaterThan(before * 1.5)

    // Help > Keyboard Shortcuts lists them by value
    await clickMenu(app, ['Help', 'Keyboard Shortcuts'])
    const panel = page.locator('.modal-shortcuts')
    await expect(panel.locator('.sc-row', { hasText: 'Georgia' }).locator('kbd')).toHaveText(
      shown('⌃⇧F6', 'Ctrl+Shift+F6'),
    )
    await expect(panel.locator('.sc-row', { hasText: 'Font Size: 24' }).locator('kbd')).toHaveText(
      shown('⌃⇧F8', 'Ctrl+Shift+F8'),
    )
    await page.keyboard.press('Escape')
    await expect(panel).toBeHidden()

    await saveFromMenu(app)
    await expect
      .poll(() => savedParagraph(docPath, 'First line'))
      .toMatch(/<w:rFonts [^>]*w:ascii="Georgia"/)
    expect(await savedParagraph(docPath, 'Second line')).toContain('<w:pStyle w:val="Heading1"/>')
    expect(await savedParagraph(docPath, 'Third line')).toContain('<w:sz w:val="48"/>')
    expect(await savedParagraph(docPath, 'Fourth line')).not.toMatch(/Georgia|Heading1|w:sz /)
  } finally {
    await closeAndSaveVideo(first.launched, `${VIDEO_DIR}-assign`).catch(() => {})
  }

  const second = await openDoc(docPath, userDataDir)
  try {
    const { app } = second.launched
    const page = second.page

    // after a restart the chords still apply their values
    await expect(async () => {
      await selectLine(page, 'Fourth line')
      await page.keyboard.press('Control+Shift+F5')
      expect(await drawnFont(page, 'Fourth line')).toContain('Verdana')
    }).toPass({ timeout: 15_000 })
    await selectLine(page, 'Third line')
    await page.keyboard.press('Control+Shift+F7')
    await expect(page.locator('.doc-page h1')).toHaveText(['Second line', 'Third line'])

    // and the dialog shows them on the font and the style that hold them
    const dialog = await openCustomizeKeyboard(app, page)
    const currentKeys = dialog.locator('.ck-current button')
    await dialog.locator('[data-category="fonts"]').click()
    await expect(dialog.locator('[data-command="font-family:Georgia"] kbd')).toHaveText(
      shown('⌃⇧F6', 'Ctrl+Shift+F6'),
    )
    await dialog.locator('[data-category="styles"]').click()
    await dialog.locator('[data-command="apply-style:Heading1"]').click()
    await expect(currentKeys).toHaveText([shown('⌃⇧F7', 'Ctrl+Shift+F7')])

    // Remove takes the key from the style
    await currentKeys.first().click()
    await dialog.getByRole('button', { name: 'Remove' }).click()
    await expect(currentKeys).toHaveCount(0)
    await expect.poll(stored).toEqual({
      'font-family:Georgia': ['Ctrl+Shift+F6'],
      'font-family:Verdana': ['Ctrl+Shift+F5'],
      'font-size:24': ['Ctrl+Shift+F8'],
    })
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect(dialog).toBeHidden()

    await saveFromMenu(app)
    await expect
      .poll(() => savedParagraph(docPath, 'Fourth line'))
      .toMatch(/<w:rFonts [^>]*w:ascii="Verdana"/)
    expect(await savedParagraph(docPath, 'Third line')).toContain('<w:pStyle w:val="Heading1"/>')
  } finally {
    await closeAndSaveVideo(second.launched, `${VIDEO_DIR}-restart`).catch(() => {})
    rmSync(dir, { recursive: true, force: true })
  }
})
