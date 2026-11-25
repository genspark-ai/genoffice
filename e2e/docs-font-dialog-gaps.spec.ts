/**
 * Font dialog: Underline style and colour, Emphasis mark, the Advanced
 * Typography group and Default…. Each choice shows on the page, is in the
 * saved file, and reads back into the dialog; a line the dialog did not touch
 * keeps its bytes.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-font-dialog-gaps'
const KEPT_RPR = '<w:rPr><w:u w:val="dashLong" w:color="00B050" w:themeColor="accent6"/></w:rPr>'

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

const select = (page: Page, from: number, to: number) =>
  page.evaluate(
    ([from, to]) => {
      const { editor } = (
        window as never as {
          __aidocs: {
            editor: { chain(): { focus(): { setTextSelection(r: object): { run(): boolean } } } }
          }
        }
      ).__aidocs
      editor.chain().focus().setTextSelection({ from, to }).run()
    },
    [from, to],
  )

const documentXml = async (docPath: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file('word/document.xml')?.async('string') ?? ''

const css = (page: Page, selector: string, prop: string) =>
  page
    .locator(selector)
    .first()
    .evaluate((el, prop) => getComputedStyle(el).getPropertyValue(prop), prop)

test('underline style and colour, emphasis mark and OpenType features apply, save and read back', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-font-gaps-')))
  const docPath = join(dir, 'font.docx')
  writeFileSync(
    docPath,
    await docxWith(
      para('plain words 123') + `<w:p><w:r>${KEPT_RPR}<w:t>kept line</w:t></w:r></w:p>`,
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
    await expect(page.locator('.doc-page p', { hasText: 'plain words' })).toBeVisible({
      timeout: 30_000,
    })
    // a line style from the file is drawn as such, in its colour
    expect(await css(page, '.doc-page u', 'text-decoration-style')).toBe('dashed')
    expect(await css(page, '.doc-page u', 'text-decoration-color')).toBe('rgb(0, 176, 80)')

    const dialog = page.locator('.modal.font-dialog')
    const pick = async (control: string, option: string) => {
      await dialog.getByRole('button', { name: control, exact: true }).click()
      await page.getByRole('option', { name: option, exact: true }).click()
    }
    await select(page, 1, 16)
    await clickMenu(app, ['Format', 'Font…'])
    await expect(dialog).toBeVisible()
    await expect(dialog.getByLabel('Underline color')).toBeDisabled()
    await pick('Underline style', 'Double underline')
    await dialog.getByLabel('Underline color').fill('#c00000')
    await expect(dialog.getByLabel('Automatic')).not.toBeChecked()
    await pick('Emphasis mark', 'Dot')
    await dialog.screenshot({ path: test.info().outputPath('font-tab.png') })

    await dialog.getByRole('tab', { name: 'Advanced' }).click()
    await expect(dialog.getByLabel('Snap to grid when document grid is defined')).toBeChecked()
    await pick('Ligatures', 'Standard and Contextual')
    await pick('Number spacing', 'Tabular')
    await pick('Number forms', 'Old-style')
    await pick('Stylistic sets', '2')
    await dialog.getByLabel('Use contextual alternates').check()
    await dialog.screenshot({ path: test.info().outputPath('advanced-tab.png') })
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect(dialog).toBeHidden()

    const first = '.doc-page p:first-of-type'
    expect(await css(page, `${first} u`, 'text-decoration-style')).toBe('double')
    expect(await css(page, `${first} u`, 'text-decoration-color')).toBe('rgb(192, 0, 0)')
    const span = `${first} span[data-doc-style]`
    expect(await css(page, span, 'font-feature-settings')).toBe(
      '"calt", "clig", "dlig" 0, "hlig" 0, "liga", "onum", "ss02", "tnum"',
    )
    expect(await css(page, span, 'text-emphasis-style')).toContain('dot')

    await saveFromMenu(app)
    await expect
      .poll(() => documentXml(docPath), { timeout: 30_000 })
      .toContain('<w:u w:val="double" w:color="C00000"/>')
    const xml = await documentXml(docPath)
    expect(xml).toContain('<w:em w:val="dot"/>')
    expect(xml).toContain(
      '<w14:ligatures w14:val="standardContextual"/><w14:numForm w14:val="oldStyle"/>' +
        '<w14:numSpacing w14:val="tabular"/>' +
        '<w14:stylisticSets><w14:styleSet w14:id="2"/></w14:stylisticSets><w14:cntxtAlts/>',
    )
    expect(xml).toMatch(/<w:document[^>]*xmlns:w14=/)
    // the run the dialog never touched keeps its line, theme link included
    expect(xml).toContain(`<w:r>${KEPT_RPR}<w:t>kept line</w:t></w:r>`)

    // the dialog shows what the run has; back to none removes it
    await select(page, 1, 16)
    await clickMenu(app, ['Format', 'Font…'])
    await expect(dialog.getByRole('button', { name: 'Emphasis mark', exact: true })).toContainText(
      'Dot',
    )
    await expect(dialog.getByLabel('Underline color')).toHaveValue('#c00000')
    await dialog.getByRole('tab', { name: 'Advanced' }).click()
    await expect(dialog.getByRole('button', { name: 'Ligatures', exact: true })).toContainText(
      'Standard and Contextual',
    )
    await expect(dialog.getByLabel('Use contextual alternates')).toBeChecked()
    await pick('Ligatures', 'None')
    await dialog.getByLabel('Use contextual alternates').uncheck()
    await dialog.getByRole('tab', { name: 'Font', exact: true }).click()
    await pick('Underline style', '(none)')
    await pick('Emphasis mark', '(none)')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect(dialog).toBeHidden()
    await expect(page.locator(`${first} u`)).toHaveCount(0)
    await saveFromMenu(app)
    await expect
      .poll(async () => (await documentXml(docPath)).includes('w:val="double"'), {
        timeout: 30_000,
      })
      .toBe(false)
    const cleared = await documentXml(docPath)
    expect(cleared).not.toContain('<w:em ')
    expect(cleared).not.toContain('w14:ligatures')
    expect(cleared).not.toContain('w14:cntxtAlts')
    expect(cleared).toContain('<w14:numForm w14:val="oldStyle"/>')
    expect(cleared).toContain(`<w:r>${KEPT_RPR}<w:t>kept line</w:t></w:r>`)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Default… keeps the font for new blank documents', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-font-gaps-')))
  const docPath = join(dir, 'default.docx')
  writeFileSync(docPath, await docxWith(para('apple')))
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const apple = page.locator('.doc-page p', { hasText: 'apple' })
    await expect(apple).toBeVisible({ timeout: 30_000 })
    const dialog = page.locator('.modal.font-dialog')
    await select(page, 1, 6)
    await clickMenu(app, ['Format', 'Font…'])
    await dialog.getByRole('button', { name: 'Font style', exact: true }).click()
    await page.getByRole('option', { name: 'Italic', exact: true }).click()
    await dialog.getByRole('button', { name: 'Size', exact: true }).click()
    await page.getByRole('option', { name: '20', exact: true }).click()
    await dialog.locator('button[name="default"]').click()
    await expect(dialog).toBeHidden()
    await expect(page.locator('.status-msg')).toContainText('this font')
    // Default… applies to this document as OK does
    await saveFromMenu(app)
    await expect
      .poll(() => documentXml(docPath), { timeout: 30_000 })
      .toContain('<w:sz w:val="40"/>')
    expect(await documentXml(docPath)).toContain('<w:i/>')
    // the file is on disk before the save is over, and until then New Document asks about unsaved changes
    await expect(page.locator('.status-msg')).toContainText('Saved')

    // a new blank document starts in the kept font
    await clickMenu(app, ['File', 'New Document'])
    await expect(apple).toHaveCount(0, { timeout: 30_000 })
    const blank = page.locator('.doc-page p').first()
    await blank.click()
    await page.keyboard.type('fresh')
    await expect(page.locator('.doc-page p', { hasText: 'fresh' })).toHaveCSS(
      'font-style',
      'italic',
    )
    const size = await css(page, '.doc-page p', 'font-size')
    expect(Math.round(parseFloat(size) * 0.75)).toBe(20)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
