/**
 * Design > Themes, the foot of the gallery: Save Current Theme writes an Office
 * theme file that the gallery lists under Custom, Browse for Themes takes the
 * theme of a theme file or a document, and Reset to Theme from Template puts
 * back the theme the document was opened with.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  realpathSync,
  rmSync,
  writeFileSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, type LaunchedApp } from './helpers'
import { saveFromMenu } from './docs-objects-fixture'
import { themedDocx } from '../packages/docx-engine/tests/helpers/themed-docx'
import { OFFICE_THEMES, buildThmx, themeXmlOf } from '../packages/docx-engine/src/index'

const VIDEO_DIR = 'docs-design-theme-files'

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

const repeatLabel = (app: ElectronApplication) =>
  app.evaluate(
    ({ Menu }) =>
      Menu.getApplicationMenu()?.items.find((item) => item.label === 'Edit')?.submenu?.items[1]
        ?.label,
  )

/** what the theme decides on the page */
const look = (page: Page) =>
  page.evaluate(() => {
    const css = (el: Element | null | undefined, prop: string) =>
      el ? getComputedStyle(el).getPropertyValue(prop) : null
    const cells = [
      ...document.querySelectorAll('.doc-page table:not(.doc-chart-table) :is(td, th)'),
    ]
    const cell = (text: string) => cells.find((c) => c.textContent?.trim() === text)
    const line = (text: string) =>
      [...document.querySelectorAll('.doc-page :is(h1, p)')].find((p) =>
        p.textContent?.includes(text),
      )
    return {
      headFill: css(cell('Region'), 'background-color'),
      ownFill: css(cell('20'), 'background-color'),
      boxFill: css(document.querySelector('.doc-page .doc-textbox'), 'background-color'),
      heading: css(line('Overview'), 'color'),
      bodyFace: css(line('plain'), 'font-family')?.split(',')[0]?.replace(/"/g, ''),
    }
  })

const OFFICE = {
  headFill: 'rgb(68, 114, 196)',
  ownFill: 'rgb(112, 173, 71)',
  boxFill: 'rgb(237, 125, 49)',
  heading: 'rgb(68, 114, 196)',
  bodyFace: 'Calibri',
}
const FACET = {
  headFill: 'rgb(144, 194, 38)',
  ownFill: 'rgb(145, 134, 85)',
  boxFill: 'rgb(84, 160, 33)',
  heading: 'rgb(144, 194, 38)',
  bodyFace: 'Trebuchet MS',
}
// Facet with the Blue Warm colours over it
const MINE = {
  headFill: 'rgb(74, 102, 172)',
  ownFill: 'rgb(157, 144, 160)',
  boxFill: 'rgb(98, 157, 209)',
  heading: 'rgb(74, 102, 172)',
  bodyFace: 'Trebuchet MS',
}

async function open(userDataDir: string, docPath: string): Promise<LaunchedApp> {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
    userDataDir,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await expect(page.locator('.doc-page :is(h1, p)', { hasText: 'Overview' }).first()).toBeVisible({
    timeout: 30_000,
  })
  await expect(page.locator('.doc-page .doc-chart-canvas svg rect').first()).toBeVisible()
  await page.locator('.doc-page p', { hasText: 'plain' }).first().click()
  return launched
}

const gallery = (page: Page) => page.locator('.theme-gallery')

async function openThemes(page: Page) {
  const tab = page.locator('.ribbon-tab', { hasText: /^Design$/ })
  if (!(await tab.evaluate((el) => el.classList.contains('active')))) await tab.click()
  await page.locator('.rb-big', { has: page.getByText('Themes', { exact: true }) }).click()
  await expect(gallery(page)).toBeVisible()
}

const savedTheme = async (docPath: string) =>
  (await (
    await JSZip.loadAsync(readFileSync(docPath))
  )
    .file('word/theme/theme1.xml')
    ?.async('string')) ?? ''

const themeFacts = (xml: string) => ({
  name: /<a:theme [^>]*name="([^"]*)"/.exec(xml)?.[1],
  colors: /<a:clrScheme name="([^"]*)"/.exec(xml)?.[1],
  accent1: /<a:accent1>\s*<a:srgbClr val="([0-9A-F]{6})"/.exec(xml)?.[1],
  minor: /<a:minorFont>\s*<a:latin typeface="([^"]*)"/.exec(xml)?.[1],
  // Facet's format scheme has this gradient; the Office one does not
  facetFormat: xml.includes('<a:lin ang="5400000" scaled="0"/>'),
})

const status = (page: Page) => page.locator('.statusbar, .status-bar').first()

/** copies of what the app wrote, kept for a check outside this run */
function keep(file: string, name: string) {
  const dir = process.env.KEEP_THEME_FILES_DIR
  if (!dir) return
  mkdirSync(dir, { recursive: true })
  copyFileSync(file, join(dir, name))
}

test('Design > Themes > Save Current Theme writes a theme file, lists it under Custom, and it applies to another document', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-theme-files-')))
  const userDataDir = mkdtempSync(join(dir, 'user-data'))
  const docPath = join(dir, 'themed.docx')
  const otherPath = join(dir, 'other.docx')
  writeFileSync(docPath, await themedDocx())
  writeFileSync(otherPath, await themedDocx(false))
  let launched = await open(userDataDir, docPath)
  try {
    let { app } = launched
    let page = await waitForPageWithUrl(app, '://docs/')
    await expect.poll(() => look(page)).toMatchObject(OFFICE)

    await openThemes(page)
    await expect(gallery(page).locator('.theme-card')).toHaveCount(29)
    await expect(gallery(page).locator('.layout-menu-head')).toHaveCount(0)
    await expect(gallery(page).locator('.theme-foot button')).toHaveText([
      'Reset to Theme from Template',
      'Browse for Themes…',
      'Save Current Theme…',
    ])
    // the foot is on the arrow keys' way like the cards
    await gallery(page).locator('.theme-card').last().focus()
    await page.keyboard.press('ArrowDown')
    await expect(gallery(page).getByTestId('theme-reset')).toBeFocused()
    await page.keyboard.press('ArrowDown')
    await page.keyboard.press('ArrowDown')
    await expect(gallery(page).getByTestId('theme-save')).toBeFocused()
    await gallery(page).getByTestId('theme-facet').click()
    await expect.poll(() => look(page)).toMatchObject(FACET)
    await page.locator('.rb-big', { has: page.getByText('Colors', { exact: true }) }).click()
    await page.getByTestId('theme-colors-blueWarm').click()
    await expect.poll(() => look(page)).toMatchObject(MINE)

    // the save dialog starts in the themes folder, under the theme's name
    const themesDir = await app.evaluate(({ app, dialog }) => {
      const state = globalThis as unknown as { themeSaveAsked?: string }
      const dir = `${app.getPath('userData')}/Themes`
      dialog.showSaveDialog = (async (...args: unknown[]) => {
        const options = args[args.length - 1] as { defaultPath?: string }
        state.themeSaveAsked = options.defaultPath
        return { canceled: false, filePath: `${dir}/Mine` }
      }) as typeof dialog.showSaveDialog
      return dir
    })
    await openThemes(page)
    await gallery(page).getByTestId('theme-save').press('Enter')
    await expect(status(page)).toContainText('Theme "Mine" saved')
    expect(
      await app.evaluate(
        () => (globalThis as unknown as { themeSaveAsked?: string }).themeSaveAsked,
      ),
    ).toBe(join(themesDir, 'Facet.thmx'))
    const thmxPath = join(themesDir, 'Mine.thmx')
    const thmx = await JSZip.loadAsync(readFileSync(thmxPath))
    expect(
      Object.keys(thmx.files)
        .filter((name) => !name.endsWith('/'))
        .sort(),
    ).toEqual([
      '[Content_Types].xml',
      '_rels/.rels',
      'theme/theme/_rels/themeManager.xml.rels',
      'theme/theme/theme1.xml',
      'theme/theme/themeManager.xml',
    ])
    expect(themeFacts(await thmx.file('theme/theme/theme1.xml')!.async('string'))).toEqual({
      name: 'Mine',
      colors: 'Blue Warm',
      accent1: '4A66AC',
      minor: 'Trebuchet MS',
      facetFormat: true,
    })
    keep(thmxPath, 'Mine.thmx')

    // saving a theme is not an edit of the document: the pick before it is still the last step
    await openThemes(page)
    await expect(gallery(page).locator('.layout-menu-head')).toHaveText(['Custom', 'Office'])
    await expect(gallery(page).locator('.theme-card')).toHaveCount(30)
    const card = gallery(page).getByTestId('theme-custom-Mine')
    await expect(card).toHaveClass(/active/)
    await expect(card.locator('.theme-strip > span').first()).toHaveCSS(
      'background-color',
      'rgb(74, 102, 172)',
    )
    await page.screenshot({ path: test.info().outputPath('themes-custom.png') })
    await page.keyboard.press('Escape')
    await clickMenu(app, ['Edit', 'Undo'])
    await expect.poll(() => look(page)).toMatchObject(FACET)

    // another document, after a restart: the theme is in the gallery and applies as a pick does
    await closeAndSaveVideo(launched, `${VIDEO_DIR}-save`)
    launched = await open(userDataDir, otherPath)
    app = launched.app
    page = await waitForPageWithUrl(app, '://docs/')
    const before = await look(page)
    expect(before).toMatchObject(OFFICE)
    await openThemes(page)
    await gallery(page).getByTestId('theme-custom-Mine').click()
    await expect(gallery(page)).toBeHidden()
    await expect.poll(() => look(page)).toMatchObject(MINE)
    await expect.poll(() => repeatLabel(app)).toBe('Repeat')
    await clickMenu(app, ['Edit', 'Undo'])
    await expect.poll(() => look(page)).toEqual(before)
    await clickMenu(app, ['Edit', 'Redo'])
    await expect.poll(() => look(page)).toMatchObject(MINE)
    const live = await look(page)
    await saveFromMenu(app)
    await expect
      .poll(async () => themeFacts(await savedTheme(otherPath)))
      .toEqual({
        name: 'Mine',
        colors: 'Blue Warm',
        accent1: '4A66AC',
        minor: 'Trebuchet MS',
        facetFormat: true,
      })
    keep(otherPath, 'saved-theme-applied.docx')

    await closeAndSaveVideo(launched, `${VIDEO_DIR}-apply`)
    launched = await open(userDataDir, otherPath)
    app = launched.app
    page = await waitForPageWithUrl(app, '://docs/')
    await expect.poll(() => look(page)).toEqual(live)

    // right-click > Delete takes the card and its file away
    await openThemes(page)
    await gallery(page).getByTestId('theme-custom-Mine').click({ button: 'right' })
    await page.getByTestId('theme-delete').click()
    await expect(gallery(page).locator('.theme-card')).toHaveCount(29)
    await expect(gallery(page).locator('.layout-menu-head')).toHaveCount(0)
    await expect.poll(() => existsSync(thmxPath)).toBe(false)
    await expect.poll(() => look(page)).toEqual(live)
  } finally {
    await closeAndSaveVideo(launched, `${VIDEO_DIR}-reopen`)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Design > Themes > Browse for Themes takes the theme of a theme file or a document; Reset to Theme from Template puts back the one at open', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-theme-files-')))
  const userDataDir = mkdtempSync(join(dir, 'user-data'))
  const docPath = join(dir, 'themed.docx')
  const thmxPath = join(dir, 'Facet copy.thmx')
  const sourcePath = join(dir, 'source.docx')
  const barePath = join(dir, 'bare.docx')
  const facet = OFFICE_THEMES.find((theme) => theme.id === 'facet')!
  writeFileSync(docPath, await themedDocx())
  writeFileSync(
    thmxPath,
    await buildThmx(
      themeXmlOf({
        fonts: facet.fonts,
        colors: facet.colors,
        format: { name: facet.name, xml: facet.format },
      }),
    ),
  )
  writeFileSync(barePath, await themedDocx(false))
  let launched = await open(userDataDir, docPath)
  try {
    let { app } = launched
    let page = await waitForPageWithUrl(app, '://docs/')
    await expect.poll(() => look(page)).toMatchObject(OFFICE)
    const before = await look(page)
    const browse = async (file: string | null) => {
      await app.evaluate(({ dialog }, file) => {
        dialog.showOpenDialog = (async () =>
          file
            ? { canceled: false, filePaths: [file] }
            : { canceled: true, filePaths: [] }) as typeof dialog.showOpenDialog
      }, file)
      await openThemes(page)
      await gallery(page).getByTestId('theme-browse').click()
      await expect(gallery(page)).toBeHidden()
    }

    // nothing to reset yet, a cancelled pick, and a document without a theme part: no change, no undo step
    await openThemes(page)
    await gallery(page).getByTestId('theme-reset').click()
    await browse(null)
    await browse(barePath)
    await expect(status(page)).toContainText('bare.docx has no theme')
    expect(await look(page)).toEqual(before)
    await page.locator('.doc-page p', { hasText: 'plain' }).first().click()
    await page.keyboard.type('x')
    await clickMenu(app, ['Edit', 'Undo'])
    await clickMenu(app, ['Edit', 'Undo'])
    expect(await look(page)).toEqual(before)
    await expect(page.locator('.doc-page p', { hasText: 'plain' }).first()).not.toContainText('x')

    // a theme file: one undo step, Redo, Repeat
    await browse(thmxPath)
    await expect.poll(() => look(page)).toMatchObject(FACET)
    const live = await look(page)
    await expect.poll(() => repeatLabel(app)).toBe('Repeat')
    await clickMenu(app, ['Edit', 'Undo'])
    await expect.poll(() => look(page)).toEqual(before)
    await clickMenu(app, ['Edit', 'Redo'])
    await expect.poll(() => look(page)).toEqual(live)
    await saveFromMenu(app)
    await expect
      .poll(async () => themeFacts(await savedTheme(docPath)))
      .toEqual({
        name: 'Facet',
        colors: 'Facet',
        accent1: '90C226',
        minor: 'Trebuchet MS',
        facetFormat: true,
      })
    await expect(status(page)).toContainText('Saved')
    keep(docPath, 'browsed-thmx.docx')
    copyFileSync(docPath, sourcePath)

    // the theme at open comes back, although the file on disk has the other one by now
    await openThemes(page)
    await gallery(page).getByTestId('theme-reset').click()
    await expect.poll(() => look(page)).toEqual(before)
    await clickMenu(app, ['Edit', 'Undo'])
    await expect.poll(() => look(page)).toEqual(live)
    await clickMenu(app, ['Edit', 'Redo'])
    await expect.poll(() => look(page)).toEqual(before)
    await saveFromMenu(app)
    await expect
      .poll(async () => themeFacts(await savedTheme(docPath)))
      .toEqual({
        name: 'Office Theme',
        colors: 'Office',
        accent1: '4472C4',
        minor: 'Calibri',
        facetFormat: false,
      })
    await expect(status(page)).toContainText('Saved')

    // a document gives its theme the same way
    await browse(sourcePath)
    await expect.poll(() => look(page)).toEqual(live)
    await saveFromMenu(app)
    await expect
      .poll(async () => themeFacts(await savedTheme(docPath)))
      .toMatchObject({ name: 'Facet', accent1: '90C226', facetFormat: true })

    await closeAndSaveVideo(launched, `${VIDEO_DIR}-browse`)
    launched = await open(userDataDir, docPath)
    app = launched.app
    page = await waitForPageWithUrl(app, '://docs/')
    await expect.poll(() => look(page)).toEqual(live)
    // reopened on Facet, that is now the theme Reset goes back to
    await openThemes(page)
    await gallery(page).getByTestId('theme-office').click()
    await expect.poll(async () => (await look(page)).bodyFace).not.toBe('Trebuchet MS')
    await openThemes(page)
    await gallery(page).getByTestId('theme-reset').click()
    await expect.poll(() => look(page)).toEqual(live)
  } finally {
    await closeAndSaveVideo(launched, `${VIDEO_DIR}-reopen`)
    rmSync(dir, { recursive: true, force: true })
  }
})
