/**
 * Design tab: the Themes gallery is a grid of Word's Office themes, a pick
 * repaints tables, shapes and charts at once and equals save and reopen, the
 * Design picks are undo steps, and style sets can be saved, reset and deleted.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, type LaunchedApp } from './helpers'
import { saveFromMenu } from './docs-objects-fixture'
import { themedDocx } from '../packages/docx-engine/tests/helpers/themed-docx'

const VIDEO_DIR = 'docs-design-themes'

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

/** Edit's second item: Repeat under whichever wording it has now */
const repeatLabel = (app: ElectronApplication) =>
  app.evaluate(
    ({ Menu }) =>
      Menu.getApplicationMenu()?.items.find((item) => item.label === 'Edit')?.submenu?.items[1]
        ?.label,
  )

/** what the theme decides on the page: table style, a cell's own fill, the text box, the chart, the faces */
const look = (page: Page) =>
  page.evaluate(() => {
    const css = (el: Element | null | undefined, prop: string) =>
      el ? getComputedStyle(el).getPropertyValue(prop) : null
    const cells = [
      ...document.querySelectorAll('.doc-page table:not(.doc-chart-table) :is(td, th)'),
    ]
    const cell = (text: string) => cells.find((c) => c.textContent?.trim() === text)
    const box = document.querySelector('.doc-page .doc-textbox')
    const line = (text: string) =>
      [...document.querySelectorAll('.doc-page :is(h1, p)')].find((p) =>
        p.textContent?.includes(text),
      )
    const run = (text: string) =>
      [...(line(text)?.querySelectorAll('span') ?? [])].find((s) => s.textContent?.includes(text))
    const bars = [...document.querySelectorAll('.doc-page .doc-chart-canvas svg rect[fill]')]
      .map((r) => r.getAttribute('fill')!.toUpperCase())
      .filter((fill) => /^#[0-9A-F]{6}$/.test(fill) && fill !== '#FFFFFF')
    return {
      headFill: css(cell('Region'), 'background-color'),
      headInk: css(cell('Region')?.querySelector('p'), 'color'),
      bandFill: css(cell('North'), 'background-color'),
      rule: css(cell('North'), 'border-top-color'),
      ownFill: css(cell('20'), 'background-color'),
      boxFill: css(box, 'background-color'),
      boxLine: css(box, 'border-top-color'),
      boxInk: css(box?.querySelector('span'), 'color'),
      bars: [...new Set(bars)],
      heading: css(line('Overview'), 'color'),
      headingFace: css(line('Overview'), 'font-family'),
      displayFace: css(run('display face'), 'font-family'),
      bodyFace: css(line('plain'), 'font-family'),
    }
  })

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

/** clicking the tab that is already open would fold the ribbon away */
async function designTab(page: Page) {
  const tab = page.locator('.ribbon-tab', { hasText: /^Design$/ })
  if (!(await tab.evaluate((el) => el.classList.contains('active')))) await tab.click()
}
const bigButton = (page: Page, label: string) =>
  page.locator('.rb-big', { has: page.getByText(label, { exact: true }) })

const OFFICE = {
  headFill: 'rgb(68, 114, 196)',
  ownFill: 'rgb(112, 173, 71)',
  boxFill: 'rgb(237, 125, 49)',
  boxLine: 'rgb(68, 114, 196)',
  heading: 'rgb(68, 114, 196)',
}
const FACET = {
  headFill: 'rgb(144, 194, 38)',
  ownFill: 'rgb(145, 134, 85)',
  boxFill: 'rgb(84, 160, 33)',
  boxLine: 'rgb(144, 194, 38)',
  heading: 'rgb(144, 194, 38)',
}

test('Design > Themes: the grid of Office themes; a pick repaints table, shape and chart, undoes, and equals save and reopen', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-themes-')))
  const userDataDir = mkdtempSync(join(dir, 'user-data'))
  const docPath = join(dir, 'themed.docx')
  writeFileSync(docPath, await themedDocx())
  let launched = await open(userDataDir, docPath)
  try {
    let { app } = launched
    let page = await waitForPageWithUrl(app, '://docs/')
    await expect.poll(() => look(page)).toMatchObject(OFFICE)
    const before = await look(page)
    expect(before.bars).toEqual(expect.arrayContaining(['#4472C4', '#ED7D31']))
    await page.screenshot({ path: test.info().outputPath('page-office.png') })

    await designTab(page)
    await bigButton(page, 'Themes').click()
    const gallery = page.locator('.theme-gallery')
    await expect(gallery.locator('.theme-card')).toHaveCount(29)
    await expect(gallery.locator('.theme-card.active .theme-name')).toHaveText(
      'Office 2013 - 2022 Theme',
    )
    // a card draws its own theme: Facet's faces and first accent
    const facet = gallery.getByTestId('theme-facet')
    await expect(facet.locator('.theme-thumb > span').first()).toHaveCSS(
      'font-family',
      /Trebuchet MS/,
    )
    await expect(facet.locator('.theme-strip > span').first()).toHaveCSS(
      'background-color',
      'rgb(144, 194, 38)',
    )
    // the cards are reachable from the keyboard: arrows walk the grid, Enter picks
    await gallery.locator('.theme-card').first().focus()
    await page.keyboard.press('ArrowRight')
    await expect(gallery.locator('.theme-card').nth(1)).toBeFocused()
    await page.keyboard.press('ArrowDown')
    await expect(gallery.locator('.theme-card').nth(5)).toBeFocused()
    await page.screenshot({ path: test.info().outputPath('themes-gallery.png') })
    await facet.focus()
    await page.keyboard.press('Enter')
    await expect(gallery).toBeHidden()

    await expect.poll(() => look(page)).toMatchObject(FACET)
    const live = await look(page)
    expect(live.bars).toEqual(expect.arrayContaining(['#90C226', '#54A021']))
    expect(live.bars).not.toContain('#4472C4')
    for (const face of [live.headingFace, live.displayFace, live.bodyFace])
      expect(face).toContain('Trebuchet MS')
    expect(live.bandFill).not.toBe(before.bandFill)
    expect(live.rule).not.toBe(before.rule)
    await page.screenshot({ path: test.info().outputPath('page-facet.png') })

    // one undo step, and Redo puts the theme back
    await expect.poll(() => repeatLabel(app)).toBe('Repeat')
    await clickMenu(app, ['Edit', 'Undo'])
    await expect.poll(() => look(page)).toEqual(before)
    await clickMenu(app, ['Edit', 'Redo'])
    await expect.poll(() => look(page)).toEqual(live)
    await bigButton(page, 'Themes').click()
    await expect(gallery.locator('.theme-card.active .theme-name')).toHaveText('Facet')
    await page.keyboard.press('Escape')

    await saveFromMenu(app)
    await expect
      .poll(async () => {
        const zip = await JSZip.loadAsync(readFileSync(docPath))
        const theme = (await zip.file('word/theme/theme1.xml')?.async('string')) ?? ''
        const body = (await zip.file('word/document.xml')?.async('string')) ?? ''
        return [
          /<a:accent1>\s*<a:srgbClr val="90C226"/.test(theme),
          /<a:theme [^>]*name="Facet"/.test(theme),
          theme.includes('<a:lin ang="5400000" scaled="0"/>'),
          body.includes('w:themeFill="accent6"'),
          body.includes('<a:schemeClr val="accent2"/>'),
          body.includes('w:asciiTheme="majorHAnsi"'),
        ]
      })
      .toEqual([true, true, true, true, true, true])
    await expect.poll(() => look(page)).toEqual(live)

    // the file reads back the way the page already looked
    await closeAndSaveVideo(launched, `${VIDEO_DIR}-pick`)
    launched = await open(userDataDir, docPath)
    app = launched.app
    page = await waitForPageWithUrl(app, '://docs/')
    await expect.poll(() => look(page)).toEqual(live)
  } finally {
    await closeAndSaveVideo(launched, `${VIDEO_DIR}-reopen`)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Design > Colors, Fonts, Paragraph Spacing and a style set are undo steps; Thatch keeps its outlined title', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-themes-')))
  const docPath = join(dir, 'themed.docx')
  writeFileSync(docPath, await themedDocx())
  const launched = await open(mkdtempSync(join(dir, 'user-data')), docPath)
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const before = await look(page)
    await designTab(page)

    await bigButton(page, 'Colors').click()
    const colors = page.locator('.theme-list')
    await expect(colors.locator('button')).toHaveCount(24)
    await expect(colors.locator('button.active')).toHaveText('Office 2013 - 2022')
    await colors.getByTestId('theme-colors-blueWarm').click()
    await expect
      .poll(() => look(page))
      .toMatchObject({
        headFill: 'rgb(74, 102, 172)',
        boxFill: 'rgb(98, 157, 209)',
        headingFace: before.headingFace,
      })

    await bigButton(page, 'Fonts').click()
    const fonts = page.locator('.theme-list')
    await expect(fonts.locator('button')).toHaveCount(26)
    await fonts.getByTestId('theme-fonts-georgia').click()
    await expect.poll(async () => (await look(page)).bodyFace).toContain('Georgia')
    await expect.poll(async () => (await look(page)).displayFace).toContain('Georgia')

    await clickMenu(app, ['Edit', 'Undo'])
    await expect.poll(async () => (await look(page)).bodyFace).toBe(before.bodyFace)
    expect((await look(page)).headFill).toBe('rgb(74, 102, 172)')
    await clickMenu(app, ['Edit', 'Undo'])
    await expect.poll(() => look(page)).toEqual(before)
    await clickMenu(app, ['Edit', 'Redo'])
    await expect.poll(async () => (await look(page)).headFill).toBe('rgb(74, 102, 172)')
    await clickMenu(app, ['Edit', 'Undo'])
    await expect.poll(() => look(page)).toEqual(before)

    // Paragraph Spacing
    const plain = page.locator('.doc-page p', { hasText: 'plain' }).first()
    const gap = () => plain.evaluate((el) => getComputedStyle(el).lineHeight)
    const single = await gap()
    await page.locator('.rb-small', { hasText: 'Paragraph Spacing' }).click()
    await page.locator('[data-rb-panel] button', { hasText: 'Double' }).click()
    await expect.poll(gap).not.toBe(single)
    await expect.poll(() => repeatLabel(app)).toBe('Repeat')
    await clickMenu(app, ['Edit', 'Undo'])
    await expect.poll(gap).toBe(single)

    // a style set
    const heading = page.locator('.doc-page :is(h1, p)', { hasText: 'Overview' }).first()
    const caps = () => heading.evaluate((el) => getComputedStyle(el).textTransform)
    await page.getByTestId('style-set-more').click()
    await page.locator('.style-set-menu').getByTestId('style-set-modern').click()
    await expect.poll(caps).toBe('uppercase')
    await clickMenu(app, ['Edit', 'Undo'])
    await expect.poll(caps).toBe('none')
    await clickMenu(app, ['Edit', 'Redo'])
    await expect.poll(caps).toBe('uppercase')
    await clickMenu(app, ['Edit', 'Undo'])
    await expect.poll(caps).toBe('none')

    // Thatch: Word's white title with an outline in the first accent, on the page and in the file
    const title = page.locator('.doc-page p', { hasText: 'Report' }).first()
    const stroke = () =>
      title.evaluate((el) => {
        const s = getComputedStyle(el.querySelector('span') ?? el)
        return [s.webkitTextStrokeWidth, s.color]
      })
    await page.getByTestId('style-set-more').click()
    await page.locator('.style-set-menu').getByTestId('style-set-thatch').click()
    await expect.poll(async () => parseFloat((await stroke())[0]!)).toBeGreaterThan(1)
    expect((await stroke())[1]).toBe('rgb(255, 255, 255)')
    await page.screenshot({ path: test.info().outputPath('thatch.png') })
    await saveFromMenu(app)
    await expect
      .poll(async () => {
        const zip = await JSZip.loadAsync(readFileSync(docPath))
        const styles = (await zip.file('word/styles.xml')?.async('string')) ?? ''
        return [
          /<w:styles\b[^>]*xmlns:w14=/.test(styles),
          /w:styleId="Title"[\s\S]*?<w14:textOutline w14:w="13335"/.test(styles),
        ]
      })
      .toEqual([true, true])
  } finally {
    await closeAndSaveVideo(launched, `${VIDEO_DIR}-undo`)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Design > style sets: Save as a New Style Set survives a restart, applies, deletes; Reset to the Default Style Set; Set as Default', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-themes-')))
  const userDataDir = mkdtempSync(join(dir, 'user-data'))
  const docPath = join(dir, 'themed.docx')
  writeFileSync(docPath, await themedDocx())
  let launched = await open(userDataDir, docPath)
  try {
    let { app } = launched
    let page = await waitForPageWithUrl(app, '://docs/')
    const heading = () => page.locator('.doc-page :is(h1, p)', { hasText: 'Overview' }).first()
    const caps = () => heading().evaluate((el) => getComputedStyle(el).textTransform)
    const menu = () => page.locator('.style-set-menu')
    const openSets = async () => {
      await designTab(page)
      await page.getByTestId('style-set-more').click()
      await expect(menu()).toBeVisible()
    }

    await openSets()
    await expect(menu().locator('.layout-menu-head')).toHaveText(['This Document', 'Built-In'])
    // arrows cross the sections: down from the lone card to the row below it, and along that row
    await menu().getByTestId('style-set-current').focus()
    await page.keyboard.press('ArrowDown')
    await expect(menu().getByTestId('style-set-distinctive')).toBeFocused()
    await page.keyboard.press('ArrowRight')
    await page.keyboard.press('ArrowDown')
    await expect(menu().getByTestId('style-set-modern')).toBeFocused()
    await menu().getByTestId('style-set-modern').click()
    await expect.poll(caps).toBe('uppercase')

    // keep the look under a name
    await openSets()
    await menu().getByTestId('style-set-save').click()
    await expect(page.locator('input[name="name"]')).toBeFocused()
    await page.keyboard.type('House style')
    await page.keyboard.press('Enter')
    await openSets()
    await expect(menu().locator('.layout-menu-head')).toHaveText([
      'This Document',
      'Custom',
      'Built-In',
    ])
    await expect(menu().getByTestId('style-set-custom-House style')).toBeVisible()
    await page.screenshot({ path: test.info().outputPath('style-sets-custom.png') })

    // Reset: the styles the document had when it was opened
    await menu().getByTestId('style-set-reset').click()
    await expect.poll(caps).toBe('none')
    await clickMenu(app, ['Edit', 'Undo'])
    await expect.poll(caps).toBe('uppercase')
    await clickMenu(app, ['Edit', 'Redo'])
    await expect.poll(caps).toBe('none')

    // Set as Default keeps the theme and the style set for new documents
    await page.locator('.rb-big', { has: page.getByText('Themes', { exact: true }) }).click()
    await page.locator('.theme-gallery').getByTestId('theme-facet').click()
    await openSets()
    await menu().getByTestId('style-set-custom-House style').click()
    await expect.poll(caps).toBe('uppercase')
    await page.locator('.rb-small', { hasText: 'Set as Default' }).click()
    await expect(page.locator('.statusbar, .status-bar').first()).toContainText('style set')

    // after a restart the default and the saved set are still there
    await closeAndSaveVideo(launched, `${VIDEO_DIR}-sets`)
    launched = await open(userDataDir, docPath)
    app = launched.app
    page = await waitForPageWithUrl(app, '://docs/')
    await expect.poll(caps).toBe('none')

    // the saved set gives its look to the file as it was
    await openSets()
    await menu().getByTestId('style-set-custom-House style').click()
    await expect.poll(caps).toBe('uppercase')

    // right-click > Delete takes the saved set away
    await openSets()
    await menu().getByTestId('style-set-custom-House style').click({ button: 'right' })
    await page.getByTestId('style-set-delete').click()
    await expect(menu().getByTestId('style-set-custom-House style')).toHaveCount(0)
    await expect(menu().locator('.layout-menu-head')).toHaveText(['This Document', 'Built-In'])

    // with the file saved, a new blank document starts on the kept theme and style set
    await saveFromMenu(app)
    await expect
      .poll(async () => {
        const zip = await JSZip.loadAsync(readFileSync(docPath))
        return ((await zip.file('word/styles.xml')?.async('string')) ?? '').includes('<w:caps/>')
      })
      .toBe(true)
    await clickMenu(app, ['File', 'New Document'])
    const blank = async (): Promise<Page | null> => {
      for (const candidate of app.windows()) {
        if (!candidate.url().includes('://docs/')) continue
        const empty = await candidate
          .evaluate(() => {
            const paper = document.querySelector('.doc-page[contenteditable="true"]')
            return !!paper?.querySelector('p') && (paper.textContent ?? '').trim() === ''
          })
          .catch(() => false)
        if (empty) return candidate
      }
      return null
    }
    await expect.poll(async () => !!(await blank()), { timeout: 30_000 }).toBe(true)
    const fresh = (await blank())!
    const first = fresh.locator('.doc-page[contenteditable="true"] p').first()
    await expect(first).toHaveCSS('font-family', /Trebuchet MS/)
    // Modern's body text is 10 pt
    await expect(first).toHaveCSS('font-size', /^13\.3/)
    await first.click()
    await designTab(fresh)
    await bigButton(fresh, 'Themes').click()
    await expect(fresh.locator('.theme-gallery .theme-card.active .theme-name')).toHaveText('Facet')
    await fresh.keyboard.press('Escape')
  } finally {
    await closeAndSaveVideo(launched, `${VIDEO_DIR}-restart`)
    rmSync(dir, { recursive: true, force: true })
  }
})
