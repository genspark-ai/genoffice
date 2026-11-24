import { test, expect, type Page } from '@playwright/test'
import { mkdirSync, mkdtempSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import type { ElectronApplication } from 'playwright'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * Word's last rung: a group the window has no room for becomes one menu
 * button named after it, in Word's order per tab, and its controls drop down
 * from the button and run the same commands.
 */

const SHOTS = process.env.RIBBON_LADDER_SHOTS
const COMPACT_KEY = process.platform === 'darwin' ? 'Alt+Meta+k' : 'Control+Alt+k'
// off macOS Home carries the Editing group and keeps to stripping captions; its groups do not fold
const HOME_FOLDS = process.platform === 'darwin'

/** the order Word folds each tab's groups in; groups in one bracket fold at the same width */
const ORDER: Record<string, string[][]> = {
  Home: [['Paragraph'], ['Font']],
  Insert: [['Links'], ['Pages', 'Text'], ['Header & Footer'], ['Symbols']],
  Review: [['Protect'], ['Tracking'], ['Changes'], ['Comments']],
}

async function minimalDocx(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  zip.file(
    'word/document.xml',
    '<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:body></w:document>',
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function showTab(page: Page, name: string): Promise<void> {
  // by position: the tab names are translated
  const index = ['Home', 'Insert', 'Draw', 'Design', 'Layout', 'References', 'Mailings', 'Review']
  const tab = page.locator('.ribbon-tab:not(.ribbon-tab-file)').nth(index.indexOf(name))
  // clicking the active tab collapses the ribbon, so only click an inactive one
  if (!/\bactive\b/.test((await tab.getAttribute('class')) ?? '')) await tab.click()
  await expect(page.locator('[data-ribbon-body] button').first()).toBeVisible()
}

async function resize(app: ElectronApplication, page: Page, width: number): Promise<void> {
  await app.evaluate(({ BrowserWindow }, w) => {
    for (const win of BrowserWindow.getAllWindows()) win.setContentSize(w, 800)
    return w
  }, width)
  // a screen narrower than the width asked for keeps the window at its own
  await expect
    .poll(() => page.evaluate(() => window.innerWidth), { timeout: 5_000 })
    .toBe(width)
    .catch(() => {})
  // the ribbon fits itself again on the frame after the resize
  await page.evaluate(
    () => new Promise((done) => requestAnimationFrame(() => requestAnimationFrame(done))),
  )
}

const overflow = (page: Page): Promise<number> =>
  page.evaluate(() => {
    const body = document.querySelector('[data-ribbon-body]')!
    return body.scrollWidth - body.clientWidth
  })

interface Rungs {
  /** names of the groups showing as a menu button */
  menus: string[]
  /** stacks of small controls that have not given up their captions */
  captioned: number
  paneGone: boolean
  stylesButton: boolean
  /** big buttons that went straight to their icon */
  bigIcons: number
  /** room left at the end of the ribbon, negative when it scrolls */
  slack: number
}

const rungs = (page: Page): Promise<Rungs> =>
  page.evaluate(() => {
    const body = document.querySelector('[data-ribbon-body]')!
    const menus = [...body.querySelectorAll('.ribbon-group.rb-group-menu')]
    return {
      menus: menus.map(
        (group) =>
          group.querySelector('.rb-group-menu-button > span:last-child')?.textContent?.trim() ?? '',
      ),
      captioned: [...body.querySelectorAll('.rb-col')].filter(
        (stack) =>
          stack.querySelector('.rb-small-label') &&
          !stack.classList.contains('rb-col-tight') &&
          !stack.closest('.rb-group-menu'),
      ).length,
      paneGone: !!body.querySelector('.rb-styles-pane.rb-fit-gone'),
      stylesButton: !!body.querySelector('.rb-styles-gallery.rb-big'),
      bigIcons: body.querySelectorAll('.rb-big.rb-big-tight').length,
      slack: Math.round(
        getComputedStyle(body).direction === 'rtl'
          ? Math.min(...[...body.children].map((el) => el.getBoundingClientRect().left)) -
              body.getBoundingClientRect().left
          : body.getBoundingClientRect().right -
              Math.max(...[...body.children].map((el) => el.getBoundingClientRect().right)),
      ),
    }
  })

const menuButton = (page: Page, name: string) =>
  page.locator('[data-ribbon-body] .rb-group-menu-button', { hasText: new RegExp(`^${name}$`) })

const openItems = (page: Page) =>
  page.locator('[data-ribbon-body] .rb-group-menu.rb-group-open > .ribbon-group-items')

const docJson = (page: Page): Promise<string> =>
  page.evaluate(() =>
    JSON.stringify(
      (
        window as never as { __aidocs: { editor: { getJSON(): unknown } } }
      ).__aidocs.editor.getJSON(),
    ),
  )

test.describe('Docs ribbon groups become menu buttons in a narrow window', () => {
  let dir: string
  test.beforeEach(() => {
    dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-ribbon-menus-')))
  })
  test.afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  const open = async (name: string, lang?: string) => {
    const docPath = join(dir, 'menus.docx')
    writeFileSync(docPath, await minimalDocx())
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: name,
      openFile: docPath,
      ...(lang ? { lang } : {}),
    })
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await page.waitForFunction(
      () => Boolean((window as never as { __aidocs?: { editor?: unknown } }).__aidocs?.editor),
      undefined,
      { timeout: 30_000 },
    )
    await page.locator('.doc-page p').first().click()
    return { launched, page }
  }

  test('groups fold in the order Word folds them, after the smaller steps', async () => {
    test.setTimeout(240_000)
    const { launched, page } = await open('docs-ribbon-group-menus-order')
    try {
      const widths = [1700, 1500, 1360, 1240, 1120, 1000, 900, 800, 720]
      const seen: Record<string, Rungs[]> = { Home: [], Insert: [], Review: [] }
      for (const width of widths) {
        await resize(launched.app, page, width)
        for (const tab of Object.keys(ORDER)) {
          await showTab(page, tab)
          const at = await rungs(page)
          seen[tab].push(at)
          const folded = (names: string[]) => names.filter((name) => at.menus.includes(name))
          const steps = ORDER[tab]
          steps.forEach((names, i) => {
            // a group folds only once every group Word folds before it has
            if (folded(names).length === 0) return
            for (const earlier of steps.slice(0, i))
              expect(folded(earlier), `${tab} at ${width}: ${at.menus}`).toEqual(earlier)
          })
          const known = steps.flat()
          expect(
            at.menus.filter((name) => !known.includes(name)),
            `${tab} at ${width}`,
          ).toEqual([])
          // the ribbon scrolls only when every group that can fold has
          if (at.menus.length < known.length && (tab !== 'Home' || HOME_FOLDS))
            await expect
              .poll(() => overflow(page), { message: `${tab} at ${width}` })
              .toBeLessThanOrEqual(0)
          if (tab === 'Home' && at.menus.length > 0) {
            // the cards became the Styles button and Styles Pane left before any group folded
            expect(at.stylesButton, `Home at ${width}`).toBe(true)
            expect(at.paneGone, `Home at ${width}`).toBe(true)
          }
          // on Review no group folds while a stack still has captions to give up
          if (tab === 'Review' && at.menus.length > 0)
            expect(at.captioned, `Review at ${width}`).toBe(0)
          if (SHOTS) {
            mkdirSync(SHOTS, { recursive: true })
            await page
              .locator('.ribbon')
              .first()
              .screenshot({ path: join(SHOTS, `${tab.toLowerCase()}-${width}.png`) })
          }
        }
      }
      if (SHOTS) writeFileSync(join(SHOTS, 'menus.json'), JSON.stringify({ widths, seen }, null, 2))
      // Layout and Mailings: the two big buttons go to icons, and only then the group folds
      const twoIcons: Record<string, string> = { Layout: 'Arrange', Mailings: 'Create' }
      const fine: Record<string, Rungs[]> = {}
      for (const [tab, group] of Object.entries(twoIcons)) {
        await showTab(page, tab)
        const states: Rungs[] = []
        for (let width = 1700; width >= 720; width -= 10) {
          await resize(launched.app, page, width)
          states.push(await rungs(page))
        }
        fine[tab] = states
        if (SHOTS) writeFileSync(join(SHOTS, 'fine.json'), JSON.stringify(fine, null, 2))
        const firstMenu = states.findIndex((at) => at.menus.includes(group))
        const firstIcons = states.findIndex((at) => at.bigIcons > 0)
        expect(firstIcons, `${tab}: big buttons as icons`).toBeGreaterThanOrEqual(0)
        expect(firstMenu, `${tab}: ${group} as a menu`).toBeGreaterThanOrEqual(firstIcons)
        // inside Arrange the two get their captions back in the menu; beside Create they stay icons
        if (tab === 'Layout') expect(firstMenu).toBeGreaterThan(firstIcons)
        else for (const at of states.slice(firstMenu)) expect(at.bigIcons, tab).toBe(2)
        expect(states[0].bigIcons + states[0].menus.length, `${tab} at its widest`).toBe(0)
      }
      for (const [tab, list] of Object.entries(seen)) {
        // a group that folded stays folded as the window narrows
        list.forEach((at, i) => {
          if (i > 0)
            for (const name of list[i - 1].menus) expect(at.menus, `${tab} #${i}`).toContain(name)
        })
        expect(list[0].menus, `${tab} at its widest`).toEqual([])
        if (tab !== 'Home' || HOME_FOLDS)
          expect(list.at(-1)!.menus.length, `${tab} at its narrowest`).toBeGreaterThan(0)
        else expect(list.at(-1)!.menus, 'Home off macOS').toEqual([])
      }
    } finally {
      await closeAndSaveVideo(launched, 'docs-ribbon-group-menus-order')
    }
  })

  test('a folded group opens from its button and runs its commands', async () => {
    test.setTimeout(240_000)
    const { launched, page } = await open('docs-ribbon-group-menus-run')
    try {
      await resize(launched.app, page, 720)
      const inWindow = async () => {
        const box = (await openItems(page).boundingBox())!
        expect(box.x).toBeGreaterThanOrEqual(0)
        expect(box.x + box.width).toBeLessThanOrEqual(await page.evaluate(() => window.innerWidth))
      }

      await showTab(page, 'Insert')
      const pages = menuButton(page, 'Pages')
      await expect(pages).toBeVisible()
      const pageBreak = page
        .locator('[data-ribbon-body]')
        .getByRole('button', { name: 'Page Break' })
      await expect(pageBreak).toBeHidden()
      await pages.click()
      await expect(openItems(page)).toBeVisible()
      await inWindow()
      const beforeBreak = await docJson(page)
      await pageBreak.click()
      await expect.poll(() => docJson(page)).not.toBe(beforeBreak)
      // a command closes the menu
      await expect(openItems(page)).toHaveCount(0)

      // a dropdown of a control inside still opens, and Escape closes one layer at a time
      await menuButton(page, 'Text').click()
      await expect(openItems(page)).toBeVisible()
      await inWindow()
      await openItems(page).locator('.rb-split-wrap > button').first().click()
      const panel = openItems(page).locator('[data-rb-panel]')
      await expect(panel).toBeVisible()
      await page.keyboard.press('Escape')
      await expect(panel).toHaveCount(0)
      await expect(openItems(page)).toBeVisible()
      await page.keyboard.press('Escape')
      await expect(openItems(page)).toHaveCount(0)

      // a press outside closes it
      const links = menuButton(page, 'Links')
      await links.click()
      await expect(openItems(page)).toBeVisible()
      await page.locator('.doc-page p').first().click()
      await expect(openItems(page)).toHaveCount(0)

      // from the keyboard: Enter opens it with the focus inside, Escape hands the focus back
      await links.focus()
      await page.keyboard.press('Enter')
      await expect(openItems(page)).toBeVisible()
      await expect
        .poll(() =>
          page.evaluate(
            () => !!document.activeElement?.closest('.rb-group-open .ribbon-group-items'),
          ),
        )
        .toBe(true)
      await page.keyboard.press('Escape')
      await expect(openItems(page)).toHaveCount(0)
      await expect(links).toBeFocused()

      await showTab(page, 'Review')
      await menuButton(page, 'Tracking').click()
      const track = openItems(page).locator('[data-tip^="Track Changes"]').first()
      await expect(track).toBeVisible()
      await expect(track).not.toHaveClass(/\bactive\b/)
      await track.click()
      await expect(openItems(page)).toHaveCount(0)
      await menuButton(page, 'Tracking').click()
      await expect(track).toHaveClass(/\bactive\b/)

      // a tab switch leaves no menu open, and the groups fold again on the way back
      await showTab(page, 'Insert')
      await expect(openItems(page)).toHaveCount(0)
      await expect(pages).toBeVisible()

      if (HOME_FOLDS) {
        await showTab(page, 'Home')
        const center = page.locator('[data-ribbon-body] [data-tip^="Center"]').first()
        await expect(center).toBeHidden()
        await menuButton(page, 'Paragraph').click()
        await expect(openItems(page)).toBeVisible()
        await inWindow()
        if (SHOTS) {
          mkdirSync(SHOTS, { recursive: true })
          await page.screenshot({ path: join(SHOTS, 'home-720-paragraph-open.png') })
        }
        const before = await docJson(page)
        await center.click()
        await expect.poll(() => docJson(page)).not.toBe(before)
        expect(await docJson(page)).toContain('center')
        await expect(openItems(page)).toHaveCount(0)
        // the font box inside the menu still drops its list
        await menuButton(page, 'Font').click()
        await inWindow()
        await openItems(page).locator('.rb-split-wrap .rb-caret').first().click()
        await expect(openItems(page).locator('[data-rb-panel]')).toBeVisible()
        await page.keyboard.press('Escape')
        await page.keyboard.press('Escape')
        await expect(openItems(page)).toHaveCount(0)
        await showTab(page, 'Insert')
      }

      // the compact ribbon shows every control as an icon, so nothing folds
      const folded = () => page.locator('[data-ribbon-body] .rb-group-menu').count()
      const compact = () => page.evaluate(() => !!document.querySelector('.ribbon-compact'))
      for (let i = 0; i < 3 && !(await compact()); i += 1) {
        await page.keyboard.press(COMPACT_KEY)
        await page.waitForTimeout(500)
      }
      expect(await compact()).toBe(true)
      await expect.poll(folded).toBe(0)
      await expect(pageBreak).toBeVisible()
      for (let i = 0; i < 3 && (await compact()); i += 1) {
        await page.keyboard.press(COMPACT_KEY)
        await page.waitForTimeout(500)
      }
      await expect.poll(folded).toBeGreaterThan(0)
    } finally {
      await closeAndSaveVideo(launched, 'docs-ribbon-group-menus-run')
    }
  })

  test('in a right-to-left UI the menu drops from the group and stays in the window', async () => {
    test.setTimeout(180_000)
    const { launched, page } = await open('docs-ribbon-group-menus-rtl', 'ar')
    try {
      expect(await page.evaluate(() => getComputedStyle(document.body).direction)).toBe('rtl')
      await resize(launched.app, page, 720)
      const width = await page.evaluate(() => window.innerWidth)
      for (const tab of [...(HOME_FOLDS ? ['Home'] : []), 'Insert', 'Review']) {
        await showTab(page, tab)
        const buttons = page.locator('[data-ribbon-body] button.rb-group-menu-button')
        await expect.poll(() => buttons.count(), { message: tab }).toBeGreaterThan(0)
        for (let i = 0; i < (await buttons.count()); i += 1) {
          const button = buttons.nth(i)
          await button.scrollIntoViewIfNeeded()
          await button.click()
          await expect(openItems(page)).toBeVisible()
          const menu = (await openItems(page).boundingBox())!
          const from = (await button.boundingBox())!
          expect(menu.x, `${tab} #${i}`).toBeGreaterThanOrEqual(0)
          expect(menu.x + menu.width, `${tab} #${i}`).toBeLessThanOrEqual(width)
          expect(menu.y, `${tab} #${i}`).toBeGreaterThanOrEqual(from.y + from.height)
          // it hangs from the reading-direction start of its button wherever the window has the room
          if (from.x + from.width - menu.width >= 12)
            expect(
              Math.abs(menu.x + menu.width - (from.x + from.width)),
              `${tab} #${i}`,
            ).toBeLessThan(12)
          if (SHOTS && i === 0) {
            mkdirSync(SHOTS, { recursive: true })
            await page.screenshot({ path: join(SHOTS, `rtl-${tab.toLowerCase()}-open.png`) })
          }
          await page.keyboard.press('Escape')
          await expect(openItems(page)).toHaveCount(0)
        }
      }
    } finally {
      await closeAndSaveVideo(launched, 'docs-ribbon-group-menus-rtl')
    }
  })
})
