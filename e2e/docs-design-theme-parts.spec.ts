/**
 * Design > Themes outside the body: headers, footers, footnotes and endnotes
 * take a pick at once, the page equals save and reopen, and parts nobody
 * edited keep their bytes.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, type LaunchedApp } from './helpers'
import { saveFromMenu } from './docs-objects-fixture'
import {
  THEMED_PARTS,
  themedPartsDocx,
} from '../packages/docx-engine/tests/helpers/themed-parts-docx'

const VIDEO_DIR = 'docs-design-theme-parts'

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

const TEXTS = {
  headOne: 'Head one',
  fixed: 'fixed',
  firstHead: 'First head',
  headTwo: 'Head two',
  foot: 'Page foot',
  footnote: 'Foot note',
}

/**
 * what the theme decides outside the body: colour and face of each part's run, the header table's
 * rule (the endnote list at the end of the document is drawn as plain text, so it has nothing to show)
 */
const look = (page: Page) =>
  page.evaluate((texts) => {
    const root = document.querySelector('.page-wrap') ?? document.body
    const holder = (text: string): Element | null => {
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
      for (let node = walker.nextNode(); node; node = walker.nextNode()) {
        const el = node.parentElement
        if (!node.nodeValue?.includes(text) || !el) continue
        if (el.closest('.doc-page') && !el.closest('.page-gap')) continue
        if (el.getClientRects().length > 0) return el
      }
      return null
    }
    const css = (el: Element | null | undefined, prop: string) =>
      el ? getComputedStyle(el).getPropertyValue(prop) : null
    const out: Record<string, string | null> = {}
    for (const [name, text] of Object.entries(texts)) {
      const el = holder(text)
      out[name] = css(el, 'color')
      out[`${name}Face`] = css(el, 'font-family')
    }
    out.rule = css(holder('HNorth')?.closest('.page-hf-cell'), 'border-bottom-color')
    return out
  }, TEXTS)

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
  await page.locator('.doc-page p', { hasText: 'plain' }).first().click()
  return launched
}

const bigButton = (page: Page, label: string) =>
  page.locator('.rb-big', { has: page.getByText(label, { exact: true }) })

const OFFICE = {
  headOne: 'rgb(237, 125, 49)',
  fixed: 'rgb(255, 0, 0)',
  firstHead: 'rgb(68, 114, 196)',
  headTwo: 'rgb(112, 173, 71)',
  foot: 'rgb(68, 84, 106)',
  footnote: 'rgb(237, 125, 49)',
}
const FACET = {
  headOne: 'rgb(84, 160, 33)',
  fixed: 'rgb(255, 0, 0)',
  firstHead: 'rgb(144, 194, 38)',
  headTwo: 'rgb(145, 134, 85)',
  footnote: 'rgb(84, 160, 33)',
}

const savedParts = async (docPath: string) => {
  const zip = await JSZip.loadAsync(readFileSync(docPath))
  const read = (name: string) => zip.file(name)?.async('string')
  const parts: Record<string, string | undefined> = {}
  for (const name of Object.keys(THEMED_PARTS)) parts[name] = await read(name)
  return { theme: (await read('word/theme/theme1.xml')) ?? '', parts }
}

test('Design > Themes: headers, footers and footnotes follow a pick at once, undo with it, and equal save and reopen', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-theme-parts-')))
  const userDataDir = mkdtempSync(join(dir, 'user-data'))
  const docPath = join(dir, 'themed-parts.docx')
  writeFileSync(docPath, await themedPartsDocx())
  let launched = await open(userDataDir, docPath)
  try {
    let { app } = launched
    let page = await waitForPageWithUrl(app, '://docs/')
    await expect.poll(() => look(page)).toMatchObject(OFFICE)
    const before = await look(page)
    await page.locator('.page-gap-note').first().scrollIntoViewIfNeeded()
    await page.screenshot({ path: test.info().outputPath('parts-office.png') })

    const tab = page.locator('.ribbon-tab', { hasText: /^Design$/ })
    await tab.click()
    await bigButton(page, 'Themes').click()
    const gallery = page.locator('.theme-gallery')
    await gallery.getByTestId('theme-facet').click()
    await expect(gallery).toBeHidden()

    await expect.poll(() => look(page)).toMatchObject(FACET)
    const live = await look(page)
    for (const name of ['headOne', 'headTwo', 'foot', 'footnote'])
      expect(live[`${name}Face`]).toContain('Trebuchet MS')
    expect(live.foot).not.toBe(before.foot)
    expect(live.rule).not.toBe(before.rule)
    await page.locator('.page-gap-note').first().scrollIntoViewIfNeeded()
    await page.screenshot({ path: test.info().outputPath('parts-facet.png') })

    await clickMenu(app, ['Edit', 'Undo'])
    await expect.poll(() => look(page)).toEqual(before)
    await clickMenu(app, ['Edit', 'Redo'])
    await expect.poll(() => look(page)).toEqual(live)

    await saveFromMenu(app)
    await expect
      .poll(async () =>
        /<a:accent1>\s*<a:srgbClr val="90C226"/.test((await savedParts(docPath)).theme),
      )
      .toBe(true)
    // the theme part changed; the parts that reference it were not rewritten
    expect((await savedParts(docPath)).parts).toEqual(THEMED_PARTS)
    await expect.poll(() => look(page)).toEqual(live)

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
