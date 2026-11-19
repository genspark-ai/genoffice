/**
 * Automatic hyphenation outside the body: text boxes (edited in place or drawn
 * as static markup), header and footer strips (every page's copy, the strip
 * being edited, print preview, the exported PDF) and notes break at the
 * document's own break points and obey Hyphenate words in CAPS, the limit on
 * consecutive hyphens and the zone. A document that does not hyphenate is
 * drawn without any of it.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { saveFromMenu } from './docs-objects-fixture'
import {
  PLACE_SELECTOR,
  hyphenationPlacesDocx,
  placeLines,
  type Place,
} from './docs-hyphenation-places-fixture'

const VIDEO_DIR = 'docs-hyphenation-places'
const OUTSIDE: Place[] = ['boxed', 'frozen', 'header', 'footer']
const NOTES: Place[] = ['footnote', 'endnote']

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

const part = async (docPath: string, name: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file(name)?.async('string') ?? ''

async function open(docPath: string) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await expect(page.locator('.doc-page p').first()).toBeVisible({ timeout: 30_000 })
  await expect(page.locator('.page-gap-hf.page-hf-header').first()).toBeAttached({
    timeout: 30_000,
  })
  return { launched, app: launched.app, page }
}

const lower = (page: Page, place: Place, scope = '', nth = 0) =>
  placeLines(page, place, scoped(place, scope), nth)
const capitals = (page: Page, place: Place, scope = '', nth = 0) =>
  placeLines(page, place.toUpperCase(), scoped(place, scope), nth)
const scoped = (place: Place, scope: string) =>
  PLACE_SELECTOR[place]
    .split(',')
    .map((selector) => `${scope} ${selector}`)
    .join(',')

/** the header on an exported PDF's first page, line by line: its lower-case paragraph, then its capitals */
async function pdfHeaderLines(file: string): Promise<{ lower: string[]; capitals: string[] }> {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs')
  const doc = await pdfjs.getDocument({ data: new Uint8Array(readFileSync(file)) }).promise
  const items = (await (await doc.getPage(1)).getTextContent()).items as Array<{
    str: string
    hasEOL: boolean
  }>
  const lines = items
    .reduce<string[]>(
      (acc, item) => {
        acc[acc.length - 1] += item.str
        if (item.hasEOL) acc.push('')
        return acc
      },
      [''],
    )
    .map((line) => line.trim())
  const from = lines.findIndex((line) => line.startsWith('header '))
  const capsFrom = lines.findIndex((line) => line.startsWith('HEADER '))
  const rest = lines
    .slice(capsFrom + 1)
    .findIndex((line) => /^(footer|body|\d* ?footnote) /.test(line))
  return {
    lower: lines.slice(from, capsFrom),
    capitals: lines.slice(capsFrom, rest < 0 ? lines.length : capsFrom + 1 + rest),
  }
}
const endsInHyphen = (line: string) => /[A-Za-z] ?[-­‐]$/.test(line)

test('text boxes, headers, footers and notes hyphenate by the rules of Tools > Hyphenation', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-hyph-places-')))
  const docPath = join(dir, 'places.docx')
  writeFileSync(docPath, await hyphenationPlacesDocx({ settings: '<w:autoHyphenation/>' }))
  const session = await open(docPath)
  try {
    const { app, page } = session
    const dialog = () => page.locator('.hyphenation-dialog')
    const openDialog = async () => {
      await clickMenu(app, ['Tools', 'Hyphenation…'])
      await expect(dialog()).toBeVisible()
    }
    const ok = async () => {
      await dialog().locator('button.primary').click()
      await expect(dialog()).toBeHidden()
    }
    const caps = () => dialog().getByLabel('Hyphenate words in CAPS')
    const limit = () => dialog().getByLabel('Limit consecutive hyphens to')
    const zone = () => dialog().getByLabel('Hyphenation zone')

    // next to no zone first: every word that fits in part breaks
    await openDialog()
    await zone().fill('0.05 in')
    await ok()

    // our break points everywhere, capitals included, and Chromium's dictionary out
    for (const place of OUTSIDE) {
      await expect
        .poll(async () => (await capitals(page, place))?.broken, { message: place })
        .toBeGreaterThan(0)
      const text = await lower(page, place)
      expect(text!.broken, place).toBeGreaterThan(0)
      expect(text!.candidates, place).toBeGreaterThan(10)
      expect(text!.hyphens, place).toBe('manual')
    }
    const noteCandidates: Partial<Record<Place, number>> = {}
    for (const place of NOTES) {
      const text = await lower(page, place)
      expect(text!.candidates, place).toBeGreaterThan(10)
      expect(text!.hyphens, place).toBe('manual')
      noteCandidates[place] = text!.candidates
    }
    // every page's copy of the strip breaks the same way: the first page's and a later one
    for (const place of ['header', 'footer'] as const) {
      const first = await lower(page, place)
      const last = await lower(page, place, '', first!.copies - 1)
      expect(first!.copies, place).toBeGreaterThan(2)
      expect([last!.lines, last!.broken, last!.height], place).toEqual([
        first!.lines,
        first!.broken,
        first!.height,
      ])
    }
    const withCaps = Object.fromEntries(
      await Promise.all(OUTSIDE.map(async (place) => [place, await lower(page, place)] as const)),
    )
    await page.screenshot({ path: test.info().outputPath('places-default.png') })

    // Hyphenate words in CAPS off: capitals stay whole in every place, the rest still breaks
    await openDialog()
    await caps().uncheck()
    await ok()
    for (const place of OUTSIDE) {
      await expect
        .poll(async () => (await capitals(page, place))?.broken, { message: place })
        .toBe(0)
      expect((await lower(page, place))!.broken, place).toBe(withCaps[place]!.broken)
    }
    for (const place of NOTES) {
      // the capitals of the note carry no break point any more
      await expect
        .poll(async () => (await lower(page, place))?.candidates, { message: place })
        .toBeLessThan(noteCandidates[place]!)
    }
    await page.screenshot({ path: test.info().outputPath('places-no-caps.png') })

    // no two lines in a row end in a hyphen
    await openDialog()
    await limit().selectOption('1')
    await ok()
    for (const place of OUTSIDE) {
      await expect
        .poll(async () => (await lower(page, place))?.longest, { message: place })
        .toBeLessThanOrEqual(1)
    }
    expect((await lower(page, 'boxed'))!.broken).toBeGreaterThan(0)
    const limited = Object.fromEntries(
      await Promise.all(
        OUTSIDE.map(async (place) => {
          const text = await lower(page, place)
          return [place, [text!.lines, text!.broken]] as const
        }),
      ),
    )

    // the strip being edited breaks like the strip at rest
    const header = page.locator('.page-wrap > .page-hf-header').first()
    await header.dblclick()
    await expect(header.locator('.page-hf-editor')).toBeVisible()
    await expect
      .poll(async () => {
        const text = await placeLines(page, 'header', '.page-hf-editor')
        return text && [text.lines, text.broken]
      })
      .toEqual(limited.header)
    const edited = await placeLines(page, 'HEADER', '.page-hf-editor')
    expect([edited!.broken, edited!.candidates, edited!.ours]).toEqual([0, 0, true])
    await page.screenshot({ path: test.info().outputPath('header-editing.png') })
    await page.keyboard.press('Escape')
    await expect(header.locator('.page-hf-editor')).toHaveCount(0)

    // print preview carries the break points and the rules. The dialog shows a zoomed copy of
    // sheets whose line rules were checked unscaled, and where text is not laid out at
    // sub-pixel positions a zoomed line can end a word earlier or later, so neither the
    // line ends nor the limit are read off it
    await clickMenu(app, ['File', 'Print…'])
    const print = page.locator('.print-dialog')
    await expect(print.locator('.print-preview-nav')).toBeVisible({ timeout: 30_000 })
    for (const place of OUTSIDE) {
      await expect
        .poll(async () => (await lower(page, place, '.print-dialog'))?.candidates, {
          message: place,
        })
        .toBeGreaterThan(10)
      const text = await lower(page, place, '.print-dialog')
      expect(text!.hyphens, place).toBe('manual')
      expect((await capitals(page, place, '.print-dialog'))!.broken, place).toBe(0)
    }
    await page.screenshot({ path: test.info().outputPath('print-preview.png') })
    await print.getByRole('button', { name: 'Cancel', exact: true }).click()

    // a zone as wide as the column: nothing hyphenates anywhere
    await openDialog()
    await limit().selectOption('0')
    await zone().fill('3 in')
    await ok()
    for (const place of OUTSIDE) {
      await expect.poll(async () => (await lower(page, place))?.broken, { message: place }).toBe(0)
    }
    await openDialog()
    await zone().fill('0.05 in')
    await ok()
    for (const place of OUTSIDE) {
      await expect
        .poll(async () => (await lower(page, place))?.broken, { message: place })
        .toBe(withCaps[place]!.broken)
    }

    // the rules are saved; the text of the parts is untouched
    await saveFromMenu(app)
    await expect
      .poll(async () => {
        const settings = await part(docPath, 'word/settings.xml')
        return [
          /<w:autoHyphenation\/>/.test(settings),
          /<w:doNotHyphenateCaps\/>/.test(settings),
          /<w:hyphenationZone w:val="72"\/>/.test(settings),
        ]
      })
      .toEqual([true, true, true])
    for (const name of ['document', 'header1', 'footer1', 'footnotes', 'endnotes']) {
      const xml = await part(docPath, `word/${name}.xml`)
      expect(xml, name).not.toMatch(/­|softHyphen/)
      expect(xml, name).toContain('internationalization')
    }

    // the PDF breaks the header where the page does
    const pdfPath = join(dir, 'places.pdf')
    await app.evaluate(({ dialog: d, shell }, out) => {
      d.showSaveDialog = async () => ({ canceled: false, filePath: out })
      shell.showItemInFolder = () => {}
    }, pdfPath)
    const onScreen = await lower(page, 'header')
    expect(
      await page.evaluate(() =>
        (window as unknown as { __exportPdf: () => Promise<boolean> }).__exportPdf(),
      ),
    ).toBe(true)
    const inPdf = await pdfHeaderLines(pdfPath)
    expect(inPdf.lower).toHaveLength(onScreen!.lines)
    expect(inPdf.lower.filter(endsInHyphen)).toHaveLength(onScreen!.broken)
    expect(inPdf.capitals.length).toBeGreaterThan(1)
    expect(inPdf.capitals.filter(endsInHyphen)).toHaveLength(0)
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a document that does not hyphenate carries nothing of it outside the body either', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-hyph-places-')))
  const docPath = join(dir, 'plain.docx')
  writeFileSync(docPath, await hyphenationPlacesDocx({ settings: '' }))
  const session = await open(docPath)
  try {
    const { app, page } = session
    const marks = () =>
      page.evaluate(() => ({
        elements: document.querySelectorAll('.doc-shy, .doc-hyph, .doc-hyph-in').length,
        // class names and inline styles of everything drawn in the places this covers
        markup: Array.from(
          document.querySelectorAll('.page-hf, .doc-textbox, .page-gap-notes, .page-notes'),
        ).filter((el) => /hyph|doc-shy|­/.test(el.outerHTML)).length,
        css: Array.from(document.querySelectorAll('style[data-doc-css]')).filter((el) =>
          /hyphens|doc-shy/.test(el.textContent ?? ''),
        ).length,
      }))
    expect(await marks()).toEqual({ elements: 0, markup: 0, css: 0 })
    for (const place of [...OUTSIDE, ...NOTES]) {
      const text = await lower(page, place)
      expect([text!.broken, text!.candidates, text!.ours], place).toEqual([0, 0, false])
    }

    // the strip editor and the preview carry nothing either
    const header = page.locator('.page-wrap > .page-hf-header').first()
    await header.dblclick()
    await expect(header.locator('.page-hf-editor')).toBeVisible()
    expect(await marks()).toEqual({ elements: 0, markup: 0, css: 0 })
    await page.keyboard.press('Escape')
    await clickMenu(app, ['File', 'Print…'])
    const print = page.locator('.print-dialog')
    await expect(print.locator('.print-preview-nav')).toBeVisible({ timeout: 30_000 })
    expect(await marks()).toEqual({ elements: 0, markup: 0, css: 0 })
    expect((await lower(page, 'header', '.print-dialog'))!.broken).toBe(0)
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
