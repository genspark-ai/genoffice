/**
 * Tools > Hyphenation…: words in capitals, the limit on consecutive hyphens,
 * the hyphenation zone and Manual… decide where the page breaks words, are
 * written to the file and read back from it.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import {
  docxWith,
  para,
  saveFromMenu,
  settingsXml,
  withLanguage,
  withWordPart,
} from './docs-objects-fixture'

const VIDEO_DIR = 'docs-hyphenation'
// a column under 3 in wide: long words reach the end of most lines
const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="4100" w:bottom="1440" w:left="4100" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/></w:sectPr>'
const LOWER =
  'lower internationalization of organizational responsibilities demonstrates characteristically ' +
  'disproportionate institutional representation throughout intercontinental telecommunication ' +
  'infrastructure notwithstanding considerable administrative misunderstanding concerning ' +
  'environmental sustainability recommendations and multidisciplinary investigations into ' +
  'uncharacteristically complicated organizational communication problems considering ' +
  'international transportation regulations and extraordinary governmental accountability'
const CAPS =
  'UPPER INTERNATIONALIZATION OF ORGANIZATIONAL RESPONSIBILITIES DEMONSTRATES CHARACTERISTICALLY ' +
  'DISPROPORTIONATE INSTITUTIONAL REPRESENTATION THROUGHOUT INTERCONTINENTAL TELECOMMUNICATION'

/** `lang` null: the document declares no language */
async function fixture(
  settings: string,
  body = para(LOWER) + para(CAPS),
  lang: string | null = 'en-US',
): Promise<Buffer> {
  const docx = await withWordPart(await docxWith(body + SECT_PR), 'settings', settingsXml(settings))
  return lang ? withLanguage(docx, lang) : docx
}

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

/** line ends of a paragraph: how many end in a hyphen of ours, and the longest run of such lines */
const hyphenLines = (page: Page, startsWith: string) =>
  page.evaluate((startsWith) => {
    const el = Array.from(document.querySelectorAll<HTMLElement>('.ProseMirror.doc-page p')).find(
      (p) => p.textContent?.startsWith(startsWith),
    )
    if (!el) return null
    const range = document.createRange()
    range.selectNodeContents(el)
    const tops = Array.from(
      new Set(Array.from(range.getClientRects()).map((r) => Math.round(r.top))),
    )
      .sort((a, b) => a - b)
      .filter((top, i, all) => i === 0 || top - all[i - 1] > 4)
    const lineOf = (top: number) => tops.findIndex((t) => Math.abs(t - top) <= 4)
    const broken = Array.from(el.querySelectorAll('.doc-shy'))
      .map((s) => s.getBoundingClientRect())
      .filter((r) => r.width > 0.01)
      .map((r) => lineOf(Math.round(r.top)))
    let longest = 0
    let run = 0
    broken.forEach((line, i) => {
      run = i > 0 && line === broken[i - 1] + 1 ? run + 1 : 1
      longest = Math.max(longest, run)
    })
    return { lines: tops.length, hyphens: broken.length, longest }
  }, startsWith)

async function open(docPath: string) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await expect(page.locator('.doc-page p').first()).toBeVisible({ timeout: 30_000 })
  return { launched, app: launched.app, page }
}

/** line-end hyphens in the text of an exported PDF */
async function pdfLineEndHyphens(file: string): Promise<number> {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs')
  const doc = await pdfjs.getDocument({ data: new Uint8Array(readFileSync(file)) }).promise
  let count = 0
  for (let n = 1; n <= doc.numPages; n++) {
    const items = (await (await doc.getPage(n)).getTextContent()).items as Array<{
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
    count += lines.filter((line) => /[A-Za-z] ?[-\u00ad\u2010]$/.test(line)).length
  }
  return count
}

test('words in capitals, the consecutive limit and the zone change where the page hyphenates and are saved', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-hyphenation-')))
  const docPath = join(dir, 'hyphenation.docx')
  writeFileSync(docPath, await fixture('<w:autoHyphenation/>'))
  let session = await open(docPath)
  try {
    let { app, page } = session
    const dialog = () => page.locator('.hyphenation-dialog')
    const openDialog = async () => {
      await clickMenu(app, ['Tools', 'Hyphenation…'])
      await expect(dialog()).toBeVisible()
    }
    const ok = async () => {
      await dialog().locator('button.primary').click()
      await expect(dialog()).toBeHidden()
    }

    // Word's default: words in capitals hyphenate like any other
    await expect.poll(async () => (await hyphenLines(page, 'UPPER'))?.hyphens).toBeGreaterThan(1)
    const lower = await hyphenLines(page, 'lower')
    expect(lower!.hyphens).toBeGreaterThan(3)
    expect(lower!.longest).toBeGreaterThan(1)
    await page.screenshot({ path: test.info().outputPath('page-default.png') })

    // Hyphenate words in CAPS off: the capitals stay whole, the rest still breaks
    await openDialog()
    const caps = () => dialog().getByLabel('Hyphenate words in CAPS')
    const limit = () => dialog().getByLabel('Limit consecutive hyphens to')
    const zone = () => dialog().getByLabel('Hyphenation zone')
    await expect(caps()).toBeChecked()
    await expect(limit()).toHaveValue('0')
    await expect(zone()).toBeEnabled()
    await page.screenshot({ path: test.info().outputPath('dialog.png') })
    await caps().uncheck()
    await ok()
    await expect.poll(async () => (await hyphenLines(page, 'UPPER'))?.hyphens).toBe(0)
    expect((await hyphenLines(page, 'lower'))!.hyphens).toBe(lower!.hyphens)

    // no two lines in a row end in a hyphen
    await openDialog()
    await expect(caps()).not.toBeChecked()
    await limit().selectOption('1')
    await ok()
    await expect.poll(async () => (await hyphenLines(page, 'lower'))?.longest).toBe(1)
    const limited = await hyphenLines(page, 'lower')
    expect(limited!.hyphens).toBeGreaterThan(0)
    expect(limited!.hyphens).toBeLessThan(lower!.hyphens)
    await page.screenshot({ path: test.info().outputPath('page-limit.png') })

    await saveFromMenu(app)
    await expect
      .poll(async () => {
        const settings = await part(docPath, 'word/settings.xml')
        return [
          /<w:autoHyphenation\/>/.test(settings),
          /<w:consecutiveHyphenLimit w:val="1"\/>/.test(settings),
          /<w:doNotHyphenateCaps\/>/.test(settings),
          /<w:hyphenationZone/.test(settings),
        ]
      })
      .toEqual([true, true, true, false])

    // read back: the file's rules hold on the page and show in the dialog
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    session = await open(docPath)
    ;({ app, page } = session)
    await expect
      .poll(async () => (await hyphenLines(page, 'lower'))?.hyphens)
      .toBe(limited!.hyphens)
    expect((await hyphenLines(page, 'lower'))!.longest).toBe(1)
    expect((await hyphenLines(page, 'UPPER'))!.hyphens).toBe(0)
    await openDialog()
    await expect(caps()).not.toBeChecked()
    await expect(limit()).toHaveValue('1')

    // a zone as wide as the column: no gap is wider, nothing hyphenates
    await limit().selectOption('0')
    await zone().fill('3 in')
    await ok()
    await expect.poll(async () => (await hyphenLines(page, 'lower'))?.hyphens).toBe(0)
    // a narrow zone lets the words break again
    await openDialog()
    await zone().fill('0.1 in')
    await ok()
    await expect.poll(async () => (await hyphenLines(page, 'lower'))?.hyphens).toBeGreaterThan(3)

    await saveFromMenu(app)
    await expect
      .poll(async () => {
        const settings = await part(docPath, 'word/settings.xml')
        return [
          /<w:consecutiveHyphenLimit/.test(settings),
          /<w:hyphenationZone w:val="144"\/>/.test(settings),
          /<w:doNotHyphenateCaps\/>/.test(settings),
        ]
      })
      .toEqual([false, true, true])

    // the PDF breaks the same words
    const pdfPath = join(dir, 'hyphenation.pdf')
    await app.evaluate(({ dialog: d, shell }, out) => {
      d.showSaveDialog = async () => ({ canceled: false, filePath: out })
      shell.showItemInFolder = () => {}
    }, pdfPath)
    const onScreen =
      (await hyphenLines(page, 'lower'))!.hyphens + (await hyphenLines(page, 'UPPER'))!.hyphens
    expect(onScreen).toBeGreaterThan(3)
    expect(
      await page.evaluate(() =>
        (window as unknown as { __exportPdf: () => Promise<boolean> }).__exportPdf(),
      ),
    ).toBe(true)
    expect(await pdfLineEndHyphens(pdfPath)).toBe(onScreen)
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a document that does not hyphenate carries no break points; the zone is off in a Word 2013 layout', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-hyphenation-')))
  const plainPath = join(dir, 'plain.docx')
  const modernPath = join(dir, 'modern.docx')
  writeFileSync(plainPath, await fixture('<w:hyphenationZone w:val="425"/>'))
  writeFileSync(
    modernPath,
    await fixture(
      '<w:autoHyphenation/><w:hyphenationZone w:val="5000"/><w:compat><w:compatSetting w:name="compatibilityMode" ' +
        'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>',
    ),
  )
  let session = await open(plainPath)
  try {
    await expect(session.page.locator('.doc-page p', { hasText: 'lower' })).toBeVisible()
    expect(await session.page.locator('.doc-page .doc-shy, .doc-page .doc-hyph').count()).toBe(0)
    expect((await hyphenLines(session.page, 'lower'))!.hyphens).toBe(0)

    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    session = await open(modernPath)
    const { app, page } = session
    // the file's zone is wider than the column, and the page hyphenates all the same
    await expect.poll(async () => (await hyphenLines(page, 'lower'))?.hyphens).toBeGreaterThan(3)
    await clickMenu(app, ['Tools', 'Hyphenation…'])
    const dialog = page.locator('.hyphenation-dialog')
    await expect(dialog.getByLabel('Hyphenation zone')).toBeDisabled()
    await expect(dialog.getByLabel('Hyphenate words in CAPS')).toBeEnabled()
    await expect(dialog.getByLabel('Limit consecutive hyphens to')).toBeEnabled()
    await page.keyboard.press('Escape')
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a document that declares no language is not hyphenated, as in Word; the same text under en-US is', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-hyphenation-')))
  const bare = join(dir, 'no-language.docx')
  const declared = join(dir, 'en-US.docx')
  writeFileSync(bare, await fixture('<w:autoHyphenation/>', undefined, null))
  writeFileSync(declared, await fixture('<w:autoHyphenation/>'))
  let session = await open(bare)
  try {
    await expect(session.page.locator('.doc-page p', { hasText: 'lower' })).toBeVisible()
    expect(await session.page.locator('.doc-page .doc-shy').count()).toBe(0)
    expect((await hyphenLines(session.page, 'lower'))!.hyphens).toBe(0)
    // print: the same lines, no hyphen at a line end
    const pdf = join(dir, 'no-language.pdf')
    await session.app.evaluate(({ dialog: d, shell }, out) => {
      d.showSaveDialog = async () => ({ canceled: false, filePath: out })
      shell.showItemInFolder = () => {}
    }, pdf)
    expect(
      await session.page.evaluate(() =>
        (window as unknown as { __exportPdf: () => Promise<boolean> }).__exportPdf(),
      ),
    ).toBe(true)
    expect(await pdfLineEndHyphens(pdf)).toBe(0)

    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    session = await open(declared)
    await expect
      .poll(async () => (await hyphenLines(session.page, 'lower'))?.hyphens)
      .toBeGreaterThan(3)
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('text that names its own language hyphenates under defaults that name none: a run, its character style, its paragraph style', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-hyphenation-')))
  const docPath = join(dir, 'run-language.docx')
  const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
  const styles =
    `<?xml version="1.0" encoding="UTF-8"?><w:styles ${W_NS}><w:docDefaults><w:rPrDefault><w:rPr>` +
    '<w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="EnPara"><w:name w:val="EnPara"/><w:basedOn w:val="Normal"/><w:rPr><w:lang w:val="en-US"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="EnChild"><w:name w:val="EnChild"/><w:basedOn w:val="EnPara"/></w:style>' +
    '<w:style w:type="character" w:styleId="EnChar"><w:name w:val="EnChar"/><w:rPr><w:lang w:val="en-US"/></w:rPr></w:style></w:styles>'
  const p = (tag: string, pStyle = '', rPr = '') =>
    `<w:p>${pStyle ? `<w:pPr><w:pStyle w:val="${pStyle}"/></w:pPr>` : ''}<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}` +
    `<w:t>${LOWER.replace('lower', tag)}</w:t></w:r></w:p>`
  const body =
    p('direct', '', '<w:lang w:val="en-US"/>') +
    p('plain') +
    p('character', '', '<w:rStyle w:val="EnChar"/>') +
    p('paragraph', 'EnPara') +
    p('inherited', 'EnChild') +
    p('other', 'EnPara', '<w:lang w:val="zh-CN"/>')
  writeFileSync(
    docPath,
    await withWordPart(
      await withWordPart(
        await docxWith(body + SECT_PR),
        'settings',
        settingsXml('<w:autoHyphenation/>'),
      ),
      'styles',
      styles,
    ),
  )
  const session = await open(docPath)
  try {
    const { app, page } = session
    for (const tag of ['direct', 'character', 'paragraph', 'inherited'])
      await expect
        .poll(async () => (await hyphenLines(page, tag))?.hyphens, { message: tag })
        .toBeGreaterThan(3)
    // Word leaves these whole: no language, and a run whose own language has no hyphenation
    expect((await hyphenLines(page, 'plain'))!.hyphens).toBe(0)
    expect((await hyphenLines(page, 'other'))!.hyphens).toBe(0)

    // the dialog offers its settings: some of the text can be hyphenated
    await clickMenu(app, ['Tools', 'Hyphenation…'])
    const dialog = page.locator('.hyphenation-dialog')
    await expect(dialog.getByLabel('Automatically hyphenate document')).toBeEnabled()
    await expect(dialog.getByLabel('Automatically hyphenate document')).toBeChecked()
    await page.keyboard.press('Escape')
    await expect(dialog).toBeHidden()

    // an edit and a save keep the languages where they were
    await page.locator('.doc-page p', { hasText: 'plain' }).click()
    await page.keyboard.type('zq')
    await expect(page.locator('.doc-page p', { hasText: 'plain' })).toContainText('zq')
    await saveFromMenu(app)
    await expect
      .poll(async () => await part(docPath, 'word/document.xml'), { timeout: 30_000 })
      .toContain('zq')
    const saved = await part(docPath, 'word/document.xml')
    expect(saved.match(/<w:lang w:val="en-US"\/>/g)).toHaveLength(1)
    expect(saved).toContain('<w:rStyle w:val="EnChar"/>')
    expect(await part(docPath, 'word/styles.xml')).not.toMatch(
      /<w:rPrDefault>(?:(?!<\/w:rPrDefault>).)*<w:lang/s,
    )

    const pdf = join(dir, 'run-language.pdf')
    await app.evaluate(({ dialog: d, shell }, out) => {
      d.showSaveDialog = async () => ({ canceled: false, filePath: out })
      shell.showItemInFolder = () => {}
    }, pdf)
    expect(
      await page.evaluate(() =>
        (window as unknown as { __exportPdf: () => Promise<boolean> }).__exportPdf(),
      ),
    ).toBe(true)
    expect(await pdfLineEndHyphens(pdf)).toBeGreaterThan(12)
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Manual… asks about each word that could break at a line end and puts an optional hyphen where told', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-hyphenation-')))
  const docPath = join(dir, 'manual.docx')
  writeFileSync(docPath, await fixture('', para(LOWER)))
  const session = await open(docPath)
  try {
    const { app, page } = session
    const text = () =>
      page.evaluate(
        () =>
          (
            window as unknown as {
              __aidocs: { editor: { state: { doc: { textContent: string } } } }
            }
          ).__aidocs.editor.state.doc.textContent,
      )
    const lines = async () => (await hyphenLines(page, 'lower'))!.lines
    const linesBefore = await lines()

    await clickMenu(app, ['Tools', 'Hyphenation…'])
    await page.locator('.hyphenation-dialog').getByRole('button', { name: 'Manual…' }).click()
    const manual = page.locator('.manual-hyphenation-dialog')
    const word = manual.locator('.manual-hyphenation-word')
    /** the word asked about, once the dialog has found it */
    const shown = async () => {
      await expect(word).toHaveAttribute('aria-busy', 'false')
      return ((await word.textContent()) ?? '').replace(/-/g, '')
    }
    await expect(word.locator('.manual-hyphenation-point.active')).toHaveCount(1, {
      timeout: 15_000,
    })
    await page.screenshot({ path: test.info().outputPath('manual.png') })

    // Yes: an optional hyphen at the offered place
    const first = await shown()
    expect(LOWER).toContain(first)
    const head = await word.evaluate((el) => {
      let out = ''
      for (const piece of Array.from(el.children)) {
        const point = piece.querySelector('button')
        if (point?.classList.contains('active')) break
        out += (piece.textContent ?? '').replace(/^-/, '')
      }
      return out
    })
    // pressed twice in a row it is still one answer: the word is gone from the dialog with the first
    await manual.getByRole('button', { name: 'Yes' }).evaluate(async (yes: HTMLButtonElement) => {
      yes.click()
      await Promise.resolve()
      yes.click()
    })
    await expect.poll(text).toContain(`${head}\u00ad${first.slice(head.length)}`)
    expect(await text()).not.toContain('\u00ad\u00ad')

    // No: the word stays whole and the run moves on
    const second = await shown()
    expect(second).not.toBe(first)
    await manual.getByRole('button', { name: 'No' }).click()

    // another place in the word, then stop
    const third = await shown()
    expect(third).not.toBe(second)
    const points = word.locator('.manual-hyphenation-point')
    await points.first().click()
    await expect(points.first()).toHaveClass(/active/)
    const thirdHead = ((await word.locator('span').first().textContent()) ?? '').replace(/-/g, '')
    await manual.getByRole('button', { name: 'Yes' }).click()
    await expect.poll(text).toContain(`${thirdHead}\u00ad${third.slice(thirdHead.length)}`)
    await manual.getByRole('button', { name: 'Cancel' }).click()
    await expect(manual).toBeHidden()

    // only the two answers stay: no automatic hyphenation, no candidates left on the page
    await expect(page.locator('.doc-page .doc-shy')).toHaveCount(0)
    expect((await text()).match(/\u00ad/g)).toHaveLength(2)
    expect(await text()).toContain(second)
    expect(await lines()).toBeLessThanOrEqual(linesBefore)

    await saveFromMenu(app)
    await expect
      .poll(async () => {
        const xml = await part(docPath, 'word/document.xml')
        return [
          (xml.match(/<w:softHyphen\/>/g) ?? []).length,
          /autoHyphenation/.test(await part(docPath, 'word/settings.xml')),
        ]
      })
      .toEqual([2, false])
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
