/**
 * Format > Paragraph… > Asian Typography > Options… > Compress punctuation and
 * Japanese kana: the radio writes w:characterSpacingControl, and a justified
 * line takes up a mark that may not start the next one by tightening its kana.
 */
import { test, expect, type ElectronApplication, type Locator, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu, settingsXml, withWordPart } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-compress-kana'
const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '</w:sectPr>'
const WORD_2013 =
  '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
const KANJI = '\u56fd'
const KANA = '\u3042'
const MARK = '\u3005'
const SWEEP = 36

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

const settingsOf = async (docPath: string) =>
  (await (
    await JSZip.loadAsync(readFileSync(docPath))
  )
    .file('word/settings.xml')
    ?.async('string')) ?? ''

async function open(docPath: string) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await expect(page.locator('.doc-page p', { hasText: 'head' }).first()).toBeVisible({
    timeout: 30_000,
  })
  return { launched, app: launched.app, page }
}

/** per paragraph: the length of its first line, and the tightest letter-spacing on it in em */
const firstLines = (page: Page) =>
  page.evaluate(() =>
    [...document.querySelectorAll<HTMLElement>('.doc-page p')].map((para) => {
      const walker = document.createTreeWalker(para, NodeFilter.SHOW_TEXT)
      const range = document.createRange()
      const em = parseFloat(getComputedStyle(para).fontSize)
      let top: number | null = null
      let length = 0
      let tightest = 0
      while (walker.nextNode()) {
        const node = walker.currentNode as Text
        for (let i = 0; i < node.data.length; i++) {
          if (node.data[i] === '\u2060') continue
          range.setStart(node, i)
          range.setEnd(node, i + 1)
          const rect = range.getBoundingClientRect()
          top ??= rect.top
          if (rect.top > top + 1) return { length, tightest }
          length++
          const spacing = parseFloat(getComputedStyle(node.parentElement!).letterSpacing) || 0
          tightest = Math.min(tightest, spacing / em)
        }
      }
      return { length, tightest }
    }),
  )

async function openTypography(app: ElectronApplication, page: Page) {
  await page.locator('.doc-page p', { hasText: 'head' }).first().click()
  await clickMenu(app, ['Format', 'Paragraph…'])
  const dialog = page.locator('.modal', { hasText: 'Paragraph' }).first()
  await dialog.getByRole('tab', { name: 'Asian Typography' }).click()
  await dialog.getByRole('button', { name: 'Options…' }).click()
  const typography = page.getByRole('dialog', { name: 'Typography' })
  await expect(typography).toBeVisible()
  return { dialog, typography }
}

async function confirm(dialog: Locator, typography: Locator) {
  await typography.getByRole('button', { name: 'OK', exact: true }).click()
  await expect(typography).toBeHidden()
  await dialog.getByRole('button', { name: 'OK', exact: true }).click()
  await expect(dialog).toBeHidden()
}

test('Compress punctuation and Japanese kana: the option is written, kept and drawn', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-compress-kana-')))
  const docPath = join(dir, 'kana.docx')
  const ja = (text: string, jc: string) =>
    `<w:p><w:pPr><w:jc w:val="${jc}"/></w:pPr><w:r><w:rPr><w:lang w:eastAsia="ja-JP"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`
  // one of these lengths leaves the mark where the second line would start
  const sweep = (ch: string, jc: string) =>
    Array.from({ length: SWEEP }, (_, i) => ja(ch.repeat(22 + i) + MARK + ch.repeat(3), jc))
  const body =
    para('head') +
    sweep(KANA, 'both').join('') +
    sweep(KANJI, 'both').join('') +
    sweep(KANA, 'left').join('') +
    SECT_PR
  writeFileSync(
    docPath,
    await withWordPart(await docxWith(body), 'settings', settingsXml(WORD_2013)),
  )
  const groups = async (page: Page) => {
    const lines = (await firstLines(page)).slice(1)
    return {
      kana: lines.slice(0, SWEEP),
      kanji: lines.slice(SWEEP, 2 * SWEEP),
      left: lines.slice(2 * SWEEP),
    }
  }
  const lengths = (lines: Array<{ length: number }>) => lines.map((line) => line.length)

  let run = await open(docPath)
  let before: Awaited<ReturnType<typeof groups>>
  let drawn: number[]
  try {
    const { app, page } = run
    before = await groups(page)
    // somewhere in the sweep the mark does not fit and takes the kana before it down
    expect(new Set(lengths(before.kana)).size).toBeGreaterThan(1)
    expect(before.kana.every((line) => line.tightest === 0)).toBe(true)

    const { dialog, typography } = await openTypography(app, page)
    await expect(typography.getByLabel('Do not compress')).toBeChecked()
    const kana = typography.getByLabel('Compress punctuation and Japanese kana')
    await expect(kana).toBeEnabled()
    await kana.check()
    await page.screenshot({ path: test.info().outputPath('typography-kana.png') })
    await confirm(dialog, typography)

    // kana lines take the mark up; kanji have nothing to give, unjustified lines are left alone
    await expect
      .poll(async () => {
        const now = lengths((await groups(page)).kana)
        return now.filter((length, i) => length > before.kana[i].length).length
      })
      .toBeGreaterThan(0)
    const after = await groups(page)
    drawn = lengths(after.kana)
    after.kana.forEach((line, i) => {
      expect(line.length, `kana ${i}`).toBeGreaterThanOrEqual(before.kana[i].length)
      // a sixteenth of the advance at most, plus the overshoot that makes the line fit
      expect(line.tightest, `kana ${i}`).toBeGreaterThan(-1 / 16 - 0.01)
      if (line.length > before.kana[i].length) expect(line.tightest, `kana ${i}`).toBeLessThan(0)
    })
    expect(lengths(after.kanji)).toEqual(lengths(before.kanji))
    expect(after.kanji.every((line) => line.tightest === 0)).toBe(true)
    expect(lengths(after.left)).toEqual(lengths(before.left))
    await page.screenshot({ path: test.info().outputPath('page-kana.png') })

    await saveFromMenu(app)
    await expect
      .poll(async () => (await settingsOf(docPath)).match(/<w:characterSpacingControl[^>]*>/g), {
        timeout: 30_000,
      })
      .toEqual(['<w:characterSpacingControl w:val="compressPunctuationAndJapaneseKana"/>'])
    expect(await settingsOf(docPath)).toContain('w:name="compatibilityMode"')
  } finally {
    await closeAndSaveVideo(run.launched, VIDEO_DIR)
  }

  // reopened: the same lines, the dialog shows the option; punctuation only puts the lines back
  run = await open(docPath)
  try {
    const { app, page } = run
    await expect.poll(async () => lengths((await groups(page)).kana)).toEqual(drawn)
    const { dialog, typography } = await openTypography(app, page)
    await expect(typography.getByLabel('Compress punctuation and Japanese kana')).toBeChecked()
    await typography.getByLabel('Compress punctuation only').check()
    await confirm(dialog, typography)
    await expect.poll(async () => lengths((await groups(page)).kana)).toEqual(lengths(before.kana))
    expect((await groups(page)).kana.every((line) => line.tightest === 0)).toBe(true)
    await saveFromMenu(app)
    await expect
      .poll(async () => (await settingsOf(docPath)).match(/<w:characterSpacingControl[^>]*>/g), {
        timeout: 30_000,
      })
      .toEqual(['<w:characterSpacingControl w:val="compressPunctuation"/>'])
  } finally {
    await closeAndSaveVideo(run.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
