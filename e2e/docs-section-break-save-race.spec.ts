/**
 * A section break brings a save of its own. A Save pressed before that save
 * has finished queues behind it, and used to be built from the document as it
 * was before the first save reloaded it: the file lost the break and held one
 * paragraph twice. Both ways to a break are driven here with no wait before
 * the Save.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-section-break-save-race'
const RUNS = 3
const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/></w:sectPr>'
const LINES = Array.from({ length: 30 }, (_, i) => `line ${String(i + 1).padStart(2, '0')}`)

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

/** what the saved file holds: its section properties, and the lines in document order */
async function savedFile(docPath: string) {
  const zip = await JSZip.loadAsync(readFileSync(docPath))
  const xml = (await zip.file('word/document.xml')?.async('string')) ?? ''
  return {
    sections: (xml.match(/<w:sectPr\b/g) ?? []).length,
    lines: [...xml.matchAll(/<w:t(?: [^>]*)?>(line \d\d)<\/w:t>/g)].map((m) => m[1]),
  }
}

const linesOnPage = (page: Page) =>
  page
    .locator('.doc-page p')
    .evaluateAll((paras) =>
      paras.map((p) => p.textContent ?? '').filter((text) => text.startsWith('line ')),
    )

async function openLines() {
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-break-save-race-')))
  const docPath = join(dir, 'lines.docx')
  writeFileSync(docPath, await docxWith(LINES.map(para).join('') + SECT_PR))
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  const fifth = page.locator('.doc-page p', { hasText: 'line 05' }).first()
  await expect(fifth).toBeVisible({ timeout: 30_000 })
  await fifth.click()
  await page.keyboard.press('End')
  return { dir, docPath, launched, page }
}

/** the file and the page once the saves have settled: the break is there and every line once */
async function expectIntact(docPath: string, page: Page) {
  await expect
    .poll(() => savedFile(docPath), { timeout: 30_000 })
    .toEqual({ sections: 2, lines: LINES })
  // a wrong second save lands within a moment of the first: look again after it would have
  await page.waitForTimeout(1500)
  expect(await savedFile(docPath)).toEqual({ sections: 2, lines: LINES })
  expect(await linesOnPage(page)).toEqual(LINES)
}

for (let run = 1; run <= RUNS; run++) {
  test(`Columns for This point forward, then Save at once, keeps the break and every paragraph (run ${run})`, async () => {
    test.setTimeout(120_000)
    const { dir, docPath, launched, page } = await openLines()
    try {
      await clickMenu(launched.app, ['Format', 'Columns…'])
      const dialog = page.locator('.columns-dialog')
      await dialog.locator('input[name="columns"]').nth(1).check()
      await dialog.locator('select[name="applyTo"]').selectOption('forward')
      await dialog.locator('button.primary').click()
      await saveFromMenu(launched.app)
      await expectIntact(docPath, page)
    } finally {
      await closeAndSaveVideo(launched, VIDEO_DIR)
      rmSync(dir, { recursive: true, force: true })
    }
  })

  test(`a section break from the Layout tab, then Cmd+S at once, keeps the break and every paragraph (run ${run})`, async () => {
    test.setTimeout(120_000)
    const { dir, docPath, launched, page } = await openLines()
    try {
      await page.locator('.ribbon-tab', { hasText: /^Layout$/ }).click()
      await page.getByRole('button', { name: 'Breaks', exact: true }).click()
      await page
        .locator('[data-rb-panel] button')
        .filter({ has: page.locator('b', { hasText: /^Continuous$/ }) })
        .click()
      await page.keyboard.press('ControlOrMeta+s')
      await expectIntact(docPath, page)
    } finally {
      await closeAndSaveVideo(launched, VIDEO_DIR)
      rmSync(dir, { recursive: true, force: true })
    }
  })
}
