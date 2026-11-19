/**
 * A header or footer command chosen in Draft or Web Layout opens the strip in
 * Print Layout, and Merge All Windows brings the windows Arrange All made back
 * as tabs, unsaved edits included.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { saveFromMenu } from './docs-objects-fixture'
import { clickMenu } from './docs-word-acceptance-kit'
import { DRAFT_FOOTER, DRAFT_HEADER, draftDocx } from './docs-draft-fixture'

const VIDEO_DIR = 'docs-draft-hf-and-gather-windows'

const part = async (docPath: string, name: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file(name)?.async('string') ?? ''

const ribbonTab = (page: Page, name: string) =>
  page
    .locator('.ribbon-tab', { hasText: new RegExp(`^${name}$`) })
    .first()
    .click()

const ribbonButton = (page: Page, name: string) =>
  page.locator('.ribbon :is(.rb-big, .rb-small)', { hasText: new RegExp(`^${name}$`) })

const visibleWindows = (app: ElectronApplication) =>
  app.evaluate(({ BrowserWindow }) =>
    BrowserWindow.getAllWindows()
      .filter((w) => w.isVisible() && !w.isMinimized())
      .map((w) => w.getBounds()),
  )

test('a header or footer command in Draft and Web Layout opens the strip in Print Layout', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-draft-hf-')))
  const docPath = join(dir, 'report.docx')
  writeFileSync(docPath, await draftDocx())
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const canvas = page.locator('.doc-zoom')
    const strip = page.locator('.page-hf-editor')
    await expect(page.locator('.doc-page p', { hasText: 'Opening paragraph' })).toBeVisible({
      timeout: 30_000,
    })
    // the open-time respell kick focuses the body and shields keystrokes for a moment
    await page.waitForTimeout(1500)

    // View > Header and Footer, from Draft
    await clickMenu(app, ['View', 'Draft'])
    await expect(canvas).toHaveClass(/\bview-draft\b/)
    await clickMenu(app, ['View', 'Header and Footer'])
    await expect(canvas).toHaveClass(/\bview-print\b/)
    await expect(page.locator('.page-hf-header .page-hf-editor')).toBeFocused()
    await page.keyboard.press('End')
    await page.keyboard.type(' from draft')
    await page.screenshot({ path: test.info().outputPath('draft-header-open.png') })
    await page.keyboard.press('Escape')
    await expect(strip).toHaveCount(0)
    await expect(page.locator('.page-wrap > .page-hf-header')).toContainText(
      `${DRAFT_HEADER} from draft`,
    )
    // the view stays where the command took it
    await expect(canvas).toHaveClass(/\bview-print\b/)

    // Insert > Footer > Edit Footer on the ribbon, from Web Layout
    await clickMenu(app, ['View', 'Web Layout'])
    await expect(canvas).toHaveClass(/\bview-web\b/)
    await ribbonTab(page, 'Insert')
    const footer = ribbonButton(page, 'Footer')
    await footer.scrollIntoViewIfNeeded()
    await footer.click()
    await page.locator('[data-rb-panel] button', { hasText: /^Edit Footer$/ }).click()
    await expect(canvas).toHaveClass(/\bview-print\b/)
    await expect(page.locator('.page-hf-footer .page-hf-editor')).toBeFocused()
    await page.keyboard.press('End')
    await page.keyboard.type(' from web')
    await page.keyboard.press('Escape')
    await expect(strip).toHaveCount(0)

    await saveFromMenu(app)
    await expect
      .poll(() => part(docPath, 'word/header1.xml'), { timeout: 30_000 })
      .toContain(`${DRAFT_HEADER} from draft`)
    expect(await part(docPath, 'word/footer1.xml')).toContain(`${DRAFT_FOOTER} from web`)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR).catch(() => {})
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Merge All Windows brings the arranged windows back as tabs with their unsaved edits', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-gather-')))
  const docPath = join(dir, 'report.docx')
  writeFileSync(docPath, await draftDocx())
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    // with a file on argv the first window may be the editor view, not the shell's
    const shell = await waitForPageWithUrl(app, 'shell/out')
    const tabs = shell.locator('.tab-bar .tab-item:not(.tab-home)')
    const first = await waitForPageWithUrl(app, '://docs/')
    const opening = first.locator('.doc-page p', { hasText: 'Opening paragraph' })
    await expect(opening).toBeVisible({ timeout: 30_000 })
    await first.waitForTimeout(1500)
    await opening.click()
    await first.keyboard.press('End')
    await first.keyboard.type(' unsaved words')

    await expect
      .poll(
        async () => {
          for (const candidate of app.windows()) {
            const opened = await candidate
              .evaluate(async () => {
                const api = (window as never as { aiOffice?: { newDoc(): Promise<void> } }).aiOffice
                if (!api) return false
                await api.newDoc()
                return true
              })
              .catch(() => false)
            if (opened) return true
          }
          return false
        },
        { timeout: 30_000 },
      )
      .toBe(true)
    const docsPages = () => app.windows().filter((p) => p.url().includes('://docs/'))
    await expect.poll(() => docsPages().length, { timeout: 30_000 }).toBe(2)
    const second = docsPages().find((p) => p !== first)!
    await expect(second.locator('.status-page')).toContainText(/\d/, { timeout: 30_000 })
    await expect(tabs).toHaveCount(2)
    const before = (await visibleWindows(app))[0]

    await ribbonTab(second, 'View')
    // every document is a tab: there is no window to bring back
    await expect(ribbonButton(second, 'Merge All Windows')).toBeDisabled()
    const arrange = ribbonButton(second, 'Arrange All')
    await arrange.scrollIntoViewIfNeeded()
    await arrange.click()
    await expect.poll(async () => (await visibleWindows(app)).length, { timeout: 30_000 }).toBe(2)
    await expect(tabs).toHaveCount(1)

    // asked from the window that moved out
    await ribbonTab(first, 'View')
    const merge = ribbonButton(first, 'Merge All Windows')
    await merge.scrollIntoViewIfNeeded()
    await expect(merge).toBeEnabled()
    await merge.click()

    await expect.poll(async () => (await visibleWindows(app)).length, { timeout: 30_000 }).toBe(1)
    await expect(tabs).toHaveCount(2)
    // the shell window is back where it was before it became a tile
    await expect.poll(async () => (await visibleWindows(app))[0]).toEqual(before)
    // the document the command came from is the tab in front
    await expect(tabs.filter({ hasText: 'report.docx' })).toHaveClass(/\bactive\b/)

    // the same live document: its unsaved words are there, it takes more, and it saves
    const edited = first.locator('.doc-page p', { hasText: 'unsaved words' })
    await expect(edited).toBeVisible()
    await edited.click()
    await first.keyboard.press('End')
    await first.keyboard.type(' and more')
    await expect(first.locator('.doc-page p', { hasText: 'unsaved words and more' })).toBeVisible()
    await shell.screenshot({ path: test.info().outputPath('merged-tabs.png') })
    await saveFromMenu(app)
    await expect
      .poll(() => part(docPath, 'word/document.xml'), { timeout: 30_000 })
      .toContain('unsaved words and more')

    // with nothing detached the command is greyed again, in both documents
    await expect(ribbonButton(first, 'Merge All Windows')).toBeDisabled()
    await shell.locator('.tab-bar .tab-item:not(.tab-home)', { hasNotText: 'report.docx' }).click()
    await expect(ribbonButton(second, 'Merge All Windows')).toBeDisabled()
    await expect(tabs).toHaveCount(2)
    expect(await visibleWindows(app)).toHaveLength(1)
  } finally {
    await launched.app
      .evaluate(({ BrowserWindow }) => {
        for (const w of BrowserWindow.getAllWindows()) w.destroy()
      })
      .catch(() => {})
    await closeAndSaveVideo(launched, VIDEO_DIR).catch(() => {})
    rmSync(dir, { recursive: true, force: true })
  }
})
