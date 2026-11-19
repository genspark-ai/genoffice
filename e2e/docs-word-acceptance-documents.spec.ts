/** Word acceptance tasks, documents and saving: each walked along Word's path and read back from the saved file. */
import { test, expect } from '@playwright/test'
import { existsSync, mkdtempSync, readdirSync, realpathSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'
import { para, saveFromMenu } from './docs-objects-fixture'
import {
  caretAtEnd,
  clickMenu,
  open,
  part,
  pickSavePath,
  savedBody,
} from './docs-word-acceptance-kit'

const VIDEO_DIR = 'docs-word-acceptance-documents'
const PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

test('File > New Document starts a blank document whose typed paragraph is saved', async () => {
  test.setTimeout(180_000)
  const s = await open(VIDEO_DIR, para('existing'), 'existing.docx')
  try {
    const { app, page, dir } = s
    const existing = page.locator('.doc-page p', { hasText: 'existing' })
    await expect(existing).toBeVisible({ timeout: 30_000 })

    await clickMenu(app, ['File', 'New Document'])
    await expect(existing).toHaveCount(0, { timeout: 30_000 })
    await page.locator('.doc-page p').first().click()
    await page.keyboard.type('A first paragraph.')
    await expect(page.locator('.doc-page p', { hasText: 'A first paragraph.' })).toBeVisible()

    // the first save of a new document does not ask where: it lands in the default folder
    await saveFromMenu(app)
    const savedFile = () => readdirSync(dir).find((name) => name !== 'existing.docx')
    await expect.poll(savedFile, { timeout: 30_000 }).toMatch(/\.docx$/)
    await expect
      .poll(() => savedBody(join(dir, savedFile()!)), { timeout: 30_000 })
      .toMatch(/<w:t[^>]*>A first paragraph\.<\/w:t>/)
    expect(await savedBody(s.docPath)).not.toContain('A first paragraph.')
  } finally {
    await s.close()
  }
})

test('File > Save As writes a .docx that opens again with its text', async () => {
  test.setTimeout(180_000)
  const out = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-acceptance-copy-')))
  const copy = join(out, 'copy.docx')
  try {
    const s = await open(VIDEO_DIR, para('kept text'), 'source.docx')
    try {
      const { app, page } = s
      await expect(page.locator('.doc-page p', { hasText: 'kept text' })).toBeVisible({
        timeout: 30_000,
      })
      await caretAtEnd(page, 'kept text')
      await page.keyboard.type(' and more')
      await pickSavePath(app, copy)
      await clickMenu(app, ['File', 'Save As…'])
      await expect
        .poll(() => (existsSync(copy) ? savedBody(copy) : ''), { timeout: 30_000 })
        .toContain('kept text and more')
      expect(await savedBody(s.docPath)).not.toContain('and more')
    } finally {
      await s.close()
    }

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: VIDEO_DIR,
      openFile: copy,
    })
    try {
      const page = await waitForPageWithUrl(launched.app, '://docs/')
      await expect(page.locator('.doc-page p', { hasText: 'kept text and more' })).toBeVisible({
        timeout: 30_000,
      })
    } finally {
      await closeAndSaveVideo(launched, VIDEO_DIR)
    }
  } finally {
    rmSync(out, { recursive: true, force: true })
  }
})

test('File > Print shows as many pages as the status bar counts; File > Properties saves title and author', async () => {
  test.setTimeout(180_000)
  const s = await open(
    VIDEO_DIR,
    para('first page') + PAGE_BREAK + para('second page') + PAGE_BREAK + para('third page'),
    'pages.docx',
  )
  try {
    const { app, page, docPath } = s
    await expect(page.locator('.doc-page p', { hasText: 'third page' })).toBeAttached({
      timeout: 30_000,
    })
    const statusPage = page.locator('.status-page')
    await expect(statusPage).toHaveText(/^Page \d+ of 3$/)

    // the shortcut is the menu's accelerator, which a synthetic key press does not reach
    const accelerator = await app.evaluate(
      ({ Menu }) =>
        Menu.getApplicationMenu()
          ?.items.find((item) => item.label === 'File')
          ?.submenu?.items.find((item) => item.label === 'Print…')?.accelerator,
    )
    expect(accelerator).toBe('CmdOrCtrl+P')
    await clickMenu(app, ['File', 'Print…'])
    const print = page.locator('.print-dialog')
    await expect(print.locator('.print-preview-nav')).toContainText('1 / 3', { timeout: 30_000 })
    await expect(print.locator('.print-page-count')).toContainText('3')
    await page.screenshot({ path: test.info().outputPath('print-preview.png') })
    await print.getByRole('button', { name: 'Cancel', exact: true }).click()
    await expect(print).toBeHidden()

    await clickMenu(app, ['File', 'Properties…'])
    const properties = page.locator('.properties-dialog')
    await properties.locator('input[name="core.title"]').fill('Annual report')
    await properties.locator('input[name="core.creator"]').fill('Ada Lovelace')
    await properties.locator('button.primary').click()
    await expect(properties).toBeHidden()

    await saveFromMenu(app)
    await expect
      .poll(() => part(docPath, 'docProps/core.xml'), { timeout: 30_000 })
      .toContain('<dc:title>Annual report</dc:title>')
    expect(await part(docPath, 'docProps/core.xml')).toContain(
      '<dc:creator>Ada Lovelace</dc:creator>',
    )
  } finally {
    await s.close()
  }
})
