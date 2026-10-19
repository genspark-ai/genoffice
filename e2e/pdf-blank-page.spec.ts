import { test, expect } from '@playwright/test'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PDFDocument } from 'pdf-lib'
import { launchShell, waitForPageWithUrl, closeAndSaveVideo } from './helpers'

test('Insert blank page is a pending, undoable edit that lands in the file on save', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-pdf-blank-'))
  const pdfPath = join(dir, 'two.pdf')
  const doc = await PDFDocument.create()
  doc.addPage([400, 600])
  doc.addPage([612, 792])
  await writeFile(pdfPath, await doc.save({ useObjectStreams: false }))

  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'pdf-blank-page',
    openFile: pdfPath,
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://pdf/')
    const thumbs = page.locator('.pdf-thumb')
    await expect(thumbs).toHaveCount(2)

    await thumbs.first().locator('.pdf-thumb-box').click({ button: 'right' })
    await page.getByRole('button', { name: 'Insert blank page', exact: true }).click()

    // Appears at once as page 2, nothing written yet
    await expect(thumbs).toHaveCount(3)
    await expect(page.locator('.pdf-page')).toHaveCount(3)
    await expect(page.locator('.tb-save-pending')).toBeVisible()
    const before = await readFile(pdfPath)
    expect((await PDFDocument.load(before)).getPageCount()).toBe(2)

    await page.keyboard.press('ControlOrMeta+z')
    await expect(thumbs).toHaveCount(2)
    await page.keyboard.press('ControlOrMeta+Shift+z')
    await expect(thumbs).toHaveCount(3)

    await page.keyboard.press('ControlOrMeta+s')
    await expect(page.locator('.tb-save-pending')).toHaveCount(0)
    await expect(thumbs).toHaveCount(3)

    const saved = await PDFDocument.load(await readFile(pdfPath))
    expect(saved.getPageCount()).toBe(3)
    const blank = saved.getPage(1)
    expect([blank.getWidth(), blank.getHeight()]).toEqual([400, 600])
    expect(saved.getPage(2).getWidth()).toBe(612)
  } finally {
    await closeAndSaveVideo(launched, 'pdf-blank-page')
  }
})
