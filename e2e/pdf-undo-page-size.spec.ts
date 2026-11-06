import { test, expect } from '@playwright/test'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PDFDocument } from 'pdf-lib'
import { launchShell, waitForPageWithUrl, closeAndSaveVideo } from './helpers'

test('Page size, written to the file at once, is undone and redone in the file', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-pdf-page-size-'))
  const pdfPath = join(dir, 'two.pdf')
  const doc = await PDFDocument.create()
  doc.addPage([400, 600])
  doc.addPage([400, 600])
  await writeFile(pdfPath, await doc.save({ useObjectStreams: false }))
  const sizes = async () =>
    (await PDFDocument.load(await readFile(pdfPath)))
      .getPages()
      .map((p) => [p.getWidth(), p.getHeight()])

  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'pdf-undo-page-size',
    openFile: pdfPath,
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://pdf/')
    await expect(page.locator('.pdf-thumb')).toHaveCount(2)

    await page.getByText('Pages', { exact: true }).first().click()
    await page.getByRole('button', { name: 'Page size', exact: true }).click()
    await page.getByRole('button', { name: 'A4', exact: true }).click()
    await expect.poll(sizes).toEqual([
      [595, 842],
      [595, 842],
    ])

    await page.keyboard.press('ControlOrMeta+z')
    await expect.poll(sizes).toEqual([
      [400, 600],
      [400, 600],
    ])
    await page.keyboard.press('ControlOrMeta+Shift+z')
    await expect.poll(sizes).toEqual([
      [595, 842],
      [595, 842],
    ])
  } finally {
    await closeAndSaveVideo(launched, 'pdf-undo-page-size')
  }
})
