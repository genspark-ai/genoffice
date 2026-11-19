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

    // The step is kept once the editor has reloaded the written file, a moment after the write:
    // a key pressed before then finds nothing to undo, so it is pressed until the file changes
    const stepUntil = async (keys: string, expected: number[][]) =>
      expect(async () => {
        await page.keyboard.press(keys)
        await expect.poll(sizes, { timeout: 1500 }).toEqual(expected)
      }).toPass({ timeout: 15_000 })
    await stepUntil('ControlOrMeta+z', [
      [400, 600],
      [400, 600],
    ])
    await stepUntil('ControlOrMeta+Shift+z', [
      [595, 842],
      [595, 842],
    ])
  } finally {
    await closeAndSaveVideo(launched, 'pdf-undo-page-size')
  }
})
