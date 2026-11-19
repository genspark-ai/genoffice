import { test, expect } from '@playwright/test'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PDFDocument, StandardFonts } from 'pdf-lib'
import { launchShell, waitForPageWithUrl, closeAndSaveVideo } from './helpers'

test('redaction stays in the current tab and accumulates on its working copy', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-redaction-'))
  const source = join(dir, 'original.pdf')
  const copy = join(dir, 'redacted.pdf')
  const pdf = await PDFDocument.create()
  const page = pdf.addPage([300, 200])
  const font = await pdf.embedFont(StandardFonts.Helvetica)
  page.drawText('SECRET-ONE', { x: 20, y: 130, size: 18, font })
  page.drawText('SECRET-TWO', { x: 20, y: 60, size: 18, font })
  page.drawText('PUBLIC', { x: 20, y: 20, size: 18, font })
  const original = await pdf.save()
  await writeFile(source, original)
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'pdf-redaction',
    openFile: source,
  })
  try {
    const editor = await waitForPageWithUrl(launched.app, '://pdf/')
    await expect(editor.locator('.pdf-page').first()).toBeVisible()
    editor.on('dialog', (dialog) => dialog.accept())
    await editor.getByRole('button', { name: 'Annotate', exact: true }).click()
    await editor.getByRole('button', { name: 'Redact area', exact: true }).click()
    const mark = async (top: number, bottom: number) => {
      const box = await editor.locator('.pdf-redaction-layer').first().boundingBox()
      if (!box) throw new Error('Missing redaction layer')
      await editor.mouse.move(box.x + (box.width * 18) / 300, box.y + (box.height * top) / 200)
      await editor.mouse.down()
      await editor.mouse.move(box.x + (box.width * 160) / 300, box.y + (box.height * bottom) / 200)
      await editor.mouse.up()
      await expect(editor.locator('.pdf-redaction-mark')).toHaveCount(1)
    }
    await mark(45, 75)
    await launched.app.evaluate(({ dialog }) => {
      dialog.showSaveDialog = async () => ({ canceled: true, filePath: '' })
    })
    await editor.getByRole('button', { name: 'Apply redactions', exact: true }).click()
    await expect(editor.locator('.app')).not.toHaveAttribute('inert')
    await expect(editor.locator('.pdf-redaction-mark')).toHaveCount(1)
    await launched.app.evaluate(({ dialog }, target) => {
      dialog.showSaveDialog = async () => ({ canceled: false, filePath: target })
    }, copy)
    await editor.getByRole('button', { name: 'Apply redactions', exact: true }).click()
    await expect(editor.locator('.pdf-redaction-mark')).toHaveCount(0)
    await expect(editor.locator('.textLayer')).not.toContainText('SECRET-ONE')
    await expect(editor.locator('.textLayer')).toContainText('SECRET-TWO')
    expect(await editor.evaluate(() => window.pdfApi.consumePending())).toBe(copy)
    const shellPage = await waitForPageWithUrl(launched.app, 'shell/out')
    await expect(shellPage.locator('.tab-title').filter({ hasText: 'redacted.pdf' })).toBeVisible()
    await launched.app.evaluate(({ dialog }) => {
      dialog.showSaveDialog = async () => {
        throw new Error('Unexpected second Save As')
      }
    })
    await mark(115, 145)
    await editor.getByRole('button', { name: 'Apply redactions', exact: true }).click()
    await expect(editor.locator('.pdf-redaction-mark')).toHaveCount(0)
    await expect(editor.locator('.textLayer')).not.toContainText('SECRET-ONE')
    await expect(editor.locator('.textLayer')).not.toContainText('SECRET-TWO')
    await expect(editor.locator('.textLayer')).toContainText('PUBLIC')
    expect(new Uint8Array(await readFile(source))).toEqual(original)
    await editor.reload()
    await expect(editor.locator('.textLayer')).toContainText('PUBLIC')
    await expect(editor.locator('.textLayer')).not.toContainText('SECRET')
  } finally {
    await closeAndSaveVideo(launched, 'pdf-redaction')
    await rm(dir, { recursive: true, force: true })
  }
})
