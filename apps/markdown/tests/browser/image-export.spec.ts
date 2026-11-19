import { expect, test } from '@playwright/test'
import { PDFDocument } from 'pdf-lib'
import { openSource } from './helpers'

for (const outcome of ['canceled', 'failed', 'completed']) {
  const canceled = outcome === 'canceled'
  test(`PNG export ${outcome} with a broken image`, async ({ page }) => {
    const pdf = await PDFDocument.create()
    for (let i = 0; i < 2; i++)
      pdf.addPage([120, 120]).drawRectangle({ x: 10, y: 10, width: 80, height: 80 })
    const pdfBase64 = Buffer.from(await pdf.save()).toString('base64')
    await openSource(page, false, false, '# Export\n\n![broken](/missing-export-image.png)\n')
    await page.evaluate(
      ({ canceled, outcome, pdfBase64 }) => {
        window.markdownApi.prepareImageExport = async () => {
          document.body.dataset.picked = 'true'
          return canceled ? { ok: true, canceled: true } : { ok: true, id: 'fixture', pdfBase64 }
        }
        window.markdownApi.writeExportImage = async (_id, pageNumber, bytes) => {
          document.body.dataset.pngHeader = bytes.slice(0, 11)
          return pageNumber === 1 || outcome === 'completed'
            ? { ok: true }
            : { ok: false, error: 'Disk is full' }
        }
        window.markdownApi.finishImageExport = async (_id, success) => {
          document.body.dataset.finished = String(success)
          return success
            ? { ok: true, path: '/export/report-images' }
            : { ok: true, canceled: true }
        }
        window.dispatchEvent(new CustomEvent('test:export', { detail: 'png' }))
      },
      { canceled, outcome, pdfBase64 },
    )
    await expect(page.locator('body')).toHaveAttribute('data-picked', 'true')
    if (canceled) {
      await expect(page.locator('.status-export')).toHaveCount(0)
      await expect(page.locator('body')).not.toHaveAttribute('data-png-header')
      await expect(page.locator('body')).not.toHaveAttribute('data-finished')
    } else if (outcome === 'completed') {
      await expect(page.locator('.status-export')).toContainText('/export/report-images')
      await expect(page.locator('body')).toHaveAttribute('data-png-header', 'iVBORw0KGgo')
      await expect(page.locator('body')).toHaveAttribute('data-finished', 'true')
    } else {
      await expect(page.locator('.status-export')).toHaveText('Image export failed: Disk is full')
      await expect(page.locator('body')).toHaveAttribute('data-png-header', 'iVBORw0KGgo')
      await expect(page.locator('body')).toHaveAttribute('data-finished', 'false')
    }
  })
}
