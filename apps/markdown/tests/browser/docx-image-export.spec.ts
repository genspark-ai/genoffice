import { expect, test } from '@playwright/test'
import { parseDocx } from '@genoffice/docx-engine'
import { openSource } from './helpers'

const SVG =
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1600 400"><rect width="1600" height="400" fill="#123456"/></svg>'

test('an SVG picture is rasterized into the docx instead of collapsing to alt text', async ({
  page,
}) => {
  await openSource(page, false, false, '# Export\n\n![demo](demo.svg)\n\ntail\n')
  await page.evaluate((svgBase64) => {
    window.markdownApi.readImage = async (src) =>
      src === 'demo.svg' ? { base64: svgBase64, mime: 'image/svg+xml' } : null
    window.markdownApi.exportDocx = async (request) => {
      document.body.dataset.docx = request.base64
      return { ok: true, path: '/export/demo.docx' }
    }
    window.dispatchEvent(new CustomEvent('test:export', { detail: 'docx' }))
  }, Buffer.from(SVG).toString('base64'))
  await expect(page.locator('body')).toHaveAttribute('data-docx', /.+/)
  const base64 = await page.locator('body').getAttribute('data-docx')
  const parsed = await parseDocx(new Uint8Array(Buffer.from(base64!, 'base64')))
  const image = parsed.blocks.find((b) => b.type === 'image')
  expect(image?.imageDataUrl).toMatch(/^data:image\/png;base64,/)
  expect(image?.imageWidthPx).toBe(620)
  expect(image?.imageHeightPx).toBe(155)
  const texts = parsed.blocks.map((b) => (b.runs ?? []).map((r) => r.text).join(''))
  expect(texts).not.toContain('[demo]')
})
