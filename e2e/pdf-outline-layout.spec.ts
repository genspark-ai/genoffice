import { test, expect } from '@playwright/test'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'

test('long PDF outlines scroll without compressing entries', async ({ page }) => {
  const css = readFileSync(resolve('apps/pdf/src/renderer/styles.css'), 'utf8')
  await page.setContent(`<!doctype html>
    <style>${css}</style>
    <div class="pdf-thumbs pdf-outline-pane" style="height:300px;width:240px">
      <div class="pdf-outline-header">Outline<button class="rb-icon">Collapse</button></div>
      <div class="pdf-outline">
        <div class="pdf-outline-note">Generated outline</div>
        ${Array.from({ length: 40 }, (_, i) => `<button class="pdf-outline-item">Chapter ${i + 1}</button>`).join('')}
      </div>
    </div>
  `)
  const item = page.locator('.pdf-outline-item').first()
  const height = await item.evaluate((el) => el.getBoundingClientRect().height)
  const lineHeight = await item.evaluate((el) => parseFloat(getComputedStyle(el).lineHeight))
  expect(height).toBeGreaterThanOrEqual(lineHeight + 10)
  const header = await page.locator('.pdf-outline-header').boundingBox()
  await page.locator('.pdf-outline-item').last().click()
  expect(await page.locator('.pdf-outline').evaluate((el) => el.scrollTop)).toBeGreaterThan(0)
  expect(await page.locator('.pdf-outline-header').boundingBox()).toEqual(header)
})
