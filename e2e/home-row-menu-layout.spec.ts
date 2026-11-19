import { test, expect } from '@playwright/test'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'

// Exercise the production menu CSS in a real browser: jsdom cannot lay out anchors.
const css = readFileSync(resolve('apps/shell/src/renderer/src/home.css'), 'utf8')

for (const scenario of [
  { name: 'bottom row flips upward', height: 600, top: 550, above: true },
  { name: 'top row opens downward', height: 600, top: 20, above: false },
  { name: 'short window keeps every action reachable', height: 220, top: 100 },
]) {
  test(scenario.name, async ({ page }) => {
    await page.setViewportSize({ width: 800, height: scenario.height })
    await page.setContent(`<!doctype html>
      <style>${css}</style>
      <div style="height: calc(100vh - 16px)">
        <span class="recent-actions" style="position:absolute;top:${scenario.top}px;right:20px">
          <button class="more-btn">…</button>
        </span>
      </div>
    `)
    await page.locator('.more-btn').evaluate((button) => {
      button.addEventListener('click', () => {
        const menu = document.createElement('div')
        menu.className = 'row-menu'
        menu.setAttribute('role', 'menu')
        for (const label of [
          'Open',
          'Reveal',
          'Copy path',
          'Move',
          'Rename',
          'Duplicate',
          'Remove',
          'Delete',
        ]) {
          const item = document.createElement('button')
          item.textContent = label
          menu.append(item)
        }
        button.after(menu)
      })
    })
    const initialHeight = await page.evaluate(() => document.documentElement.scrollHeight)
    await page.locator('.more-btn').click()
    const menu = page.locator('.row-menu')
    await expect(menu).toBeVisible()
    await expect
      .poll(async () => {
        const box = await menu.boundingBox()
        return !!box && box.y >= 0 && box.y + box.height <= scenario.height
      })
      .toBe(true)
    const box = (await menu.boundingBox())!
    const trigger = (await page.locator('.more-btn').boundingBox())!
    if (scenario.above === true) expect(box.y + box.height).toBeLessThanOrEqual(trigger.y)
    if (scenario.above === false) expect(box.y).toBeGreaterThanOrEqual(trigger.y + trigger.height)
    await menu.getByText('Delete', { exact: true }).click()
    expect(await page.evaluate(() => window.scrollY)).toBe(0)
    expect(await page.evaluate(() => document.documentElement.scrollHeight)).toBe(initialHeight)
    if (scenario.above === false) {
      await page.setViewportSize({ width: 800, height: 220 })
      await expect
        .poll(async () => {
          const resized = await menu.boundingBox()
          return !!resized && resized.y >= 0 && resized.y + resized.height <= 220
        })
        .toBe(true)
      await menu.getByText('Delete', { exact: true }).click()
      expect(await page.evaluate(() => window.scrollY)).toBe(0)
    }
  })
}
