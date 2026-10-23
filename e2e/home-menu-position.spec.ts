import { test, expect } from '@playwright/test'
import { mkdirSync, mkdtempSync, realpathSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo } from './helpers'

/**
 * Row menus (the ⋯ button on folder rows) are position:fixed and anchor via
 * inline top/right computed from getBoundingClientRect against the layout
 * viewport. On scaled Linux desktops window.innerWidth can disagree with that
 * viewport and pushed the menus past the window's left edge (genoffice#1946); these
 * assertions pin the menu next to its button and inside the window on every
 * platform.
 */
test.describe('home row menu position', () => {
  let root: string

  test.beforeEach(() => {
    root = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-root-')))
    mkdirSync(join(root, 'Clients'), { recursive: true })
  })

  test.afterEach(() => {
    rmSync(root, { recursive: true, force: true })
  })

  test('the tree row menu opens below its button and inside the window', async () => {
    const launched = await launchShell({
      onboardingSeen: true,
      settings: { defaultSaveDir: root },
      videoDir: 'home-menu-position',
    })
    const { page } = launched
    try {
      const tree = page.locator('.folder-panel .tree')
      const row = tree.locator('.tree-row', {
        has: page.locator('.tree-name', { hasText: 'Clients' }),
      })
      await expect(row).toBeVisible()
      const more = row.locator('.folder-more-btn')
      await more.hover()
      await more.click()
      const menu = page.locator('.folder-menu')
      await expect(menu).toBeVisible()

      const {
        menu: m,
        button: b,
        viewport: vp,
      } = await page.evaluate(() => {
        const rect = (el: Element) => {
          const r = el.getBoundingClientRect()
          return { left: r.left, right: r.right, top: r.top, bottom: r.bottom }
        }
        return {
          menu: rect(document.querySelector('.folder-menu')!),
          button: rect(document.querySelector('.folder-more-btn[aria-expanded="true"]')!),
          viewport: {
            w: document.documentElement.clientWidth,
            h: document.documentElement.clientHeight,
          },
        }
      })
      // inside the window, right edge on the button's right edge, below it
      expect(m.left).toBeGreaterThanOrEqual(0)
      expect(m.right).toBeLessThanOrEqual(vp.w)
      expect(Math.abs(m.right - b.right)).toBeLessThanOrEqual(2)
      expect(m.top).toBeGreaterThanOrEqual(b.bottom - 1)
      expect(m.top).toBeLessThanOrEqual(b.bottom + 12)
    } finally {
      await closeAndSaveVideo(launched, 'home-menu-position')
    }
  })

  test('the tree right-click menu opens to the right of the pointer and inside the window', async () => {
    const launched = await launchShell({
      onboardingSeen: true,
      settings: { defaultSaveDir: root },
      videoDir: 'home-menu-position',
    })
    const { page } = launched
    try {
      const row = page.locator('.folder-panel .tree .tree-row', {
        has: page.locator('.tree-name', { hasText: 'Clients' }),
      })
      await expect(row).toBeVisible()
      const box = (await row.boundingBox())!
      const pointer = { x: box.x + 24, y: box.y + box.height / 2 }
      await page.mouse.click(pointer.x, pointer.y, { button: 'right' })
      const menu = page.locator('.folder-menu')
      await expect(menu).toBeVisible()
      const m = await page.evaluate(() => {
        const r = document.querySelector('.folder-menu')!.getBoundingClientRect()
        return {
          left: r.left,
          right: r.right,
          top: r.top,
          w: document.documentElement.clientWidth,
        }
      })
      // never cut off at the left edge, and it starts at the pointer instead of ending there
      expect(m.left).toBeGreaterThanOrEqual(0)
      expect(m.right).toBeLessThanOrEqual(m.w)
      expect(Math.abs(m.left - (pointer.x + 2))).toBeLessThanOrEqual(2)
    } finally {
      await closeAndSaveVideo(launched, 'home-menu-position')
    }
  })
})
