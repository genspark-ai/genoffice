import { test, expect, type ElectronApplication } from '@playwright/test'
import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * Pins the whole Docs menu bar (labels, accelerators, roles, check/enable
 * state, nesting) so a refactor of how the menu is built can prove it changed
 * nothing. Regenerate with UPDATE_E2E_BASELINES=1 when a change is intended.
 */

const BASELINE = join(__dirname, 'baselines', 'docs-menu-tree.json')
const MENUS = ['File', 'Edit', 'View', 'Insert', 'Format', 'Tools', 'Table']

interface MenuNode {
  label: string
  type?: string
  role?: string
  accelerator?: string
  disabled?: true
  checked?: boolean
  submenu?: MenuNode[]
}

const menuTree = (app: ElectronApplication, menus: string[]): Promise<MenuNode[]> =>
  app.evaluate(({ Menu }, menus) => {
    type Node = {
      label: string
      type?: string
      role?: string
      accelerator?: string
      disabled?: true
      checked?: boolean
      submenu?: Node[]
    }
    const walk = (item: Electron.MenuItem): Node => {
      const node: Node = { label: item.label }
      if (item.type !== 'normal' && item.type !== 'submenu') node.type = item.type
      if (item.role) node.role = item.role
      // a role item's accelerator is Electron's per-platform default, not ours
      if (item.accelerator && !item.role) node.accelerator = String(item.accelerator)
      if (!item.enabled) node.disabled = true
      if (item.type === 'checkbox' || item.type === 'radio') node.checked = item.checked
      // the recent-files list is per-machine state, not menu structure
      if (item.submenu && item.label !== 'Open Recent') node.submenu = item.submenu.items.map(walk)
      return node
    }
    return (Menu.getApplicationMenu()?.items ?? [])
      .filter((item) => menus.includes(item.label))
      .map(walk)
  }, menus)

// macOS keeps Settings in the application menu and binds Go To differently;
// fold the other platforms onto the macOS shape so one baseline serves both
function toMacShape(nodes: MenuNode[]): MenuNode[] {
  return nodes
    .filter((n) => n.accelerator !== 'CmdOrCtrl+,')
    .map((n) => ({
      ...n,
      ...(n.accelerator === 'Ctrl+G' ? { accelerator: 'Alt+Cmd+G' } : {}),
      ...(n.submenu ? { submenu: toMacShape(n.submenu) } : {}),
    }))
}

test('the Docs menu bar matches its baseline', async () => {
  test.setTimeout(120_000)
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-menu-tree' })
  const { page, app } = launched
  try {
    await page.evaluate(() =>
      (window as never as { aiOffice: { newDoc(): Promise<void> } }).aiOffice.newDoc(),
    )
    await waitForPageWithUrl(app, '://docs/')
    await expect
      .poll(async () => (await menuTree(app, MENUS)).map((m) => m.label), { timeout: 30_000 })
      .toEqual(MENUS)
    // what is greyed out depends on where the focus is: put the caret in the document
    const docPage = await waitForPageWithUrl(app, '://docs/')
    await docPage.locator('.ProseMirror').first().click()
    const read = async () => JSON.stringify(toMacShape(await menuTree(app, MENUS)), null, 2) + '\n'
    if (process.env.UPDATE_E2E_BASELINES || !existsSync(BASELINE)) {
      // the renderer's first state report greys items out a moment after the menu is built
      await page.waitForTimeout(3000)
      writeFileSync(BASELINE, await read())
    }
    // polled: enabled and checked state arrive from the renderer after the menu is built
    await expect.poll(read, { timeout: 15_000 }).toBe(readFileSync(BASELINE, 'utf8'))
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-tree')
  }
})
