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
const MENUS = ['File', 'Edit', 'View', 'Insert', 'Format', 'Tools', 'Table', 'Window']

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

/** the chords Word for Mac alone has, by the path of the item that shows them */
const MAC_ONLY_CHORDS: Record<string, string> = {
  'File/New from Template…': 'Shift+Cmd+P',
  'Edit/Paste Special…': 'Ctrl+Cmd+V',
  'Edit/Find/Find Next': 'Cmd+G',
  'Edit/Find/Find Previous': 'Shift+Cmd+G',
  'Edit/Find/Replace…': 'Shift+Cmd+H',
  'Edit/Toggle Drawing': 'Ctrl+Cmd+Z',
  'View/Focus': 'Ctrl+Shift+Cmd+F',
  'View/Ribbon': 'Alt+Cmd+R',
  'Insert/Equation…': 'Ctrl+=',
  'Format/Format Object…': 'Shift+Cmd+1',
  'Format/Selection Pane…': 'Alt+Cmd+U',
  'Tools/Thesaurus…': 'Ctrl+Alt+Cmd+R',
  'Table/Delete/Rows': 'Ctrl+Cmd+X',
}

/** items whose chord Word for Mac gives to another command (Normal style, Replace) */
const NO_CHORD_ON_MAC = new Set(['Window/New Window', 'File/Back to Home'])

// the shell's interface zoom is Cmd+Ctrl on macOS; the baseline holds the chord of the other platforms
function interfaceZoomAsElsewhere(node: MenuNode): MenuNode {
  const key = /^Cmd\+Ctrl\+([-=0])$/.exec(node.accelerator ?? '')?.[1]
  return {
    ...node,
    ...(key ? { accelerator: `CmdOrCtrl+Alt+${key}` } : {}),
    ...(node.submenu ? { submenu: node.submenu.map(interfaceZoomAsElsewhere) } : {}),
  }
}

// macOS keeps Settings in the application menu, binds Go To differently and has
// Word for Mac's own chords; fold the other platforms onto the macOS shape so
// one baseline serves both
function toMacShape(nodes: MenuNode[], trail = ''): MenuNode[] {
  if (process.platform === 'darwin') return nodes.map(interfaceZoomAsElsewhere)
  return (
    nodes
      // Preferences and the suite's Settings are in the application menu on macOS
      .filter((n) => n.accelerator !== 'CmdOrCtrl+,' && n.label !== 'Settings…')
      .map((n) => {
        const path = trail + n.label
        const macOnly = n.accelerator ? undefined : MAC_ONLY_CHORDS[path]
        // the key order of the walk: label, type, role, accelerator, disabled, checked, submenu
        const { disabled, checked, submenu, ...head } = n
        if (NO_CHORD_ON_MAC.has(path)) delete head.accelerator
        return {
          ...head,
          ...(n.accelerator === 'Ctrl+G' ? { accelerator: 'Alt+Cmd+G' } : {}),
          ...(path === 'Edit/Paste and Match Formatting'
            ? { accelerator: 'CmdOrCtrl+Alt+Shift+V' }
            : {}),
          ...(macOnly ? { accelerator: macOnly } : {}),
          ...(disabled ? { disabled } : {}),
          ...(checked === undefined ? {} : { checked }),
          ...(submenu ? { submenu: toMacShape(submenu, `${path}/`) } : {}),
        }
      })
  )
}

// the Window menu begins and ends with the system's own items, which differ by platform:
// the baseline holds the commands between them
function windowCommands(nodes: MenuNode[]): MenuNode[] {
  return nodes.map((top) => {
    if (top.label !== 'Window' || !top.submenu) return top
    const ours = top.submenu.filter(
      (item) =>
        !item.role &&
        !['Minimize', 'Close Window', 'Minimize All', 'Zoom All'].includes(item.label ?? ''),
    )
    while (ours[0]?.type === 'separator') ours.shift()
    while (ours.at(-1)?.type === 'separator') ours.pop()
    // the menu's own role is macOS' too
    return { label: top.label, submenu: ours }
  })
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
    const read = async () =>
      JSON.stringify(toMacShape(windowCommands(await menuTree(app, MENUS))), null, 2) + '\n'
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
