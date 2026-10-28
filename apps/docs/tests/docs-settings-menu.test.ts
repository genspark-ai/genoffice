// The docs application menu moves its non-mac Settings entry from the Tools
// menu into the top-level Settings Center menu (settingsMenuTemplate), inserted
// immediately before the Window menu; the macOS app-menu entry stays where Word
// keeps it. docs-main is the full Electron bootstrap, so the wiring is asserted
// against its source (same pattern as open-file.test.ts), while the template's
// own shape is exercised directly.
import { describe, expect, it, vi } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'

vi.mock('electron', () => ({}))

const { appMenuLabels, settingsMenuTemplate } =
  await import('../../../packages/electron-utils/src/app-menu')

const source = readFileSync(join(import.meta.dirname, '../src/main/docs-main.ts'), 'utf8')

describe('docs-main settings menu wiring', () => {
  it('imports settingsMenuTemplate from @genoffice/electron-utils', () => {
    expect(source).toMatch(
      /import\s*\{[^}]*\bsettingsMenuTemplate\b[^}]*\}\s*from\s*'@genoffice\/electron-utils'/s,
    )
  })

  it('keeps exactly one settingsMenuItem: the macOS application-menu entry', () => {
    // the app-menu extra (Word keeps Settings under the app name on macOS)
    expect(source).toContain("extra: [settingsMenuItem(labels, () => sendCommand('preferences'))]")
    // the non-mac Tools-menu placement is gone
    expect(source).not.toMatch(/\.\.\.\(isMac \? \[\] : \[settingsMenuItem\(/)
    expect(source.match(/settingsMenuItem\(/g)).toHaveLength(1)
  })

  it('inserts the top-level Settings Center menu immediately before the Window menu', () => {
    const settingsMenu = source.indexOf('settingsMenuTemplate(labels,')
    const windowMenu = source.indexOf('windowMenuTemplate(process.platform')
    const toolsMenu = source.indexOf("label: tm('menuTools')")
    expect(toolsMenu).toBeGreaterThan(-1)
    expect(settingsMenu).toBeGreaterThan(toolsMenu)
    expect(settingsMenu).toBeLessThan(windowMenu)
    // the fallback action is preserved (shell-registered invoker wins when present)
    expect(source).toContain("settingsMenuTemplate(labels, () => sendCommand('preferences'))")
  })
})

describe('settingsMenuTemplate', () => {
  it('carries the Settings item with the Ctrl/Cmd+, accelerator under the Settings Center title', () => {
    const labels = appMenuLabels('en')
    const template = settingsMenuTemplate(labels)
    expect(template.label).toBe(labels.settingsMenu)
    const items = template.submenu as Array<Record<string, unknown>>
    expect(items).toHaveLength(1)
    expect(items[0].label).toBe(labels.settings)
    expect(items[0].accelerator).toBe('CmdOrCtrl+,')
  })

  it('routes the item click to the fallback when no shell invoker is registered', async () => {
    const fallback = vi.fn()
    const template = settingsMenuTemplate(appMenuLabels('en'), fallback)
    const item = (template.submenu as Array<{ click?: () => void }>)[0]
    expect(item.click).toBeTypeOf('function')
    await item.click!()
    expect(fallback).toHaveBeenCalledTimes(1)
  })
})
