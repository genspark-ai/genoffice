// Slides menu: the Settings entry moves out of the File menu into the shared
// top-level Settings Center menu, placed immediately before Help; the ⌘/Ctrl+,
// accelerator stays on the item and the macOS app-menu entry stays untouched.
import { describe, expect, it, vi } from 'vitest'
import type { MenuItemConstructorOptions } from 'electron'

vi.mock('electron', () => ({
  app: {
    getPath: () => '/tmp',
    getName: () => 'slides-test',
    getVersion: () => '0',
    whenReady: () => Promise.resolve(),
  },
  dialog: {},
  clipboard: { writeBuffer: () => {} },
  ipcMain: {
    handle: () => {},
    on: () => {},
    removeHandler: () => {},
  },
  BrowserWindow: class {},
  webContents: { getAllWebContents: () => [], fromId: () => null },
  // return the template so the test can assert its structure
  Menu: {
    buildFromTemplate: (template: MenuItemConstructorOptions[]) => template,
    setApplicationMenu: () => {},
  },
  session: { defaultSession: { protocol: {}, setDisplayMediaRequestHandler: () => {} } },
  nativeImage: {},
  shell: {},
  desktopCapturer: { getSources: async () => [] },
  WebContentsView: class {},
}))
// The shaping chain loads a harfbuzz wasm that cannot resolve in the unit-test env
vi.mock('../src/main/shaped-metrics', () => ({
  shapedMetricsReady: async () => {},
  refineComplexWidths: () => {},
  complexScriptOf: () => null,
  initShapedMetrics: () => {},
  gtMeasure: (t: string) => ({ width: t.length * 8 }),
  shapedFamily: () => null,
  shapedMeasure: () => ({ width: 0 }),
}))

import { buildSlidesMenu } from '../src/main/slides-main'

const SETTINGS_ACCELERATOR = 'CmdOrCtrl+,'

function submenu(entry: MenuItemConstructorOptions): MenuItemConstructorOptions[] {
  return (entry.submenu as MenuItemConstructorOptions[] | undefined) ?? []
}

describe('slides menu (Settings Center)', () => {
  it('places the settings menu top-level immediately before Help', () => {
    const template = buildSlidesMenu() as unknown as MenuItemConstructorOptions[]

    // top-level entries carrying the settings item: the macOS app menu (kept)
    // plus the Settings Center menu — and nothing else
    const settingsTop = template.filter((m) =>
      submenu(m).some((i) => i.accelerator === SETTINGS_ACCELERATOR),
    )
    expect(settingsTop).toHaveLength(process.platform === 'darwin' ? 2 : 1)

    const settingsMenu = settingsTop.at(-1)!
    // immediately before Help, the last top-level entry (slides has no Window menu)
    expect(template[template.length - 2]).toBe(settingsMenu)
    expect(template.at(-1)!.role).toBe('help')

    // single item, accelerator unchanged
    const items = submenu(settingsMenu)
    expect(items).toHaveLength(1)
    expect(items[0]!.accelerator).toBe(SETTINGS_ACCELERATOR)
  })

  it('keeps the settings item out of the File menu', () => {
    const template = buildSlidesMenu() as unknown as MenuItemConstructorOptions[]

    const file = template.find((m) => submenu(m).some((i) => i.accelerator === 'CmdOrCtrl+O'))!
    expect(file).toBeDefined()
    expect(submenu(file).some((i) => i.accelerator === SETTINGS_ACCELERATOR)).toBe(false)
  })
})
