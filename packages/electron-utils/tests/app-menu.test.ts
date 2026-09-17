import { beforeEach, describe, expect, it, vi } from 'vitest'

const electronMock = vi.hoisted(() => ({
  getFocusedWebContents: vi.fn(),
  showMessageBox: vi.fn(),
  writeText: vi.fn(),
}))

vi.mock('electron', () => ({
  webContents: {
    getFocusedWebContents: electronMock.getFocusedWebContents,
  },
  app: {
    getVersion: () => '1.2.3',
  },
  dialog: {
    showMessageBox: electronMock.showMessageBox,
  },
  clipboard: {
    writeText: electronMock.writeText,
  },
}))

import {
  aboutMenuItem,
  appMenuLabels,
  checkUpdatesMenuItem,
  editMenuTemplate,
  helpMenuTemplate,
  setUpdateCheckInvoker,
  viewMenuTemplate,
  windowMenuTemplate,
  type AppMenuLabels,
} from '../src/index'

const LANGS = [
  'zh',
  'en',
  'ja',
  'ko',
  'fr',
  'de',
  'es',
  'th',
  'id',
  'ru',
  'ar',
  'pt',
  'it',
  'pl',
  'cs',
  'nl',
  'ms',
  'he',
  'hi',
  'zh-TW',
] as const

const en = appMenuLabels('en')

type Item = { role?: string; type?: string; label?: string; accelerator?: string }
const submenuOf = (tpl: { submenu?: unknown }): Item[] => tpl.submenu as Item[]

beforeEach(() => {
  vi.clearAllMocks()
  setUpdateCheckInvoker(null)
})

describe('appMenuLabels', () => {
  it('covers all 20 languages with every key non-empty', () => {
    const keys = Object.keys(en) as (keyof AppMenuLabels)[]
    for (const lang of LANGS) {
      const labels = appMenuLabels(lang)
      for (const key of keys) expect(labels[key], `${lang}.${key}`).toBeTruthy()
    }
  })

  it('falls back to English for unknown languages', () => {
    expect(appMenuLabels('xx')).toEqual(en)
  })

  it('merges the context-menu clipboard labels', () => {
    const zh = appMenuLabels('zh')
    expect(zh.copy).toBe('复制')
    expect(zh.selectAll).toBe('全选')
  })
})

describe('windowMenuTemplate', () => {
  it('keeps the native role on macOS', () => {
    expect(windowMenuTemplate('darwin', en)).toEqual({ role: 'windowMenu', label: 'Window' })
  })

  it.each(['win32', 'linux'] as const)('builds Minimize/Close only on %s', (platform) => {
    const tpl = windowMenuTemplate(platform, en)
    expect(tpl.role).toBeUndefined()
    expect(tpl.label).toBe('Window')
    const items = submenuOf(tpl)
    expect(items.map((i) => i.label ?? i.type)).toEqual(['Minimize', 'separator', 'Close Window'])
    // no mac-only entries, no roles, no Ctrl+M
    for (const item of items) {
      expect(item.role).toBeUndefined()
      expect(item.accelerator).toBeUndefined()
      expect(item.label ?? '').not.toMatch(/zoom|front/i)
    }
  })

  it('localizes the Windows submenu', () => {
    const items = submenuOf(windowMenuTemplate('win32', appMenuLabels('zh')))
    expect(items[0]!.label).toBe('最小化')
    expect(items[2]!.label).toBe('关闭窗口')
  })
})

describe('editMenuTemplate', () => {
  it('keeps the native role on macOS', () => {
    expect(editMenuTemplate('darwin', en)).toEqual({ role: 'editMenu', label: 'Edit' })
  })

  it('mirrors the non-mac role expansion with labeled roles on Windows', () => {
    const tpl = editMenuTemplate('win32', en)
    expect(tpl.label).toBe('Edit')
    expect(submenuOf(tpl).map((i) => i.role ?? i.type)).toEqual([
      'undo',
      'redo',
      'separator',
      'cut',
      'copy',
      'paste',
      'delete',
      'separator',
      'selectAll',
    ])
    for (const item of submenuOf(tpl)) if (item.role) expect(item.label).toBeTruthy()
  })
})

describe('viewMenuTemplate', () => {
  it('mirrors the viewMenu role expansion with labeled roles', () => {
    const tpl = viewMenuTemplate(en)
    expect(tpl.label).toBe('View')
    expect(submenuOf(tpl).map((i) => i.role ?? i.type ?? i.label)).toEqual([
      'reload',
      'forceReload',
      'Developer Tools',
      'separator',
      'resetZoom',
      'zoomIn',
      'zoomOut',
      'separator',
      'togglefullscreen',
    ])
    for (const item of submenuOf(tpl)) if (item.role) expect(item.label).toBeTruthy()
  })

  it('opens DevTools via an explicit click item, not the docking role', () => {
    const item = submenuOf(viewMenuTemplate(en))[2] as Item & { click?: unknown }
    expect(item.role).toBeUndefined()
    expect(typeof item.click).toBe('function')
    expect(item.accelerator).toBe(process.platform === 'darwin' ? 'Alt+Command+I' : 'Ctrl+Shift+I')
  })

  it('closes detached DevTools when they have focus', async () => {
    const target = {
      devToolsWebContents: {},
      isDestroyed: vi.fn(() => false),
      isDevToolsOpened: vi.fn().mockReturnValueOnce(false).mockReturnValueOnce(true),
      openDevTools: vi.fn(),
      closeDevTools: vi.fn(),
    }
    electronMock.getFocusedWebContents.mockReturnValueOnce(target).mockReturnValueOnce(null)
    const item = submenuOf(viewMenuTemplate(en))[2] as Item & {
      click: () => Promise<void>
    }

    await item.click()
    await item.click()

    expect(target.openDevTools).toHaveBeenCalledWith({ mode: 'detach' })
    expect(target.closeDevTools).toHaveBeenCalledOnce()
  })
})

describe('checkUpdatesMenuItem / manual update check wiring', () => {
  it('sits directly above About in the help menu', () => {
    const items = submenuOf(helpMenuTemplate(en))
    expect(items.at(-2)!.label).toBe(en.checkUpdates)
    expect(items.at(-1)!.label).toBe(en.about)
  })

  it('invokes the shell-registered check on click, and no-ops unregistered', () => {
    const item = checkUpdatesMenuItem(en) as { label: string; click: () => void }
    expect(item.label).toBe(en.checkUpdates)
    expect(() => item.click()).not.toThrow()

    const invoke = vi.fn()
    setUpdateCheckInvoker(invoke)
    item.click()
    expect(invoke).toHaveBeenCalledOnce()
  })

  it('About dialog offers the check button only when a check is registered', async () => {
    electronMock.showMessageBox.mockResolvedValue({ response: 0 })
    const about = aboutMenuItem(en) as { click: () => Promise<void> }

    await about.click()
    expect(electronMock.showMessageBox.mock.calls.at(-1)![0].buttons).toEqual(['OK', en.copy])

    setUpdateCheckInvoker(() => {})
    await about.click()
    expect(electronMock.showMessageBox.mock.calls.at(-1)![0].buttons).toEqual([
      'OK',
      en.copy,
      en.checkUpdates,
    ])
  })

  it('About dialog third button triggers the registered check', async () => {
    const invoke = vi.fn()
    setUpdateCheckInvoker(invoke)
    electronMock.showMessageBox.mockResolvedValue({ response: 2 })
    await (aboutMenuItem(en) as { click: () => Promise<void> }).click()
    expect(invoke).toHaveBeenCalledOnce()
    expect(electronMock.writeText).not.toHaveBeenCalled()
  })

  it('About dialog copy button still copies name + version', async () => {
    electronMock.showMessageBox.mockResolvedValue({ response: 1 })
    await (aboutMenuItem(en) as { click: () => Promise<void> }).click()
    expect(electronMock.writeText).toHaveBeenCalledWith('GenOffice 1.2.3')
  })
})
