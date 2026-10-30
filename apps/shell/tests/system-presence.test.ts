import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

/**
 * System presence: the prefs store invariant, the accelerator shape check, the
 * deep-link parser, and the install/IPC behavior against a mocked Electron
 * (the desktop-ipc.test.ts style — the module never sees a real Tray or
 * globalShortcut).
 */

const state = vi.hoisted(() => ({
  trays: [] as Array<{
    icon: unknown
    menu: unknown
    tooltip: string
    destroyed: boolean
  }>,
  registered: new Map<string, () => void>(),
  loginItems: [] as boolean[],
  quitCalls: 0,
  handlers: new Map<string, (...args: never[]) => unknown>(),
  appListeners: new Map<string, Array<(...args: never[]) => void>>(),
  protocolClients: [] as string[],
  refuseRegistrations: false,
  unregisterCalls: 0,
}))

vi.mock('electron', () => ({
  app: {
    getPath: () => '/tmp/genoffice-test-unused',
    quit: () => {
      state.quitCalls += 1
    },
    setLoginItemSettings: (settings: { openAtLogin: boolean }) => {
      state.loginItems.push(settings.openAtLogin)
    },
    setAsDefaultProtocolClient: (scheme: string) => {
      state.protocolClients.push(scheme)
    },
    on: (channel: string, listener: (...args: never[]) => void) => {
      const list = state.appListeners.get(channel) ?? []
      list.push(listener)
      state.appListeners.set(channel, list)
    },
  },
  ipcMain: {
    handle: (channel: string, handler: (...args: never[]) => unknown) =>
      state.handlers.set(channel, handler),
  },
  Menu: { buildFromTemplate: (template: unknown) => template },
  Tray: class {
    menu: unknown
    tooltip = ''
    destroyed = false
    constructor(public icon: unknown) {
      state.trays.push(this as never)
    }
    setToolTip(tooltip: string) {
      this.tooltip = tooltip
    }
    setContextMenu(menu: unknown) {
      this.menu = menu
    }
    destroy() {
      this.destroyed = true
    }
  },
  globalShortcut: {
    register: (accelerator: string, callback: () => void) => {
      // the flag simulates an accelerator another app already holds: the OS
      // refuses the registration and the presence must surface that
      if (state.refuseRegistrations) return false
      state.registered.set(accelerator, callback)
      return true
    },
    unregister: (accelerator: string) => {
      state.unregisterCalls += 1
      state.registered.delete(accelerator)
    },
    unregisterAll: () => {
      state.registered.clear()
    },
  },
}))

import { HOME_CHANNELS } from '../src/shared/home-api'
import {
  DEFAULT_SYSTEM_PRESENCE_PREFS,
  deepLinksFromArgv,
  installDeepLinkListener,
  installSystemPresence,
  isAcceleratorLike,
  normalizeSystemPresencePrefs,
  parseDeepLink,
  registerDeepLinkClient,
  registerSystemPresenceIpc,
  SYSTEM_PRESENCE_KEY,
} from '../src/main/system-presence'

const fakeIcon = () => ({ fake: true }) as never

let dir: string
let settingsPath: string

const makePresence = (overrides: Partial<Parameters<typeof installSystemPresence>[0]> = {}) => {
  const calls = { show: 0, palette: 0, ask: [] as Array<string | undefined> }
  const presence = installSystemPresence({
    settingsPath: () => settingsPath,
    lang: () => 'en',
    trayIcon: fakeIcon,
    showMainWindow: () => {
      calls.show += 1
    },
    summonPalette: () => {
      calls.palette += 1
    },
    openAsk: (query) => {
      calls.ask.push(query)
    },
    ...overrides,
  })
  return { presence, calls }
}

const invoke = (channel: string, ...args: unknown[]) =>
  state.handlers.get(channel)!({ sender: { id: 1 } } as never, ...args)

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'genoffice-presence-'))
  settingsPath = join(dir, 'app-settings.json')
  state.trays.length = 0
  state.registered.clear()
  state.loginItems.length = 0
  state.quitCalls = 0
  state.handlers.clear()
  state.appListeners.clear()
  state.protocolClients.length = 0
  state.refuseRegistrations = false
  state.unregisterCalls = 0
})

afterEach(() => {
  rmSync(dir, { recursive: true, force: true })
})

describe('parseDeepLink', () => {
  it('parses genoffice://ask with and without a query', () => {
    expect(parseDeepLink('genoffice://ask?q=hello')).toEqual({ view: 'ask', query: 'hello' })
    expect(parseDeepLink('genoffice://ask?q=hello%20world')).toEqual({
      view: 'ask',
      query: 'hello world',
    })
    expect(parseDeepLink('genoffice://ask')).toEqual({ view: 'ask' })
  })

  it('parses the palette and open verbs', () => {
    expect(parseDeepLink('genoffice://palette')).toEqual({ view: 'palette' })
    expect(parseDeepLink('genoffice://open?path=%2Ftmp%2Fa.docx')).toEqual({
      view: 'open',
      path: '/tmp/a.docx',
    })
    // a link without a path has nothing to open; refusing it beats guessing
    expect(parseDeepLink('genoffice://open')).toBeNull()
  })

  it('rejects other schemes, hosts and garbage', () => {
    expect(parseDeepLink('https://ask?q=hi')).toBeNull()
    expect(parseDeepLink('genoffice://nope?path=/x')).toBeNull()
    expect(parseDeepLink('not a url')).toBeNull()
    expect(parseDeepLink('')).toBeNull()
  })
})

describe('deepLinksFromArgv', () => {
  it('finds every genoffice:// entry a Windows/Linux protocol launch carries, in order', () => {
    expect(deepLinksFromArgv(['/app/genoffice', 'genoffice://ask?q=hi'])).toEqual([
      'genoffice://ask?q=hi',
    ])
    expect(
      deepLinksFromArgv([
        '/app/genoffice',
        'genoffice://palette',
        '/tmp/a.docx',
        'genoffice://ask',
      ]),
    ).toEqual(['genoffice://palette', 'genoffice://ask'])
    expect(deepLinksFromArgv(['/app/genoffice', '/tmp/a.docx'])).toEqual([])
  })
})

describe('isAcceleratorLike', () => {
  it('accepts modifiers + one key', () => {
    for (const ok of [
      'Alt+Space',
      'CommandOrControl+Shift+K',
      'CmdOrCtrl+Alt+P',
      'Control+F5',
      'Shift+Command+3',
    ]) {
      expect(isAcceleratorLike(ok), ok).toBe(true)
    }
  })

  it('rejects bare keys, empty tails, unknown modifiers and junk', () => {
    for (const bad of [
      '',
      'Space',
      'K',
      'Alt+',
      'Bogus+K',
      'Alt+',
      42,
      null,
      ['Alt+Space'],
      `Alt+${'x'.repeat(80)}`,
    ]) {
      expect(isAcceleratorLike(bad), String(bad)).toBe(false)
    }
  })
})

describe('normalizeSystemPresencePrefs', () => {
  it('falls back per field, never trusting a bad accelerator', () => {
    expect(normalizeSystemPresencePrefs(undefined)).toEqual(DEFAULT_SYSTEM_PRESENCE_PREFS)
    expect(normalizeSystemPresencePrefs('junk')).toEqual(DEFAULT_SYSTEM_PRESENCE_PREFS)
    expect(normalizeSystemPresencePrefs({ loginItem: true })).toEqual({
      ...DEFAULT_SYSTEM_PRESENCE_PREFS,
      loginItem: true,
    })
    expect(
      normalizeSystemPresencePrefs({ hotkeyEnabled: false, hotkeyAccelerator: 'not a hotkey' }),
    ).toEqual({
      ...DEFAULT_SYSTEM_PRESENCE_PREFS,
      hotkeyEnabled: false,
    })
    expect(normalizeSystemPresencePrefs({ hotkeyAccelerator: ' Control+Shift+P ' })).toEqual({
      ...DEFAULT_SYSTEM_PRESENCE_PREFS,
      hotkeyAccelerator: 'Control+Shift+P',
    })
  })
})

describe('installSystemPresence', () => {
  it('creates the tray with the localized quick actions and registers the default hotkey', () => {
    const { presence } = makePresence()
    expect(state.trays).toHaveLength(1)
    expect(state.trays[0]!.tooltip).toBe('GenOffice')
    const menu = state.trays[0]!.menu as Array<{ label?: string; type?: string }>
    expect(menu.map((item) => item.label ?? item.type)).toEqual([
      'Show GenOffice',
      'Command Palette',
      'Ask AI',
      'separator',
      'Quit GenOffice',
    ])
    expect([...state.registered.keys()]).toEqual(['CommandOrControl+Shift+Space'])
    expect(presence.settings().hotkeyActive).toBe(true)
    presence.dispose()
  })

  it('wires the tray actions and the hotkey to the injected callbacks', () => {
    const { presence, calls } = makePresence()
    const menu = state.trays[0]!.menu as Array<{ label?: string; click?: () => void }>
    menu[0]!.click!()
    expect(calls.show).toBe(1)
    menu[1]!.click!()
    expect(calls.palette).toBe(1)
    menu[2]!.click!()
    expect(calls.ask).toEqual([undefined])
    menu[4]!.click!()
    expect(state.quitCalls).toBe(1)
    // the registered hotkey summons the palette
    state.registered.get('CommandOrControl+Shift+Space')!()
    expect(calls.palette).toBe(2)
    presence.dispose()
  })

  it('respects persisted prefs: no hotkey when disabled, login item applied on mac/windows', () => {
    writeFileSync(
      settingsPath,
      JSON.stringify({
        [SYSTEM_PRESENCE_KEY]: { loginItem: true, hotkeyEnabled: false },
      }),
    )
    const { presence } = makePresence()
    expect(state.registered.size).toBe(0)
    expect(presence.settings().hotkey.enabled).toBe(false)
    if (process.platform === 'darwin' || process.platform === 'win32') {
      expect(state.loginItems).toEqual([true])
    } else {
      // Linux: persisted but not applied (no single API), and reported unsupported
      expect(state.loginItems).toEqual([])
      expect(presence.settings().loginItemSupported).toBe(false)
    }
    presence.dispose()
  })

  it('applyPrefs persists, re-registers the hotkey and re-applies the login item', () => {
    const { presence } = makePresence()
    const next = presence.applyPrefs({
      loginItem: true,
      closeToTray: false,
      hotkeyEnabled: true,
      hotkeyAccelerator: 'Control+Shift+K',
    })
    expect(state.registered.has('CommandOrControl+Shift+Space')).toBe(false)
    expect([...state.registered.keys()]).toEqual(['Control+Shift+K'])
    expect(next.hotkey).toEqual({ enabled: true, accelerator: 'Control+Shift+K' })
    const onDisk = JSON.parse(readFileSync(settingsPath, 'utf8')) as Record<string, unknown>
    expect(onDisk[SYSTEM_PRESENCE_KEY]).toEqual({
      loginItem: true,
      closeToTray: false,
      hotkeyEnabled: true,
      hotkeyAccelerator: 'Control+Shift+K',
    })
    presence.dispose()
  })

  it('close-to-tray defaults per platform and normalizes like every other pref', () => {
    expect(DEFAULT_SYSTEM_PRESENCE_PREFS.closeToTray).toBe(process.platform !== 'darwin')
    writeFileSync(settingsPath, JSON.stringify({ [SYSTEM_PRESENCE_KEY]: { closeToTray: false } }))
    const { presence } = makePresence()
    // a stored false survives; the other prefs fall back to their defaults
    expect(presence.settings().closeToTray).toBe(false)
    // unknown shapes fall back to the default rather than guessing
    expect(normalizeSystemPresencePrefs({ closeToTray: 'yes' }).closeToTray).toBe(
      DEFAULT_SYSTEM_PRESENCE_PREFS.closeToTray,
    )
    presence.dispose()
  })

  it('the close-to-tray IPC validates, persists and answers the effective snapshot', async () => {
    const { presence } = makePresence()
    registerSystemPresenceIpc(() => presence)
    await expect(invoke(HOME_CHANNELS.systemPresenceSetCloseToTray, 'yes')).rejects.toThrow(
      /must be a boolean/,
    )
    const after = (await invoke(HOME_CHANNELS.systemPresenceSetCloseToTray, false)) as ReturnType<
      typeof presence.settings
    >
    expect(after.closeToTray).toBe(false)
    expect(after.closeToTraySupported).toBe(process.platform !== 'darwin')
    const onDisk = JSON.parse(readFileSync(settingsPath, 'utf8')) as Record<string, unknown>
    expect((onDisk[SYSTEM_PRESENCE_KEY] as { closeToTray: boolean }).closeToTray).toBe(false)
    presence.dispose()
  })

  it('toggling the login item or the hotkey carries close-to-tray through untouched', () => {
    const { presence } = makePresence()
    registerSystemPresenceIpc(() => presence)
    const before = presence.settings()
    invoke(HOME_CHANNELS.systemPresenceSetLoginItem, !before.loginItem)
    invoke(HOME_CHANNELS.systemPresenceSetHotkey, { accelerator: 'Alt+P' })
    // neither handler rebuilds the prefs from partials and drops the field
    expect(presence.settings().closeToTray).toBe(before.closeToTray)
    presence.dispose()
  })

  it('dispose unregisters every global shortcut and destroys the tray', () => {
    const { presence } = makePresence()
    presence.dispose()
    expect(state.registered.size).toBe(0)
    expect(state.trays[0]!.destroyed).toBe(true)
  })

  it('retryHotkey re-registers a refused accelerator without touching a working one', () => {
    // the accelerator is taken at install time: the snapshot says so and
    // nothing is registered
    state.refuseRegistrations = true
    const { presence } = makePresence()
    expect(presence.settings().hotkeyActive).toBe(false)
    expect(state.registered.size).toBe(0)

    // the holding app quit: a focus-time retry now succeeds
    state.refuseRegistrations = false
    const retried = presence.retryHotkey()
    expect(retried.hotkeyActive).toBe(true)
    expect([...state.registered.keys()]).toEqual(['CommandOrControl+Shift+Space'])

    // an already-active hotkey is left alone: no unregister/re-register churn
    const unregistersBefore = state.unregisterCalls
    expect(presence.retryHotkey().hotkeyActive).toBe(true)
    expect(state.unregisterCalls).toBe(unregistersBefore)
    presence.dispose()
  })
})

describe('registerSystemPresenceIpc', () => {
  it('reads, toggles and validates through the channels', async () => {
    const { presence } = makePresence()
    registerSystemPresenceIpc(() => presence)

    const initial = (await invoke(HOME_CHANNELS.systemPresenceGet)) as {
      hotkey: { enabled: boolean; accelerator: string }
    }
    expect(initial.hotkey).toEqual({
      enabled: true,
      accelerator: 'CommandOrControl+Shift+Space',
    })

    const afterLogin = (await invoke(HOME_CHANNELS.systemPresenceSetLoginItem, true)) as {
      loginItem: boolean
    }
    expect(afterLogin.loginItem).toBe(true)

    const afterHotkey = (await invoke(HOME_CHANNELS.systemPresenceSetHotkey, {
      accelerator: 'CommandOrControl+Shift+Space',
    })) as { hotkey: { enabled: boolean; accelerator: string } }
    expect(afterHotkey.hotkey.accelerator).toBe('CommandOrControl+Shift+Space')
    expect([...state.registered.keys()]).toEqual(['CommandOrControl+Shift+Space'])

    await expect(invoke(HOME_CHANNELS.systemPresenceSetLoginItem, 'yes')).rejects.toThrow(/boolean/)
    await expect(
      invoke(HOME_CHANNELS.systemPresenceSetHotkey, { accelerator: 'not a hotkey' }),
    ).rejects.toThrow(/accelerator/)
    await expect(invoke(HOME_CHANNELS.systemPresenceSetHotkey, 42)).rejects.toThrow(
      /enabled.*accelerator/,
    )
    // the refused writes left the persisted prefs untouched
    const onDisk = JSON.parse(readFileSync(settingsPath, 'utf8')) as Record<
      string,
      { hotkeyAccelerator: string }
    >
    expect(onDisk[SYSTEM_PRESENCE_KEY]!.hotkeyAccelerator).toBe('CommandOrControl+Shift+Space')
    presence.dispose()
  })

  it('reports null when the presence is not installed (headless runs)', async () => {
    registerSystemPresenceIpc(() => null)
    await expect(invoke(HOME_CHANNELS.systemPresenceGet)).resolves.toBeNull()
  })
})

describe('deep-link intake', () => {
  it('delivers open-url payloads and ignores the rest', () => {
    const delivered: unknown[] = []
    installDeepLinkListener((payload) => delivered.push(payload))
    const fire = (url: string) => {
      const event = { preventDefault: vi.fn() }
      for (const listener of state.appListeners.get('open-url') ?? []) {
        listener(event as never, url as never)
      }
      return event
    }
    const hit = fire('genoffice://ask?q=hello')
    expect(hit.preventDefault).toHaveBeenCalled()
    expect(delivered).toEqual([{ view: 'ask', query: 'hello' }])
    fire('genoffice://nope')
    expect(delivered).toHaveLength(1)
  })

  it('claims the genoffice scheme best-effort', () => {
    registerDeepLinkClient()
    expect(state.protocolClients).toEqual(['genoffice'])
  })
})
