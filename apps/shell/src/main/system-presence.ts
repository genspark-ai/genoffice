import { join } from 'node:path'
import { app, globalShortcut, ipcMain, Menu, Tray } from 'electron'
import type { NativeImage } from 'electron'
import { createI18n, onUiLangChange, type Lang } from '@genoffice/i18n'
import { readAppSettings, writeAppSetting } from './app-settings'
import {
  HOME_CHANNELS,
  type DeepLinkPayload,
  type SystemPresenceSettings,
} from '../shared/home-api'

/**
 * The workbench's system presence: a tray icon with the quick actions, a
 * global hotkey summoning the command palette, the launch-at-login toggle and
 * the genoffice:// deep-link intake. Everything user-facing is persisted in
 * app-settings.json under `systemPresence` (the settings-center store) and
 * reconciled live: a settings change re-registers the hotkey and re-applies
 * the login item without a restart.
 *
 * Electron objects (Tray, globalShortcut) come from the module like the other
 * main-process IPC modules; the volatile behavior — what "show", "palette"
 * and "ask" mean for the shell window — is injected, and tests mock
 * 'electron' (the desktop-ipc.test.ts style).
 */

/** the app-settings.json key holding these prefs */
export const SYSTEM_PRESENCE_KEY = 'systemPresence'

export interface SystemPresencePrefs {
  /** launch at login (macOS/Windows; Linux has no single API — see applyLoginItem) */
  loginItem: boolean
  hotkeyEnabled: boolean
  /** Electron accelerator, e.g. 'CommandOrControl+Shift+Space' */
  hotkeyAccelerator: string
}

export const DEFAULT_SYSTEM_PRESENCE_PREFS: SystemPresencePrefs = {
  loginItem: false,
  hotkeyEnabled: true,
  // the system window-menu accelerator on Windows and the non-breaking space
  // in macOS text fields (also Raycast's default) — a modifier-heavy combo
  // nobody else claims is the safer default
  hotkeyAccelerator: 'CommandOrControl+Shift+Space',
}

/** a stored accelerator can be longish, never arbitrary text */
const MAX_ACCELERATOR = 64

const ACCELERATOR_MODIFIERS = new Set([
  'command',
  'cmd',
  'control',
  'ctrl',
  'alt',
  'option',
  'shift',
  'super',
  'meta',
  'commandorcontrol',
  'cmdorctrl',
])

/** the non-modifier tail of an accelerator: one key name or printable char */
const ACCELERATOR_KEY_RE =
  /^([a-z0-9]|f([1-9]|1[0-9]|2[0-4])|space|tab|enter|return|escape|esc|backspace|delete|insert|home|end|pageup|pagedown|up|down|left|right|plus|minus|printscreen|num[0-9]|numdec|numadd|numsub|nummult|numdiv|[`\-=[\]\\;',./~!@#$%^&*()_+{}|:"<>?])$/i

/**
 * A pragmatic accelerator shape check: modifiers from the known set, then one
 * key. Electron's own grammar is broader (this accepts a few strings register()
 * will refuse); the registration result is what the settings snapshot reports,
 * so an accepted-but-unregistrable value never pretends to be active.
 */
export function isAcceleratorLike(value: unknown): value is string {
  if (typeof value !== 'string') return false
  const trimmed = value.trim()
  if (trimmed === '' || trimmed.length > MAX_ACCELERATOR) return false
  const parts = trimmed.split('+')
  if (parts.length < 2) {
    // a bare modifier-less accelerator (F5, Space) is legal but a poor global
    // hotkey — it would steal the key from every app; require a modifier
    return false
  }
  const key = parts[parts.length - 1]!
  if (!ACCELERATOR_KEY_RE.test(key)) return false
  return parts.slice(0, -1).every((m) => ACCELERATOR_MODIFIERS.has(m.toLowerCase()))
}

/** the store's invariant, applied to anything read: unknown fields fall back individually */
export function normalizeSystemPresencePrefs(raw: unknown): SystemPresencePrefs {
  const prefs = { ...DEFAULT_SYSTEM_PRESENCE_PREFS }
  if (raw === null || typeof raw !== 'object' || Array.isArray(raw)) return prefs
  const source = raw as Record<string, unknown>
  if (typeof source.loginItem === 'boolean') prefs.loginItem = source.loginItem
  if (typeof source.hotkeyEnabled === 'boolean') prefs.hotkeyEnabled = source.hotkeyEnabled
  if (isAcceleratorLike(source.hotkeyAccelerator)) {
    prefs.hotkeyAccelerator = source.hotkeyAccelerator.trim()
  }
  return prefs
}

export function readSystemPresencePrefs(settingsPath: string): SystemPresencePrefs {
  return normalizeSystemPresencePrefs(readAppSettings(settingsPath)[SYSTEM_PRESENCE_KEY])
}

// ---- tray menu strings: zh defines the key set, every sibling must match ----

const tPresence = createI18n({
  zh: {
    trayShow: '显示 GenOffice',
    trayPalette: '命令面板',
    trayAsk: 'AI 问答',
    trayQuit: '退出 GenOffice',
  },
  en: {
    trayShow: 'Show GenOffice',
    trayPalette: 'Command Palette',
    trayAsk: 'Ask AI',
    trayQuit: 'Quit GenOffice',
  },
  ja: {
    trayShow: 'GenOffice を表示',
    trayPalette: 'コマンドパレット',
    trayAsk: 'AI に聞く',
    trayQuit: 'GenOffice を終了',
  },
  ko: {
    trayShow: 'GenOffice 표시',
    trayPalette: '명령 팔레트',
    trayAsk: 'AI에게 묻기',
    trayQuit: 'GenOffice 종료',
  },
  fr: {
    trayShow: 'Afficher GenOffice',
    trayPalette: 'Palette de commandes',
    trayAsk: 'Demander à l’IA',
    trayQuit: 'Quitter GenOffice',
  },
  de: {
    trayShow: 'GenOffice anzeigen',
    trayPalette: 'Befehlspalette',
    trayAsk: 'KI fragen',
    trayQuit: 'GenOffice beenden',
  },
  es: {
    trayShow: 'Mostrar GenOffice',
    trayPalette: 'Paleta de comandos',
    trayAsk: 'Preguntar a la IA',
    trayQuit: 'Salir de GenOffice',
  },
  th: {
    trayShow: 'แสดง GenOffice',
    trayPalette: 'พาเลตคำสั่ง',
    trayAsk: 'ถาม AI',
    trayQuit: 'ออกจาก GenOffice',
  },
  id: {
    trayShow: 'Tampilkan GenOffice',
    trayPalette: 'Palet Perintah',
    trayAsk: 'Tanya AI',
    trayQuit: 'Keluar dari GenOffice',
  },
  ru: {
    trayShow: 'Показать GenOffice',
    trayPalette: 'Палитра команд',
    trayAsk: 'Спросить ИИ',
    trayQuit: 'Выйти из GenOffice',
  },
  ar: {
    trayShow: 'إظهار GenOffice',
    trayPalette: 'لوحة الأوامر',
    trayAsk: 'اسأل الذكاء الاصطناعي',
    trayQuit: 'إنهاء GenOffice',
  },
  pt: {
    trayShow: 'Mostrar GenOffice',
    trayPalette: 'Paleta de comandos',
    trayAsk: 'Perguntar à IA',
    trayQuit: 'Sair do GenOffice',
  },
  it: {
    trayShow: 'Mostra GenOffice',
    trayPalette: 'Tavolozza dei comandi',
    trayAsk: 'Chiedi all’IA',
    trayQuit: 'Esci da GenOffice',
  },
  pl: {
    trayShow: 'Pokaż GenOffice',
    trayPalette: 'Paleta poleceń',
    trayAsk: 'Zapytaj AI',
    trayQuit: 'Zakończ GenOffice',
  },
  cs: {
    trayShow: 'Zobrazit GenOffice',
    trayPalette: 'Paleta příkazů',
    trayAsk: 'Zeptat se AI',
    trayQuit: 'Ukončit GenOffice',
  },
  nl: {
    trayShow: 'GenOffice weergeven',
    trayPalette: 'Opdrachtenpalet',
    trayAsk: 'Vraag AI',
    trayQuit: 'GenOffice afsluiten',
  },
  ms: {
    trayShow: 'Tunjukkan GenOffice',
    trayPalette: 'Palet Perintah',
    trayAsk: 'Tanya AI',
    trayQuit: 'Keluar GenOffice',
  },
  he: {
    trayShow: 'הצגת GenOffice',
    trayPalette: 'לוח הפקודות',
    trayAsk: 'שאל את ה-AI',
    trayQuit: 'יציאה מ-GenOffice',
  },
  hi: {
    trayShow: 'GenOffice दिखाएँ',
    trayPalette: 'कमांड पैलेट',
    trayAsk: 'AI से पूछें',
    trayQuit: 'GenOffice से बाहर निकलें',
  },
  'zh-TW': {
    trayShow: '顯示 GenOffice',
    trayPalette: '命令選板',
    trayAsk: 'AI 問答',
    trayQuit: '結束 GenOffice',
  },
  vi: {
    trayShow: 'Hiện GenOffice',
    trayPalette: 'Bảng lệnh',
    trayAsk: 'Hỏi AI',
    trayQuit: 'Thoát GenOffice',
  },
})

// ---- genoffice:// deep links ----

export const DEEP_LINK_SCHEME = 'genoffice'

/**
 * Parse one deep-link URL. Three verbs exist:
 * genoffice://ask?q=hello → { view: 'ask', query: 'hello' },
 * genoffice://palette → { view: 'palette' },
 * genoffice://open?path=%2F… → { view: 'open', path: '…' }. Anything else is
 * not a deep link (null) rather than an error.
 */
export function parseDeepLink(raw: string): DeepLinkPayload | null {
  let url: URL
  try {
    url = new URL(raw)
  } catch {
    return null
  }
  if (url.protocol !== `${DEEP_LINK_SCHEME}:`) return null
  const host = url.hostname.toLowerCase()
  if (host === 'ask') {
    const query = url.searchParams.get('q')
    return { view: 'ask', ...(query ? { query } : {}) }
  }
  if (host === 'palette') return { view: 'palette' }
  if (host === 'open') {
    // a link without a path has nothing to open; refusing it beats guessing
    const path = url.searchParams.get('path')
    return path ? { view: 'open', path } : null
  }
  return null
}

/** every genoffice:// URL in a process argv, in order (Windows/Linux protocol
 *  launches; several can queue up in one argv) */
export function deepLinksFromArgv(argv: readonly string[]): string[] {
  return argv.filter((arg) => typeof arg === 'string' && arg.startsWith(`${DEEP_LINK_SCHEME}://`))
}

/**
 * Claim the genoffice:// scheme. Packaged builds carry the OS registration
 * (electron-builder's protocols); the call matters for dev and for Windows,
 * where the OS binding is written at runtime. Best-effort: a platform that
 * refuses it (unsigned dev builds) must not take the app down.
 */
export function registerDeepLinkClient(): void {
  try {
    app.setAsDefaultProtocolClient(DEEP_LINK_SCHEME)
  } catch {
    // best-effort: protocol registration is a convenience, not a launch blocker
  }
}

/**
 * The macOS half of deep-link intake: `open-url` (can fire before ready, so it
 * must be installed early — the delivery callback decides what "not ready"
 * means). Windows/Linux links arrive in argv instead; the caller checks
 * deepLinksFromArgv in its second-instance and cold-start paths.
 */
export function installDeepLinkListener(deliver: (payload: DeepLinkPayload) => void): void {
  app.on('open-url', (event, url) => {
    event.preventDefault()
    const payload = parseDeepLink(url)
    if (payload) deliver(payload)
  })
}

// ---- the presence itself ----

export interface SystemPresenceDeps {
  /** app-settings.json path; default <userData>/app-settings.json */
  settingsPath?: () => string
  lang: () => Lang
  /** focus the main window on its Home tab */
  showMainWindow: () => void
  /** focus Home and open the command palette (the Ctrl/Cmd+K path) */
  summonPalette: () => void
  /** focus the main window on Home's Ask AI view, optionally with a query */
  openAsk: (query?: string) => void
  /** default app.quit() */
  quit?: () => void
  /** the tray icon (template-rendered on macOS); provided by the caller because
   *  ?asset imports only resolve in the bundled build, not under vitest */
  trayIcon: () => NativeImage
}

export interface SystemPresence {
  /** re-read the persisted prefs and reconcile tray menu, hotkey and login item */
  refresh(): void
  /** persist + apply a new pref set; answers the live settings snapshot */
  applyPrefs(next: SystemPresencePrefs): SystemPresenceSettings
  /** the live settings snapshot the settings UI renders */
  settings(): SystemPresenceSettings
  /**
   * Re-register the hotkey from the in-memory prefs (no disk read): the app
   * that held the accelerator when registration failed may since have quit, so
   * a window gaining focus is a cheap moment to try again. An active hotkey is
   * left untouched — churn on every focus risks losing a working registration.
   */
  retryHotkey(): SystemPresenceSettings
  /** drop the tray icon and unregister every global shortcut (will-quit) */
  dispose(): void
}

export function installSystemPresence(deps: SystemPresenceDeps): SystemPresence {
  const settingsPath =
    deps.settingsPath ?? (() => join(app.getPath('userData'), 'app-settings.json'))
  const quit = deps.quit ?? (() => app.quit())
  const tray = new Tray(deps.trayIcon())
  tray.setToolTip('GenOffice')

  let registeredHotkey: string | null = null

  const rebuildMenu = (): void => {
    const lang = deps.lang()
    const t = (key: 'trayShow' | 'trayPalette' | 'trayAsk' | 'trayQuit') => tPresence(lang, key)
    tray.setContextMenu(
      Menu.buildFromTemplate([
        { label: t('trayShow'), click: () => deps.showMainWindow() },
        { label: t('trayPalette'), click: () => deps.summonPalette() },
        { label: t('trayAsk'), click: () => deps.openAsk() },
        { type: 'separator' },
        { label: t('trayQuit'), click: () => quit() },
      ]),
    )
  }

  /** re-register to match the prefs; answers whether the OS accepted it */
  const applyHotkey = (prefs: SystemPresencePrefs): boolean => {
    if (registeredHotkey !== null) {
      globalShortcut.unregister(registeredHotkey)
      registeredHotkey = null
    }
    if (!prefs.hotkeyEnabled) return false
    try {
      const ok = globalShortcut.register(prefs.hotkeyAccelerator, () => deps.summonPalette())
      if (ok) registeredHotkey = prefs.hotkeyAccelerator
      return ok
    } catch {
      return false
    }
  }

  /**
   * Launch at login: one API on macOS and Windows. Linux has no
   * app.setLoginItemSettings equivalent (a .desktop file in
   * ~/.config/autostart would be the way); the pref persists but nothing is
   * applied, and the settings snapshot says the platform cannot drive it.
   */
  const loginItemSupported = process.platform === 'darwin' || process.platform === 'win32'
  const applyLoginItem = (prefs: SystemPresencePrefs): void => {
    if (!loginItemSupported) return
    try {
      app.setLoginItemSettings({ openAtLogin: prefs.loginItem })
    } catch {
      // best-effort: a platform/session that refuses keeps the pref for next time
    }
  }

  let prefs = readSystemPresencePrefs(settingsPath())
  let hotkeyActive = applyHotkey(prefs)
  applyLoginItem(prefs)
  rebuildMenu()
  // the tray menu is native: it must follow a language switch like the app menu does
  const offLangChange = onUiLangChange(() => rebuildMenu())

  const snapshot = (): SystemPresenceSettings => ({
    loginItem: prefs.loginItem,
    hotkey: { enabled: prefs.hotkeyEnabled, accelerator: prefs.hotkeyAccelerator },
    hotkeyActive,
    loginItemSupported,
  })

  return {
    refresh() {
      prefs = readSystemPresencePrefs(settingsPath())
      hotkeyActive = applyHotkey(prefs)
      applyLoginItem(prefs)
    },
    applyPrefs(next: SystemPresencePrefs): SystemPresenceSettings {
      writeAppSetting(settingsPath(), SYSTEM_PRESENCE_KEY, next)
      prefs = next
      hotkeyActive = applyHotkey(prefs)
      applyLoginItem(prefs)
      return snapshot()
    },
    settings: snapshot,
    retryHotkey(): SystemPresenceSettings {
      if (hotkeyActive) return snapshot()
      hotkeyActive = applyHotkey(prefs)
      return snapshot()
    },
    dispose() {
      offLangChange()
      // quit-time hygiene: nothing outlives the app (unregisterAll, not just ours)
      globalShortcut.unregisterAll()
      tray.destroy()
    },
  }
}

/**
 * The settings center's IPC: read the prefs, toggle the login item, toggle or
 * retune the global hotkey. Validation follows the desktop-ipc defensive
 * pattern: a payload that does not carry the promised shape is refused with a
 * plain Error (the file stays as it was); an accelerator the OS will not
 * register is reported through `hotkeyActive` in the returned snapshot.
 */
export function registerSystemPresenceIpc(presence: () => SystemPresence | null): void {
  ipcMain.handle(
    HOME_CHANNELS.systemPresenceGet,
    async (): Promise<SystemPresenceSettings | null> => {
      return presence()?.settings() ?? null
    },
  )

  ipcMain.handle(
    HOME_CHANNELS.systemPresenceSetLoginItem,
    async (_event, raw: unknown): Promise<SystemPresenceSettings> => {
      if (typeof raw !== 'boolean') {
        throw new Error('the login-item setting must be a boolean')
      }
      const p = presence()
      if (!p) throw new Error('system presence is not available')
      const current = p.settings()
      return p.applyPrefs({
        loginItem: raw,
        hotkeyEnabled: current.hotkey.enabled,
        hotkeyAccelerator: current.hotkey.accelerator,
      })
    },
  )

  ipcMain.handle(
    HOME_CHANNELS.systemPresenceSetHotkey,
    async (_event, raw: unknown): Promise<SystemPresenceSettings> => {
      if (raw === null || typeof raw !== 'object' || Array.isArray(raw)) {
        throw new Error('the hotkey setting must be { enabled?: boolean, accelerator?: string }')
      }
      const patch = raw as Record<string, unknown>
      if (patch.enabled !== undefined && typeof patch.enabled !== 'boolean') {
        throw new Error('the hotkey setting must be { enabled?: boolean, accelerator?: string }')
      }
      if (patch.accelerator !== undefined && !isAcceleratorLike(patch.accelerator)) {
        throw new Error(
          'the hotkey accelerator must be modifiers + one key, e.g. "CommandOrControl+Shift+Space" or "Alt+P"',
        )
      }
      const p = presence()
      if (!p) throw new Error('system presence is not available')
      const current = p.settings()
      return p.applyPrefs({
        loginItem: current.loginItem,
        hotkeyEnabled: (patch.enabled as boolean | undefined) ?? current.hotkey.enabled,
        hotkeyAccelerator:
          typeof patch.accelerator === 'string'
            ? patch.accelerator.trim()
            : current.hotkey.accelerator,
      })
    },
  )
}
