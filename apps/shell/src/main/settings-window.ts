/**
 * The standalone settings window: sidebar navigation over Model, Network,
 * Rules, Skills and General.
 *
 * A real window rather than a modal card, because Rules and Skills are
 * editing surfaces the user keeps open while working, and because it absorbs
 * the language / update-channel / account items that used to live in the
 * cramped account popup.
 *
 * Content lives in renderer/settings.html; this module owns the lifecycle and
 * the IPC surface. It deliberately reuses the app-wide `ai:*` handlers for
 * anything the editors also need, and only adds `settings:*` channels for
 * things that are specific to this window.
 */
import { readFile } from 'node:fs/promises'
import { basename, join } from 'node:path'
import { BrowserWindow, dialog, ipcMain, shell } from 'electron'
import type { AgentRules, UserSkill } from '@genoffice/agent-core'
import type { AiModelSettings, SettingsSnapshot } from '../shared/settings-api'
import { SETTINGS_CHANNELS } from '../shared/settings-api'

/** the host supplies these so this module stays out of the shell's own state */
export interface SettingsWindowHost {
  loadSnapshot: () => Promise<SettingsSnapshot>
  saveAi: (settings: AiModelSettings) => AiModelSettings
  saveRules: (rules: AgentRules) => AgentRules
  saveSkill: (input: Partial<UserSkill> & { name: string; body: string }) => UserSkill
  deleteSkill: (id: string) => void
  importSkillContents: (files: Array<{ filename: string; content: string }>) => UserSkill[]
  setLanguage: (lang: string) => void
  setUpdateChannel: (channel: 'stable' | 'beta') => void
  accountLogin: () => boolean
  accountLogout: () => Promise<void>
}

let settingsWin: BrowserWindow | null = null
let host: SettingsWindowHost | null = null
let ipcRegistered = false

/** one .md may be large but is still text; refuse anything that clearly is not a skill */
const MAX_SKILL_FILE_BYTES = 512 * 1024

function registerIpc(): void {
  if (ipcRegistered) return
  ipcRegistered = true

  ipcMain.handle(SETTINGS_CHANNELS.load, async () => host?.loadSnapshot())
  ipcMain.handle(SETTINGS_CHANNELS.saveAi, (_e, settings: AiModelSettings) =>
    host?.saveAi(settings),
  )
  ipcMain.handle(SETTINGS_CHANNELS.saveRules, (_e, rules: AgentRules) => host?.saveRules(rules))
  ipcMain.handle(SETTINGS_CHANNELS.saveSkill, (_e, skill) => host?.saveSkill(skill))
  ipcMain.handle(SETTINGS_CHANNELS.deleteSkill, (_e, id: string) => host?.deleteSkill(String(id)))
  ipcMain.handle(SETTINGS_CHANNELS.importSkillContents, (_e, files) =>
    host?.importSkillContents(Array.isArray(files) ? files : []),
  )

  /** native picker: main reads the files, so the renderer never touches fs */
  ipcMain.handle(SETTINGS_CHANNELS.importSkillFiles, async (): Promise<UserSkill[]> => {
    const parent = settingsWin && !settingsWin.isDestroyed() ? settingsWin : undefined
    const result = await dialog.showOpenDialog(parent ?? (undefined as unknown as BrowserWindow), {
      title: 'Import skills',
      filters: [{ name: 'Markdown', extensions: ['md', 'markdown'] }],
      properties: ['openFile', 'multiSelections'],
    })
    if (result.canceled || !result.filePaths.length) return []
    const files: Array<{ filename: string; content: string }> = []
    for (const path of result.filePaths.slice(0, 50)) {
      try {
        const content = await readFile(path, 'utf-8')
        if (content.length > MAX_SKILL_FILE_BYTES) continue
        files.push({ filename: basename(path), content })
      } catch {
        /* unreadable file: skip it rather than failing the whole import */
      }
    }
    return host?.importSkillContents(files) ?? []
  })

  ipcMain.handle(SETTINGS_CHANNELS.setLanguage, (_e, lang: string) =>
    host?.setLanguage(String(lang)),
  )
  ipcMain.handle(SETTINGS_CHANNELS.setUpdateChannel, (_e, channel) =>
    host?.setUpdateChannel(channel === 'beta' ? 'beta' : 'stable'),
  )
  ipcMain.handle(SETTINGS_CHANNELS.accountLogin, () => host?.accountLogin() ?? false)
  ipcMain.handle(SETTINGS_CHANNELS.accountLogout, async () => host?.accountLogout())
  ipcMain.handle(SETTINGS_CHANNELS.close, () => {
    if (settingsWin && !settingsWin.isDestroyed()) settingsWin.close()
  })
}

/** push a fresh snapshot, e.g. once a browser sign-in completes */
export function refreshSettingsWindow(): void {
  if (settingsWin && !settingsWin.isDestroyed()) {
    settingsWin.webContents.send(SETTINGS_CHANNELS.changed)
  }
}

export function showSettingsWindow(
  parent: BrowserWindow | null,
  windowHost: SettingsWindowHost,
): void {
  host = windowHost
  registerIpc()

  if (settingsWin && !settingsWin.isDestroyed()) {
    settingsWin.show()
    settingsWin.focus()
    return
  }

  const win = new BrowserWindow({
    width: 880,
    height: 640,
    minWidth: 720,
    minHeight: 520,
    ...(parent && !parent.isDestroyed() ? { parent } : {}),
    show: false,
    title: 'Settings',
    autoHideMenuBar: true,
    webPreferences: {
      preload: join(__dirname, '../preload/settings.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  })
  settingsWin = win

  // links in help text open in the real browser, never in this window
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https?:\/\//i.test(url)) void shell.openExternal(url)
    return { action: 'deny' }
  })

  const devUrl = process.env.SHELL_RENDERER_URL
  void (devUrl
    ? win.loadURL(new URL('settings.html', devUrl).toString())
    : win.loadFile(join(__dirname, '../renderer/settings.html')))

  win.once('ready-to-show', () => win.show())
  win.on('closed', () => {
    settingsWin = null
  })
}
