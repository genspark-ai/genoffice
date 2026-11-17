import { BrowserWindow, dialog, type IpcMain } from 'electron'
import { LocalVlmManager, registerLocalVlmRuntime } from '@genoffice/ai-provider/local-vlm'
import {
  LOCAL_VLM_SETTINGS_KEY,
  localVlmSettingsFrom,
  mergeLocalVlmSettings,
  type LocalVlmSettings,
  type LocalVlmStatus,
} from '@genoffice/ai-provider/local-vlm-meta'
import { readAppSettings, writeAppSettingThen } from '../app-settings'
import { LOCAL_VLM_CHANNELS, type LocalVlmPickKind } from '../../shared/fork/local-vlm'

export interface LocalVlmIpcDeps {
  ipcMain: Pick<IpcMain, 'handle'>
  settingsPath: () => string
  userDataPath: () => string
  parentWindow?: () => BrowserWindow | null
}

function broadcastStatus(status: LocalVlmStatus): void {
  for (const win of BrowserWindow.getAllWindows()) {
    if (!win.isDestroyed()) win.webContents.send(LOCAL_VLM_CHANNELS.statusChanged, status)
  }
}

export function registerLocalVlm(deps: LocalVlmIpcDeps): LocalVlmManager {
  let settings: LocalVlmSettings = localVlmSettingsFrom(readAppSettings(deps.settingsPath()))

  const manager = new LocalVlmManager({
    userDataPath: deps.userDataPath(),
    getSettings: () => settings,
    onStatusChange: (next) => broadcastStatus(next),
  })
  registerLocalVlmRuntime(manager)
  // Prime browser-safe usability flag at startup without launching the server process.
  manager.status()

  const { ipcMain } = deps
  ipcMain.handle(LOCAL_VLM_CHANNELS.getStatus, () => manager.status())

  ipcMain.handle(LOCAL_VLM_CHANNELS.setSettings, (_event, patch: unknown) => {
    try {
      const next = mergeLocalVlmSettings(settings, patch)
      writeAppSettingThen(deps.settingsPath(), LOCAL_VLM_SETTINGS_KEY, next, (saved) => {
        settings = saved
        manager.updateSettings(saved)
      })
      return manager.status()
    } catch {
      return null
    }
  })

  ipcMain.handle(
    LOCAL_VLM_CHANNELS.download,
    async (_event, options?: { includeRuntime?: boolean; includeModel?: boolean }) => {
      try {
        const status = await manager.downloadAll(options)
        return { ok: true, status }
      } catch (error) {
        return {
          ok: false,
          status: manager.status(),
          error: error instanceof Error ? error.message : String(error),
        }
      }
    },
  )

  ipcMain.handle(LOCAL_VLM_CHANNELS.cancelDownload, () => manager.cancelDownload())

  ipcMain.handle(LOCAL_VLM_CHANNELS.start, async () => {
    try {
      await manager.ensureRunning()
      return { ok: true, status: manager.status() }
    } catch (error) {
      return {
        ok: false,
        status: manager.status(),
        error: error instanceof Error ? error.message : String(error),
      }
    }
  })

  ipcMain.handle(LOCAL_VLM_CHANNELS.stop, async () => {
    await manager.stop()
    return manager.status()
  })

  ipcMain.handle(LOCAL_VLM_CHANNELS.test, () => manager.testConnection())

  ipcMain.handle(LOCAL_VLM_CHANNELS.pickFile, async (_event, kind: unknown) => {
    const pickKind: LocalVlmPickKind = kind === 'mmproj' || kind === 'runtime' ? kind : 'model'
    const filters =
      pickKind === 'runtime'
        ? process.platform === 'win32'
          ? [{ name: 'Executable', extensions: ['exe'] }]
          : undefined
        : [{ name: 'GGUF Model', extensions: ['gguf'] }]
    const parent = deps.parentWindow?.() ?? undefined
    const opts: Electron.OpenDialogOptions = {
      properties: ['openFile'],
      ...(filters ? { filters } : {}),
    }
    const result = parent
      ? await dialog.showOpenDialog(parent, opts)
      : await dialog.showOpenDialog(opts)
    if (result.canceled || result.filePaths.length === 0) return null
    return result.filePaths[0] ?? null
  })

  return manager
}
