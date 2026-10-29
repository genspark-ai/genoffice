import { app, dialog, ipcMain, type BrowserWindow } from 'electron'
import {
  installPluginDir,
  installPluginZip,
  listPlugins,
  removePlugin,
  setPluginEnabled,
} from './imported-plugins'
import type { InstalledPlugin } from '../shared/found-skill'

export interface PluginsDeps {
  /** the shell window dialogs attach to */
  window: () => BrowserWindow | null
}

/**
 * Settings › Plugins: list, install, enable, remove.
 *
 * The import pickers open the native dialogs here in the main process — the
 * integrations pane's pattern — so no path ever crosses the IPC boundary: the
 * renderer asks for "import a plugin" and gets back the plugin that landed, or
 * null when the dialog was cancelled. Everything else the renderer can send is
 * a plugin name and a flag, and the underlying calls accept only names the
 * listing itself vouched for.
 */
export function registerPluginsIpc(deps: PluginsDeps): void {
  ipcMain.handle('ai:plugins-list', async (): Promise<InstalledPlugin[]> =>
    listPlugins(app.getPath('userData')),
  )

  const pick = async (opts: Electron.OpenDialogOptions): Promise<string | null> => {
    const win = deps.window()
    const r = await (win ? dialog.showOpenDialog(win, opts) : dialog.showOpenDialog(opts))
    return r.canceled ? null : (r.filePaths[0] ?? null)
  }

  // the only writes: a person pressed Install and picked something in a dialog
  // this process owns, so the source path never comes from the renderer
  ipcMain.handle('ai:plugin-import-folder', async (): Promise<InstalledPlugin | null> => {
    const dir = await pick({ properties: ['openDirectory'] })
    return dir ? installPluginDir(dir, app.getPath('userData')) : null
  })

  ipcMain.handle('ai:plugin-import-zip', async (): Promise<InstalledPlugin | null> => {
    const zip = await pick({
      properties: ['openFile'],
      filters: [{ name: 'Plugin package', extensions: ['zip'] }],
    })
    return zip ? installPluginZip(zip, app.getPath('userData')) : null
  })

  ipcMain.handle(
    'ai:plugin-set-enabled',
    async (_e, name: unknown, enabled: unknown): Promise<void> => {
      if (typeof name !== 'string' || name.trim() === '') throw new Error('unknown plugin')
      if (typeof enabled !== 'boolean') throw new Error('enabled must be a true or false flag')
      await setPluginEnabled(app.getPath('userData'), name, enabled)
    },
  )

  ipcMain.handle('ai:plugin-remove', async (_e, name: unknown): Promise<void> => {
    if (typeof name !== 'string' || name.trim() === '') throw new Error('unknown plugin')
    await removePlugin(app.getPath('userData'), name)
  })
}
