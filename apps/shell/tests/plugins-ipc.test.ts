import { describe, expect, it, vi, beforeEach } from 'vitest'

/**
 * The five plugin channels: the pickers open the native dialogs in the main
 * process (no path crosses IPC), and everything the renderer sends is a name
 * the listing vouched for. The store itself is imported-plugins.test.ts's
 * subject; here it is mocked so only the wiring is under test.
 */

const handles = new Map<string, (...args: never[]) => unknown>()
const openDialog = vi.fn()
vi.mock('electron', () => ({
  app: { getPath: () => '/app/userData' },
  ipcMain: {
    handle: (channel: string, handler: (...args: never[]) => unknown) =>
      handles.set(channel, handler),
  },
  dialog: { showOpenDialog: (...a: unknown[]) => openDialog(...a) },
}))

const list = vi.fn()
const installDir = vi.fn()
const installZip = vi.fn()
const setEnabled = vi.fn()
const remove = vi.fn()
vi.mock('../src/main/imported-plugins', () => ({
  listPlugins: (...a: unknown[]) => list(...a),
  installPluginDir: (...a: unknown[]) => installDir(...a),
  installPluginZip: (...a: unknown[]) => installZip(...a),
  setPluginEnabled: (...a: unknown[]) => setEnabled(...a),
  removePlugin: (...a: unknown[]) => remove(...a),
}))

import { registerPluginsIpc } from '../src/main/plugins-ipc'

const PLUGIN = {
  name: 'acme',
  version: '1.0.0',
  description: 'The acme plugin',
  author: 'Acme',
  enabled: true,
  components: { skills: ['acme:helper'], commands: [], mcpServers: [] },
  format: 'plain',
} as const

const invoke = (channel: string, ...args: unknown[]) =>
  handles.get(channel)!({ sender: { id: 7 } } as never, ...args)

beforeEach(() => {
  handles.clear()
  openDialog.mockReset()
  list.mockReset().mockResolvedValue([PLUGIN])
  installDir.mockReset().mockResolvedValue(PLUGIN)
  installZip.mockReset().mockResolvedValue(PLUGIN)
  setEnabled.mockReset().mockResolvedValue(undefined)
  remove.mockReset().mockResolvedValue(undefined)
  registerPluginsIpc({ window: () => null })
})

describe('ai:plugins-list', () => {
  it('lists the installed plugins from the app userData', async () => {
    await expect(invoke('ai:plugins-list')).resolves.toEqual([PLUGIN])
    expect(list).toHaveBeenCalledWith('/app/userData')
  })
})

describe('the import pickers', () => {
  it('picks a folder with a native dialog and installs it', async () => {
    openDialog.mockResolvedValue({ canceled: false, filePaths: ['/picked/acme'] })
    await expect(invoke('ai:plugin-import-folder')).resolves.toBe(PLUGIN)
    expect(openDialog).toHaveBeenCalledWith({ properties: ['openDirectory'] })
    expect(installDir).toHaveBeenCalledWith('/picked/acme', '/app/userData')
  })

  it('picks a zip file, with zip filters, and installs it', async () => {
    openDialog.mockResolvedValue({ canceled: false, filePaths: ['/picked/acme.zip'] })
    await expect(invoke('ai:plugin-import-zip')).resolves.toBe(PLUGIN)
    expect(openDialog).toHaveBeenCalledWith({
      properties: ['openFile'],
      filters: [{ name: 'Plugin package', extensions: ['zip'] }],
    })
    expect(installZip).toHaveBeenCalledWith('/picked/acme.zip', '/app/userData')
  })

  it('returns null and installs nothing when the dialog is cancelled', async () => {
    openDialog.mockResolvedValue({ canceled: true, filePaths: [] })
    await expect(invoke('ai:plugin-import-folder')).resolves.toBe(null)
    await expect(invoke('ai:plugin-import-zip')).resolves.toBe(null)
    expect(installDir).not.toHaveBeenCalled()
    expect(installZip).not.toHaveBeenCalled()
  })

  it('ignores renderer arguments entirely: only the dialog path reaches the store', async () => {
    openDialog.mockResolvedValue({ canceled: false, filePaths: ['/picked/acme'] })
    await invoke('ai:plugin-import-folder', '/etc/passwd', { path: true })
    expect(installDir).toHaveBeenCalledTimes(1)
    expect(installDir).toHaveBeenCalledWith('/picked/acme', '/app/userData')
  })

  it('attaches the dialog to the shell window when there is one', async () => {
    const win = { id: 1 } as never
    registerPluginsIpc({ window: () => win })
    openDialog.mockResolvedValue({ canceled: true })
    await invoke('ai:plugin-import-folder')
    expect(openDialog).toHaveBeenCalledWith(win, { properties: ['openDirectory'] })
  })
})

describe('ai:plugin-set-enabled / ai:plugin-remove', () => {
  it('forwards a name and flag to the store, keyed on the app userData', async () => {
    await invoke('ai:plugin-set-enabled', 'acme', false)
    expect(setEnabled).toHaveBeenCalledWith('/app/userData', 'acme', false)
    await invoke('ai:plugin-remove', 'acme')
    expect(remove).toHaveBeenCalledWith('/app/userData', 'acme')
  })

  it('refuses anything that is not a plugin name and a boolean', async () => {
    await expect(invoke('ai:plugin-set-enabled', 'acme', 'yes')).rejects.toThrow(
      'true or false flag',
    )
    await expect(invoke('ai:plugin-set-enabled', 42, false)).rejects.toThrow('unknown plugin')
    await expect(invoke('ai:plugin-set-enabled', '   ', true)).rejects.toThrow('unknown plugin')
    await expect(invoke('ai:plugin-remove', { name: 'acme' })).rejects.toThrow('unknown plugin')
    await expect(invoke('ai:plugin-remove', undefined)).rejects.toThrow('unknown plugin')
    expect(setEnabled).not.toHaveBeenCalled()
    expect(remove).not.toHaveBeenCalled()
  })
})
