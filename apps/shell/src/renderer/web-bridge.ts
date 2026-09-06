/// Web-version bootstrap for the shell renderer.
///
/// Loaded before main.tsx and only acts outside Electron: it installs the same
/// `window.aiOffice` / `window.aiOfficeProject` / `window.aiOfficeTabs` objects
/// the preload exposes in the desktop app, but backed by the HTTP/SSE transport
/// against the running Electron main process. Native-only channels (browse,
/// save-dir picker, reveal, trash, tab management) get browser equivalents so
/// the web version keeps the full feature surface. Inside Electron the preload
/// has already exposed the IPC-backed APIs and this module leaves them
/// untouched.
import { createHttpIpcTransport, isElectronRuntime } from '@genoffice/ipc-bridge/client'
import { createWebFileBridge, pickFileBytes } from '@genoffice/ipc-bridge/web-native'
import {
  createShellHomeApi,
  createShellProjectApi,
  createShellTabsApi,
} from '../shared/shell-api-factory'

if (!isElectronRuntime()) {
  const transport = createHttpIpcTransport()
  const files = createWebFileBridge(transport)
  const bridgedWindow = window as unknown as Record<string, unknown>
  bridgedWindow.aiOffice = createShellHomeApi(transport, {
    browse: async () => {
      const picked = await pickFileBytes(undefined, false)
      if (!picked) return
      const file = picked[0]
      if (!file) return
      const path = await files.writeTempFile(file.name, file.bytes)
      await transport.invoke('home:open-path', path)
    },
    pickDefaultSaveDir: async () => null,
    revealPath: async () => {},
    openTrash: async () => {},
  })
  bridgedWindow.aiOfficeProject = createShellProjectApi(transport)
  bridgedWindow.aiOfficeTabs = createShellTabsApi(transport, {
    tabsList: async () => [],
    tabsActivate: async () => {},
    tabsClose: async () => {},
    tabsShowMenu: async () => {},
    tabsShowNewMenu: async () => {},
    tabsReorder: async () => {},
  })
}
