/// Preload bridge for the editors' Files pane. Every editor view lives in the
/// shell's main process, which already serves the home screen's folder-tree
/// IPC; the editors just need the same channels exposed under one name.
/// Channel strings mirror HOME_CHANNELS in apps/shell/src/shared/home-api.ts
/// (a shell test keeps them in sync).
import { contextBridge, ipcRenderer } from 'electron'
import type { IpcRendererEvent } from 'electron'

export const FILES_PANE_CHANNELS = {
  folderRoot: 'home:folder-root',
  listFolder: 'home:folder-list',
  createFolder: 'home:folder-create',
  renameFolder: 'home:folder-rename',
  renameFile: 'home:rename-file',
  movePaths: 'home:move-paths',
  deleteFolder: 'home:folder-delete',
  deleteFiles: 'home:delete-files',
  openPath: 'home:open-path',
  revealPath: 'home:reveal-path',
  folderChanged: 'home:folder-changed',
} as const

/** the global every editor preload exposes for @genoffice/ui's FilesPane */
export const FILES_PANE_GLOBAL = 'filesPaneApi'

export function buildFilesPaneApi() {
  return {
    folderRoot: () => ipcRenderer.invoke(FILES_PANE_CHANNELS.folderRoot),
    listFolder: (dir: string) => ipcRenderer.invoke(FILES_PANE_CHANNELS.listFolder, dir),
    createFolder: (parent: string, name: string) =>
      ipcRenderer.invoke(FILES_PANE_CHANNELS.createFolder, parent, name),
    renameFolder: (dir: string, newName: string) =>
      ipcRenderer.invoke(FILES_PANE_CHANNELS.renameFolder, dir, newName),
    renameFile: (path: string, newName: string) =>
      ipcRenderer.invoke(FILES_PANE_CHANNELS.renameFile, path, newName),
    movePaths: (paths: string[], targetDir: string, onConflict: string) =>
      ipcRenderer.invoke(FILES_PANE_CHANNELS.movePaths, paths, targetDir, onConflict),
    deleteFolder: (dir: string) => ipcRenderer.invoke(FILES_PANE_CHANNELS.deleteFolder, dir),
    deleteFiles: (paths: string[]) => ipcRenderer.invoke(FILES_PANE_CHANNELS.deleteFiles, paths),
    openPath: (path: string) => ipcRenderer.invoke(FILES_PANE_CHANNELS.openPath, path),
    revealPath: (path: string) => ipcRenderer.invoke(FILES_PANE_CHANNELS.revealPath, path),
    onFolderChanged: (handler: (dirs: string[]) => void) => {
      const listener = (_event: IpcRendererEvent, dirs: unknown) => {
        if (Array.isArray(dirs)) handler(dirs.filter((d): d is string => typeof d === 'string'))
      }
      ipcRenderer.on(FILES_PANE_CHANNELS.folderChanged, listener)
      return () => ipcRenderer.removeListener(FILES_PANE_CHANNELS.folderChanged, listener)
    },
  }
}

export function installFilesPaneBridge(): void {
  contextBridge.exposeInMainWorld(FILES_PANE_GLOBAL, buildFilesPaneApi())
}
