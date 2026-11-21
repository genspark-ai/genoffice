/// Preload-side bridge for real-time document collaboration (docs).
///
/// The same surface serves a desktop window and a browser tab: the web build
/// aliases `electron` to a stub whose `ipcRenderer` rides the per-tab
/// WebSocket, so this file needs no browser-specific branch.
import { contextBridge, ipcRenderer } from 'electron'
import type { IpcRendererEvent } from 'electron'
import { COLLAB_CHANNELS, type CollabBridge, type CollabJoinReply } from '../shared/collab'

export function installCollabBridge(): void {
  const api: CollabBridge = {
    join: (path: string, stateVector: string, clientId: number): Promise<CollabJoinReply> =>
      ipcRenderer.invoke(COLLAB_CHANNELS.join, path, stateVector, clientId),
    update: (path: string, update: string): void =>
      ipcRenderer.send(COLLAB_CHANNELS.update, path, update),
    awareness: (path: string, update: string): void =>
      ipcRenderer.send(COLLAB_CHANNELS.awareness, path, update),
    leave: (path: string): void => ipcRenderer.send(COLLAB_CHANNELS.leave, path),
    onUpdate: (handler: (path: string, update: string) => void): (() => void) => {
      const listener = (_event: IpcRendererEvent, path: string, update: string): void =>
        handler(path, update)
      ipcRenderer.on(COLLAB_CHANNELS.remoteUpdate, listener)
      return () => ipcRenderer.removeListener(COLLAB_CHANNELS.remoteUpdate, listener)
    },
    onAwareness: (handler: (path: string, update: string) => void): (() => void) => {
      const listener = (_event: IpcRendererEvent, path: string, update: string): void =>
        handler(path, update)
      ipcRenderer.on(COLLAB_CHANNELS.remoteAwareness, listener)
      return () => ipcRenderer.removeListener(COLLAB_CHANNELS.remoteAwareness, listener)
    },
    onPeers: (
      handler: (path: string, peers: number, owner: boolean, clientIds: number[]) => void,
    ): (() => void) => {
      const listener = (
        _event: IpcRendererEvent,
        path: string,
        peers: number,
        owner: boolean,
        clientIds: number[],
      ): void => handler(path, peers, owner, clientIds)
      ipcRenderer.on(COLLAB_CHANNELS.peers, listener)
      return () => ipcRenderer.removeListener(COLLAB_CHANNELS.peers, listener)
    },
  }
  contextBridge.exposeInMainWorld('collabApi', api)
}
