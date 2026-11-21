/// Preload-side bridge for real-time document collaboration.
///
/// One surface, every app: the renderer hands the shell main process the
/// payloads of the room it is in, and the main process relays them to the
/// other members. The same bridge serves a desktop window and a browser tab —
/// the web build aliases `electron` to a stub whose `ipcRenderer` rides the
/// per-tab WebSocket.
import { contextBridge, ipcRenderer } from 'electron'
import type { IpcRendererEvent } from 'electron'
import { collabChannels, type CollabBridge, type CollabJoinReply } from './collab-protocol'

export function installCollabBridge(namespace: string): void {
  const channels = collabChannels(namespace)
  const api: CollabBridge = {
    join: (path: string, stateVector: string, clientId: number): Promise<CollabJoinReply> =>
      ipcRenderer.invoke(channels.join, path, stateVector, clientId),
    update: (path: string, update: string): void => ipcRenderer.send(channels.update, path, update),
    awareness: (path: string, update: string): void =>
      ipcRenderer.send(channels.awareness, path, update),
    leave: (path: string): void => ipcRenderer.send(channels.leave, path),
    onUpdate: (handler: (path: string, update: string) => void): (() => void) => {
      const listener = (_event: IpcRendererEvent, path: string, update: string): void =>
        handler(path, update)
      ipcRenderer.on(channels.remoteUpdate, listener)
      return () => ipcRenderer.removeListener(channels.remoteUpdate, listener)
    },
    onAwareness: (handler: (path: string, update: string) => void): (() => void) => {
      const listener = (_event: IpcRendererEvent, path: string, update: string): void =>
        handler(path, update)
      ipcRenderer.on(channels.remoteAwareness, listener)
      return () => ipcRenderer.removeListener(channels.remoteAwareness, listener)
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
      ipcRenderer.on(channels.peers, listener)
      return () => ipcRenderer.removeListener(channels.peers, listener)
    },
  }
  contextBridge.exposeInMainWorld('collabApi', api)
}
