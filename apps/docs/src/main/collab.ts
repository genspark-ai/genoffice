import { ipcMain } from 'electron'
import { COLLAB_CHANNELS } from '../shared/collab'
import { CollabRooms, type CollabSender } from './collab-rooms'

/**
 * The collaboration channels of the shell main process.
 *
 * The rooms themselves live in `CollabRooms`, which knows nothing about
 * electron: this file only wires them to ipcMain. Because the web bridge's ipc
 * capture mirrors every registration, the browser reaches the very same
 * handlers the desktop windows use — one relay, both transports.
 */

const rooms = new CollabRooms()

/** @internal test seam — the rooms are process state */
export function resetCollabRooms(): void {
  rooms.reset()
}

/** @internal test seam */
export function collabRoomCount(): number {
  return rooms.roomCount
}

export function registerCollabIpc(): void {
  ipcMain.handle(
    COLLAB_CHANNELS.join,
    (event, path: unknown, stateVector: unknown, clientId: unknown) => {
      const filePath = typeof path === 'string' ? path : ''
      const sender = event.sender as unknown as CollabSender
      return rooms.join(
        sender,
        filePath,
        typeof stateVector === 'string' ? stateVector : '',
        typeof clientId === 'number' ? clientId : 0,
      )
    },
  )

  ipcMain.on(COLLAB_CHANNELS.update, (event, path: unknown, update: unknown) => {
    const filePath = typeof path === 'string' ? path : ''
    if (!filePath || typeof update !== 'string') return
    rooms.update(event.sender as unknown as CollabSender, filePath, update)
  })

  ipcMain.on(COLLAB_CHANNELS.awareness, (event, path: unknown, update: unknown) => {
    const filePath = typeof path === 'string' ? path : ''
    if (!filePath || typeof update !== 'string') return
    rooms.awareness(event.sender as unknown as CollabSender, filePath, update)
  })

  ipcMain.on(COLLAB_CHANNELS.leave, (event, path: unknown) => {
    const filePath = typeof path === 'string' ? path : ''
    if (!filePath) return
    rooms.leave(event.sender as unknown as CollabSender, filePath)
  })
}
