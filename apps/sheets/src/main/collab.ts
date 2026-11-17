import { ipcMain } from 'electron'
import { CollabRooms, SHEETS_COLLAB_CHANNELS, type CollabSender } from './collab-rooms'

/**
 * The sheets app's collaboration channels.
 *
 * A room is one open workbook path. Every session that has that workbook open
 * — a desktop window and a browser tab alike — is a member of the same room,
 * and the sheets main process relays the snapshots between them.
 *
 * The channels are plain ipcMain registrations, so the web bridge dispatches
 * browser requests into the very same handlers the desktop windows use.
 */

const rooms = new CollabRooms()

/** the channels are registered once per process */
let registered = false

export function registerSheetsCollabIpc(): void {
  if (registered) return
  registered = true

  ipcMain.handle(
    SHEETS_COLLAB_CHANNELS.join,
    (event, path: unknown, _stateVector: unknown, clientId: unknown) => {
      const filePath = typeof path === 'string' ? path : ''
      const sender = event.sender as unknown as CollabSender
      return rooms.join(sender, filePath, typeof clientId === 'number' ? clientId : 0)
    },
  )

  ipcMain.on(SHEETS_COLLAB_CHANNELS.update, (event, path: unknown, snapshot: unknown) => {
    const filePath = typeof path === 'string' ? path : ''
    if (!filePath || typeof snapshot !== 'string') return
    rooms.update(event.sender as unknown as CollabSender, filePath, snapshot)
  })

  ipcMain.on(SHEETS_COLLAB_CHANNELS.awareness, (event, path: unknown, update: unknown) => {
    const filePath = typeof path === 'string' ? path : ''
    if (!filePath || typeof update !== 'string') return
    rooms.awareness(event.sender as unknown as CollabSender, filePath, update)
  })

  ipcMain.on(SHEETS_COLLAB_CHANNELS.leave, (event, path: unknown) => {
    const filePath = typeof path === 'string' ? path : ''
    if (!filePath) return
    rooms.leave(event.sender as unknown as CollabSender, filePath)
  })
}

/** @internal test seam — the rooms are process state */
export function resetSheetsCollabRooms(): void {
  rooms.reset()
}
