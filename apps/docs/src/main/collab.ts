import * as Y from 'yjs'
import { ipcMain } from 'electron'
import { COLLAB_CHANNELS, COLLAB_FRAGMENT, type CollabJoinReply } from '../shared/collab'

/**
 * Rooms for real-time document collaboration.
 *
 * A room is one open document path. Every session that has that document open
 * — a desktop window and a browser tab alike — is a member of the same room,
 * and the shell main process is the relay between them. There is no separate
 * signalling server: every session already lives in this process, so the room
 * is just a registry plus a yjs document that mirrors the room's state.
 *
 * The mirror never produces updates of its own; it only applies what members
 * send and hands the accumulated state to a late joiner. A late joiner asks
 * with its state vector and receives exactly the part it is missing.
 *
 * Seeding: the first member of an empty room owns the room's content — it is
 * the one that parsed the document from disk. Everyone else waits for that
 * seed, so two sessions never seed the same room with two copies of the file.
 */

interface CollabSender {
  id: number
  send(channel: string, ...args: unknown[]): void
  once(event: 'destroyed', listener: () => void): unknown
  isDestroyed(): boolean
}

interface RoomMember {
  sender: CollabSender
  /** the member's awareness id, so its cursor can be dropped when it leaves */
  clientId: number
}

interface Room {
  path: string
  doc: Y.Doc
  members: Map<number, RoomMember>
  /** the member allowed to seed an empty room; null while nobody owns it */
  ownerId: number | null
}

const rooms = new Map<string, Room>()

/** the channels are registered once per process */
let registered = false

function roomFor(path: string): Room {
  let room = rooms.get(path)
  if (!room) {
    room = { path, doc: new Y.Doc(), members: new Map(), ownerId: null }
    rooms.set(path, room)
  }
  return room
}

/** a room has content once somebody wrote into its document fragment */
function roomSeeded(room: Room): boolean {
  return room.doc.getXmlFragment(COLLAB_FRAGMENT).length > 0
}

function toBase64(bytes: Uint8Array): string {
  return Buffer.from(bytes).toString('base64')
}

function fromBase64(text: string): Uint8Array {
  return new Uint8Array(Buffer.from(text, 'base64'))
}

/** the part of the room state the joining client is missing */
function roomStateFor(room: Room, stateVector: string): string {
  let vector: Uint8Array
  try {
    vector = stateVector ? fromBase64(stateVector) : new Uint8Array()
  } catch {
    vector = new Uint8Array()
  }
  return toBase64(Y.encodeStateAsUpdate(room.doc, vector))
}

function sendTo(sender: CollabSender, channel: string, ...args: unknown[]): void {
  if (sender.isDestroyed()) return
  sender.send(channel, ...args)
}

/** the awareness ids of the room's members */
function roomClientIds(room: Room): number[] {
  return [...room.members.values()].map((member) => member.clientId)
}

/** tell the room how many members it has, and who owns an empty room */
function announcePeers(room: Room): void {
  const clientIds = roomClientIds(room)
  for (const [id, member] of room.members) {
    sendTo(
      member.sender,
      COLLAB_CHANNELS.peers,
      room.path,
      room.members.size,
      room.ownerId === id,
      clientIds,
    )
  }
}

/** relay an opaque payload to every member but the one that sent it */
function relay(room: Room, channel: string, fromId: number, ...args: unknown[]): void {
  for (const [id, member] of room.members) {
    if (id === fromId) continue
    sendTo(member.sender, channel, ...args)
  }
}

function leaveRoom(path: string, memberId: number): void {
  const room = rooms.get(path)
  if (!room) return
  const member = room.members.get(memberId)
  if (!member) return
  room.members.delete(memberId)
  if (room.ownerId === memberId) {
    // the owner left before publishing the room's content: hand the room to
    // the member that has been waiting longest, so the room is not stranded
    room.ownerId = room.members.size > 0 ? (room.members.keys().next().value ?? null) : null
  }
  if (room.members.size === 0) {
    rooms.delete(path)
    room.doc.destroy()
    return
  }
  announcePeers(room)
}

/**
 * Register the collaboration channels. Called from the docs app's IPC
 * registration so the web bridge reaches them too: the bridge dispatches
 * browser requests into the very same ipcMain handlers the desktop windows
 * use, and every session lives in this one process.
 */
export function registerCollabIpc(): void {
  if (registered) return
  registered = true
  ipcMain.handle(
    COLLAB_CHANNELS.join,
    (event, path: unknown, stateVector: unknown, clientId: unknown): CollabJoinReply => {
      const filePath = typeof path === 'string' ? path : ''
      if (!filePath) return { update: '', peers: 0, owner: false, clientIds: [] }
      const sender = event.sender as unknown as CollabSender
      const room = roomFor(filePath)
      const isNewMember = !room.members.has(sender.id)
      room.members.set(sender.id, {
        sender,
        clientId: typeof clientId === 'number' ? clientId : 0,
      })
      if (isNewMember) {
        // a closed window or a closed browser tab leaves the room
        sender.once('destroyed', () => leaveRoom(filePath, sender.id))
        if (room.ownerId === null && !roomSeeded(room)) room.ownerId = sender.id
        announcePeers(room)
      }
      return {
        update: roomStateFor(room, typeof stateVector === 'string' ? stateVector : ''),
        peers: room.members.size,
        owner: room.ownerId === sender.id && !roomSeeded(room),
        clientIds: roomClientIds(room),
      }
    },
  )

  ipcMain.on(COLLAB_CHANNELS.update, (event, path: unknown, update: unknown) => {
    const filePath = typeof path === 'string' ? path : ''
    if (!filePath || typeof update !== 'string') return
    const room = rooms.get(filePath)
    if (!room) return
    const sender = event.sender as unknown as CollabSender
    if (!room.members.has(sender.id)) return
    try {
      Y.applyUpdate(room.doc, fromBase64(update), 'collab-relay')
    } catch {
      return
    }
    relay(room, COLLAB_CHANNELS.remoteUpdate, sender.id, filePath, update)
  })

  ipcMain.on(COLLAB_CHANNELS.awareness, (event, path: unknown, update: unknown) => {
    const filePath = typeof path === 'string' ? path : ''
    if (!filePath || typeof update !== 'string') return
    const room = rooms.get(filePath)
    if (!room) return
    const sender = event.sender as unknown as CollabSender
    if (!room.members.has(sender.id)) return
    relay(room, COLLAB_CHANNELS.remoteAwareness, sender.id, filePath, update)
  })

  ipcMain.on(COLLAB_CHANNELS.leave, (event, path: unknown) => {
    const filePath = typeof path === 'string' ? path : ''
    if (!filePath) return
    const sender = event.sender as unknown as CollabSender
    leaveRoom(filePath, sender.id)
  })
}

/** @internal test seam — the rooms are process state, tests need a clean slate */
export function resetCollabRooms(): void {
  for (const room of rooms.values()) room.doc.destroy()
  rooms.clear()
}

/** @internal test seam */
export function collabRoomCount(): number {
  return rooms.size
}
