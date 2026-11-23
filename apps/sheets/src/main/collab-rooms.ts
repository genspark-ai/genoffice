import { collabChannels, type CollabJoinReply } from '@genoffice/electron-utils/collab-protocol'

/**
 * The rooms of the sheets app's collaboration layer.
 *
 * A room is one open workbook path. Its state is the workbook's file model as
 * a JSON snapshot: the newest snapshot wins (last write by arrival), and a
 * member that joins late gets the room's current snapshot.
 *
 * This is the "replace" kind of room, the sibling of the docs app's yjs rooms:
 * a spreadsheet has no ProseMirror tree to merge, so the shared state is the
 * whole model and a concurrent edit on both sides is resolved by arrival
 * order. The docs app's rooms merge instead; both share the wire protocol.
 *
 * The registry is deliberately free of any electron import: the sheets main
 * process wires it to ipcMain, and the tests drive it directly.
 */

export const SHEETS_COLLAB_CHANNELS = collabChannels('sheets')

/** the part of a WebContents a room needs; a browser tab has the same shape */
export interface CollabSender {
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
  /** the newest snapshot the room has seen; empty until somebody publishes one */
  snapshot: string
  members: Map<number, RoomMember>
  /** the member allowed to publish an empty room; null while nobody owns it */
  ownerId: number | null
}

export class CollabRooms {
  private readonly rooms = new Map<string, Room>()

  /** how many rooms are live */
  get roomCount(): number {
    return this.rooms.size
  }

  private roomFor(path: string): Room {
    let room = this.rooms.get(path)
    if (!room) {
      room = { path, snapshot: '', members: new Map(), ownerId: null }
      this.rooms.set(path, room)
    }
    return room
  }

  /** a room has content once somebody published a snapshot */
  private seeded(room: Room): boolean {
    return room.snapshot !== ''
  }

  /** the awareness ids of the room's members */
  private clientIds(room: Room): number[] {
    return [...room.members.values()].map((member) => member.clientId)
  }

  /** tell the room how many members it has, and who owns an empty room */
  private announcePeers(room: Room): void {
    const clientIds = this.clientIds(room)
    for (const [id, member] of room.members) {
      this.sendTo(
        member.sender,
        SHEETS_COLLAB_CHANNELS.peers,
        room.path,
        room.members.size,
        room.ownerId === id,
        clientIds,
      )
    }
  }

  /** relay an opaque payload to every member but the one that sent it */
  private relay(room: Room, channel: string, fromId: number, ...args: unknown[]): void {
    for (const [id, member] of room.members) {
      if (id === fromId) continue
      this.sendTo(member.sender, channel, ...args)
    }
  }

  private sendTo(sender: CollabSender, channel: string, ...args: unknown[]): void {
    if (sender.isDestroyed()) return
    sender.send(channel, ...args)
  }

  /**
   * Join a room. The first member of an empty room owns its content — it is
   * the one that has the file open. Everybody else waits for that member's
   * first snapshot, so two sessions never seed the same room with two copies
   * of the file.
   */
  join(sender: CollabSender, path: string, clientId: number): CollabJoinReply {
    if (!path) return { update: '', peers: 0, owner: false, clientIds: [] }
    const room = this.roomFor(path)
    const isNewMember = !room.members.has(sender.id)
    room.members.set(sender.id, { sender, clientId })
    if (isNewMember) {
      // a closed window or a closed browser tab leaves the room
      sender.once('destroyed', () => this.leave(sender, path))
      if (room.ownerId === null && !this.seeded(room)) room.ownerId = sender.id
      this.announcePeers(room)
    }
    return {
      update: room.snapshot,
      peers: room.members.size,
      owner: room.ownerId === sender.id && !this.seeded(room),
      clientIds: this.clientIds(room),
    }
  }

  /** publish a member's snapshot and relay it to the others */
  update(sender: CollabSender, path: string, snapshot: string): void {
    const room = this.rooms.get(path)
    if (!room || !room.members.has(sender.id)) return
    room.snapshot = snapshot
    this.relay(room, SHEETS_COLLAB_CHANNELS.remoteUpdate, sender.id, path, snapshot)
  }

  /** relay a member's awareness update to the others */
  awareness(sender: CollabSender, path: string, update: string): void {
    const room = this.rooms.get(path)
    if (!room || !room.members.has(sender.id)) return
    this.relay(room, SHEETS_COLLAB_CHANNELS.remoteAwareness, sender.id, path, update)
  }

  /** a member leaves the room */
  leave(sender: CollabSender, path: string): void {
    const room = this.rooms.get(path)
    if (!room) return
    if (!room.members.delete(sender.id)) return
    if (room.ownerId === sender.id) {
      // the owner left before publishing the room's content: hand the room to
      // the member that has been waiting longest, so the room is not stranded
      room.ownerId = room.members.size > 0 ? (room.members.keys().next().value ?? null) : null
    }
    if (room.members.size === 0) {
      this.rooms.delete(path)
      return
    }
    this.announcePeers(room)
  }

  /** drop every room — the rooms are process state */
  reset(): void {
    this.rooms.clear()
  }
}
