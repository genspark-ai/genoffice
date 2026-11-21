/**
 * Wire protocol for real-time document collaboration.
 *
 * The renderer never opens a socket of its own. It hands yjs payloads to the
 * shell main process, which owns the rooms (one per open document path) and
 * relays them to the other members. Every session — a desktop window and a
 * browser tab alike — already lives in that one process, so the room needs no
 * signalling server of its own.
 *
 * Payloads are base64 strings rather than byte arrays: the browser bridge
 * moves frames as JSON text, and one encoding keeps both transports identical.
 */

/** renderer → main */
export const COLLAB_CHANNELS = {
  /** join a room, get the state this client is missing */
  join: 'docs:collab:join',
  /** a local yjs update */
  update: 'docs:collab:update',
  /** a local awareness update */
  awareness: 'docs:collab:awareness',
  /** leave the room */
  leave: 'docs:collab:leave',
  /** main → renderer: an update from another member */
  remoteUpdate: 'docs:collab:remote-update',
  /** main → renderer: an awareness update from another member */
  remoteAwareness: 'docs:collab:remote-awareness',
  /** main → renderer: the room's member count changed */
  peers: 'docs:collab:peers',
} as const

/** the yjs fragment a room's document lives in — the same name on both sides */
export const COLLAB_FRAGMENT = 'default'

export interface CollabJoinReply {
  /** base64 yjs update carrying the room state this client is missing */
  update: string
  /** members in the room after this join */
  peers: number
  /** true when this client owns seeding an empty room with the parsed document */
  owner: boolean
  /** the awareness ids of the room's members, so a session can drop the
   *  cursors of members that left */
  clientIds: number[]
}

/** the surface the renderer talks to; the preload installs it, tests fake it */
export interface CollabBridge {
  /** `clientId` is the session's awareness id, so the room can list it */
  join(path: string, stateVector: string, clientId: number): Promise<CollabJoinReply>
  update(path: string, update: string): void
  awareness(path: string, update: string): void
  leave(path: string): void
  onUpdate(handler: (path: string, update: string) => void): () => void
  onAwareness(handler: (path: string, update: string) => void): () => void
  onPeers(
    handler: (path: string, peers: number, owner: boolean, clientIds: number[]) => void,
  ): () => void
}
