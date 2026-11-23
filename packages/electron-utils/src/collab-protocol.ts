/**
 * Wire protocol for real-time document collaboration.
 *
 * The renderer never opens a socket of its own. It hands payloads to the shell
 * main process, which owns the rooms (one per open document path) and relays
 * them to the other members. Every session — a desktop window and a browser tab
 * alike — already lives in that one process, so a room needs no signalling
 * server of its own.
 *
 * The channels are namespaced per app: the docs app and the sheets app each
 * have their own rooms, and they share this transport.
 *
 * Payloads are strings rather than byte arrays: the browser bridge moves frames
 * as JSON text, and one encoding keeps both transports identical.
 */

/** the channels one app's collaboration layer uses */
export interface CollabChannels {
  /** renderer → main: join a room, get the state this client is missing */
  join: string
  /** renderer → main: a local state update */
  update: string
  /** renderer → main: a local awareness update */
  awareness: string
  /** renderer → main: leave the room */
  leave: string
  /** main → renderer: an update from another member */
  remoteUpdate: string
  /** main → renderer: an awareness update from another member */
  remoteAwareness: string
  /** main → renderer: the room's member count changed */
  peers: string
}

/** the channels of one app's collaboration layer */
export function collabChannels(namespace: string): CollabChannels {
  return {
    join: `${namespace}:collab:join`,
    update: `${namespace}:collab:update`,
    awareness: `${namespace}:collab:awareness`,
    leave: `${namespace}:collab:leave`,
    remoteUpdate: `${namespace}:collab:remote-update`,
    remoteAwareness: `${namespace}:collab:remote-awareness`,
    peers: `${namespace}:collab:peers`,
  }
}

export interface CollabJoinReply {
  /** the room state this client is missing — empty when the room has none */
  update: string
  /** members in the room after this join */
  peers: number
  /** true when this client owns an empty room (it has the document parsed) */
  owner: boolean
  /** the awareness ids of the room's members, so a cursor can be dropped */
  clientIds: number[]
}

/** the surface the renderer talks to; the preload installs it */
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
