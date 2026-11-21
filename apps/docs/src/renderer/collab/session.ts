import * as Y from 'yjs'
import {
  applyAwarenessUpdate,
  encodeAwarenessUpdate,
  removeAwarenessStates,
} from 'y-protocols/awareness'
import type { Node as PmNode } from '@tiptap/pm/model'
import { prosemirrorToYXmlFragment } from 'y-prosemirror'
import { COLLAB_FRAGMENT, type CollabBridge, type CollabJoinReply } from '../../shared/collab'
import { CollabAwareness } from './awareness'

/**
 * One session's end of a collaboration room.
 *
 * A room is one open document path; every session that has that document open
 * is a member of it. The session owns a yjs document whose fragment holds the
 * room's shared content, and relays its updates through the main process.
 *
 * The document is created per open document, never per session: a yjs
 * document carries the history of everything that was ever typed into it, so
 * reusing one across documents would leak one document's edits into the next
 * room.
 *
 * Seeding: the first member of an empty room owns its content — it is the
 * one that has the document parsed from disk. It publishes that content only
 * once a second member shows up, so a session that works alone never pays for
 * the CRDT at all.
 */

/** the origin tag for updates that arrived from the room rather than locally */
const REMOTE = 'collab-remote'

/** how often the session re-publishes its cursor, so peers do not time it out */
const HEARTBEAT_MS = 10_000

export interface CollabUser {
  name: string
  color: string
}

export interface CollabSessionOptions {
  bridge: CollabBridge
  user: CollabUser
  /** the awareness the editor's cursor plugin watches; a session that is not
   *  bound to an editor (a test, a headless writer) may bring its own */
  awareness?: CollabAwareness
}

function toBase64(bytes: Uint8Array): string {
  let binary = ''
  for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i])
  return btoa(binary)
}

function fromBase64(text: string): Uint8Array {
  const binary = atob(text)
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
  return bytes
}

export class CollabSession {
  readonly awareness: CollabAwareness

  private readonly bridge: CollabBridge
  private readonly user: CollabUser
  private doc: Y.Doc | null = null
  private fragment: Y.XmlFragment | null = null
  private path: string | null = null
  private owner = false
  private peers = 1
  private published = false
  private heartbeat: ReturnType<typeof setInterval> | null = null
  private readonly peerListeners = new Set<(peers: number, owner: boolean) => void>()
  private readonly contentListeners = new Set<() => void>()

  constructor(options: CollabSessionOptions) {
    this.bridge = options.bridge
    this.user = options.user
    this.awareness = options.awareness ?? new CollabAwareness()
    this.awareness.on('update', this.onAwarenessUpdate)
    this.bridge.onUpdate((path, update) => {
      if (path === this.path) this.applyRemote(update)
    })
    this.bridge.onAwareness((path, update) => {
      if (path === this.path) this.applyRemoteAwareness(update)
    })
    this.bridge.onPeers((path, peers, owner, clientIds) => {
      if (path === this.path) this.onPeers(peers, owner, clientIds)
    })
  }

  /** the yjs fragment the editor binds to, once the room has content */
  get yFragment(): Y.XmlFragment | null {
    return this.fragment
  }

  /** true when this session owns the room's initial content */
  get ownsRoom(): boolean {
    return this.owner
  }

  /** members in the room, this session included */
  get memberCount(): number {
    return this.peers
  }

  /** true once the room's content reached this session */
  get hasRoomContent(): boolean {
    return this.fragment !== null && this.fragment.length > 0
  }

  /**
   * Join the room for a document. Resolves once the session knows whether it
   * owns the room and whether the room already carries content.
   */
  async open(path: string): Promise<void> {
    await this.close()
    this.path = path
    this.owner = false
    this.published = false
    this.peers = 1
    this.doc = new Y.Doc()
    this.fragment = this.doc.getXmlFragment(COLLAB_FRAGMENT)
    this.doc.on('update', this.onDocUpdate)
    // the awareness joins the room's document: one participant, one identity
    const awareness = this.awareness.open(this.doc)
    awareness.setLocalState({ user: this.user })
    const reply: CollabJoinReply = await this.bridge.join(
      path,
      toBase64(Y.encodeStateVector(this.doc)),
      awareness.clientID,
    )
    this.peers = reply.peers
    this.owner = reply.owner
    this.dropUnknownCursors(reply.clientIds)
    if (reply.update) this.applyRemote(reply.update)
    this.notifyPeers()
    this.heartbeat = setInterval(() => {
      // the awareness protocol times a peer out when it stops hearing from it
      this.awareness.instance?.setLocalState({ user: this.user })
    }, HEARTBEAT_MS)
  }

  /** leave the room and drop the document */
  async close(): Promise<void> {
    const path = this.path
    this.path = null
    this.fragment = null
    this.owner = false
    this.published = false
    this.peers = 1
    if (this.heartbeat) {
      clearInterval(this.heartbeat)
      this.heartbeat = null
    }
    if (this.doc) {
      this.doc.off('update', this.onDocUpdate)
      // the awareness keeps the anchor document alive; the per-document one is
      // dropped without destroying it so the awareness stays usable
      this.doc = null
    }
    this.awareness.close()
    if (path) this.bridge.leave(path)
  }

  /**
   * Publish the editor's current document into the room. Only the member that
   * owns an empty room may do this, and only once — a second publication would
   * duplicate the document.
   */
  publish(doc: PmNode): boolean {
    if (!this.fragment || !this.doc || this.published || !this.owner) return false
    this.published = true
    prosemirrorToYXmlFragment(doc, this.fragment)
    return true
  }

  onPeersChanged(handler: (peers: number, owner: boolean) => void): () => void {
    this.peerListeners.add(handler)
    return () => this.peerListeners.delete(handler)
  }

  /** called when the room's content reached this session */
  onRoomContent(handler: () => void): () => void {
    this.contentListeners.add(handler)
    return () => this.contentListeners.delete(handler)
  }

  private onDocUpdate = (update: Uint8Array, origin: unknown): void => {
    if (origin === REMOTE) return
    const path = this.path
    if (!path) return
    this.bridge.update(path, toBase64(update))
  }

  private onAwarenessUpdate = (
    changes: { added: number[]; updated: number[]; removed: number[] },
    origin: unknown,
  ): void => {
    if (origin !== 'local') return
    const path = this.path
    if (!path) return
    const changed = [...changes.added, ...changes.updated, ...changes.removed]
    if (changed.length === 0) return
    const awareness = this.awareness.instance
    if (!awareness) return
    this.bridge.awareness(path, toBase64(encodeAwarenessUpdate(awareness, changed)))
  }

  private onPeers(peers: number, owner: boolean, clientIds: number[]): void {
    this.peers = peers
    this.owner = owner
    this.dropUnknownCursors(clientIds)
    this.notifyPeers()
  }

  private notifyPeers(): void {
    for (const handler of this.peerListeners) handler(this.peers, this.owner)
  }

  /** drop the cursors of members that are no longer in the room */
  private dropUnknownCursors(clientIds: number[]): void {
    const known = new Set(clientIds)
    const awareness = this.awareness.instance
    if (!awareness) return
    const stale = [...awareness.getStates().keys()].filter((id) => !known.has(id))
    if (stale.length === 0) return
    removeAwarenessStates(awareness, stale, REMOTE)
  }

  private applyRemote(update: string): void {
    if (!this.doc) return
    try {
      Y.applyUpdate(this.doc, fromBase64(update), REMOTE)
    } catch {
      return
    }
    for (const handler of this.contentListeners) handler()
  }

  private applyRemoteAwareness(update: string): void {
    const awareness = this.awareness.instance
    if (!awareness) return
    try {
      applyAwarenessUpdate(awareness, fromBase64(update), REMOTE)
    } catch {
      /* a malformed awareness payload must not take the session down */
    }
  }
}
