import type { Node as PmNode } from '@tiptap/pm/model'
import { CollabSession, type CollabUser } from './session'
import { collabSyncController } from './sync-extension'
import { authorName } from '../author-name'
import type { CollabBridge } from '../../shared/collab'

export interface CollabCoordinatorDeps {
  bridge: CollabBridge
  /** the editor's current document, used to seed an empty room */
  currentDoc: () => PmNode | null
  user: () => CollabUser
}

/**
 * Drives one session's collaboration: it joins the room of the open document,
 * publishes the document when a second member shows up, and binds the editor
 * to the room's CRDT.
 *
 * A session that works alone never binds: the editor keeps the document it
 * parsed, exactly as before, and the CRDT costs nothing. The room only takes
 * over once somebody else has the same document open.
 */
export class CollabCoordinator {
  private session: CollabSession | null = null
  private ready = false
  private bound = false

  constructor(private readonly deps: CollabCoordinatorDeps) {}

  /** true while the editor is driven by a shared room */
  get isShared(): boolean {
    return this.bound
  }

  /** the document was opened: join its room */
  async openDocument(path: string): Promise<void> {
    this.closeDocument()
    if (!path) return
    const session = new CollabSession({
      bridge: this.deps.bridge,
      user: this.deps.user(),
      awareness: collabSyncController()?.awareness,
    })
    this.session = session
    session.onPeersChanged((peers, owner) => this.onPeers(session, peers, owner))
    session.onRoomContent(() => this.sync())
    await session.open(path)
    this.sync()
  }

  /** the document finished loading into the editor */
  documentReady(): void {
    this.ready = true
    this.sync()
  }

  /** the document was closed or replaced */
  closeDocument(): void {
    this.bound = false
    this.ready = false
    const session = this.session
    this.session = null
    collabSyncController()?.detach()
    void session?.close()
  }

  private onPeers(session: CollabSession, peers: number, owner: boolean): void {
    if (peers > 1 && owner) {
      // the member that owns an empty room publishes what is on its screen
      const doc = this.deps.currentDoc()
      if (doc) session.publish(doc)
    }
    this.sync()
  }

  /**
   * Bind the editor to the room once the room has content and the document
   * finished loading. Binding earlier would push the local parse into the room
   * before the room's own content arrived.
   */
  private sync(): void {
    const session = this.session
    if (!session || this.bound || !this.ready) return
    if (session.memberCount < 2 || !session.hasRoomContent) return
    const ctrl = collabSyncController()
    const fragment = session.yFragment
    if (!ctrl || !fragment) return
    ctrl.attach(fragment)
    this.bound = true
  }
}

let coordinator: CollabCoordinator | null = null

/** the coordinator of this renderer, if one was set up */
export function collabCoordinator(): CollabCoordinator | null {
  return coordinator
}

/** a stable colour for the local participant, derived from its name */
function colorFor(name: string): string {
  let hash = 0
  for (let i = 0; i < name.length; i++) hash = (hash * 31 + name.charCodeAt(i)) % 360
  return `hsl(${hash}, 70%, 55%)`
}

/**
 * The coordinator of this renderer. The preload installs the bridge, so a
 * renderer without it (a headless writer, a test) simply has no room.
 */
export function ensureCollabCoordinator(): CollabCoordinator | null {
  if (coordinator) return coordinator
  const bridge = (window as unknown as { collabApi?: CollabBridge }).collabApi
  if (!bridge) return null
  coordinator = new CollabCoordinator({
    bridge,
    currentDoc: () => collabSyncController()?.currentDoc ?? null,
    user: () => {
      const name = authorName().trim() || 'Guest'
      return { name, color: colorFor(name) }
    },
  })
  return coordinator
}
