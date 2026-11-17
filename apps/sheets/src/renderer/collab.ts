import type { CollabBridge } from '@genoffice/electron-utils/collab-protocol'
import { setCollabPresence, type CollabPresenceState } from '@genoffice/ui'

/**
 * The sheets app's collaboration client.
 *
 * A session that has a workbook open joins the room of that workbook's path.
 * The shared state is the workbook's file model as a JSON snapshot: the
 * newest snapshot wins, and a member that joins late gets the room's current
 * snapshot.
 *
 * A session that works alone never publishes: the room's member count is the
 * switch, so a lone editor pays nothing and behaves exactly as before.
 */

/** what a member of the room is doing right now */
export interface SheetsPresence {
  /** the member's awareness id */
  id: number
  /** the member's display name */
  name: string
  /** the id of the sheet the member is on */
  sheetId?: string
  /** the name of that sheet */
  sheetName?: string
  /** the range the member has selected, in A1 notation */
  range?: string
}

export interface SheetsCollabDeps {
  bridge: CollabBridge
  /** the workbook's file model, serialized */
  snapshot: () => string
  /** apply a snapshot that arrived from another member */
  apply: (snapshot: string) => void
  /** the local participant's presence (which sheet, which range) */
  presence: () => string
  /** how long to wait after the last edit before publishing */
  debounceMs?: number
  /** snapshots larger than this are not published — a workbook that big is
   *  cheaper to re-read from disk than to push on every keystroke */
  maxSnapshotBytes?: number
}

const DEFAULT_DEBOUNCE_MS = 400

const DEFAULT_MAX_SNAPSHOT_BYTES = 1_000_000

export class SheetsCollab {
  private path: string | null = null
  private timer: ReturnType<typeof setTimeout> | null = null
  private unsubs: Array<() => void> = []
  private peers = 1
  private readonly clientId = Math.floor(Math.random() * 0x7fffffff)
  private presence = new Map<number, CollabPresenceState>()

  constructor(private readonly deps: SheetsCollabDeps) {}

  /** members in this room, this session included */
  get memberCount(): number {
    return this.peers
  }

  /** open the room of a workbook */
  async open(path: string): Promise<void> {
    await this.close()
    if (!path) return
    this.path = path
    this.unsubs.push(
      this.deps.bridge.onUpdate((room, snapshot) => {
        if (room !== this.path) return
        this.deps.apply(snapshot)
      }),
    )
    this.unsubs.push(
      this.deps.bridge.onPeers((room, peers, _owner, clientIds) => {
        if (room !== this.path) return
        this.peers = peers
        // a member that left takes its cursor with it
        const known = new Set(clientIds)
        const stale = [...this.presence.keys()].filter((id) => !known.has(id))
        if (stale.length > 0) {
          const next = new Map(this.presence)
          for (const id of stale) next.delete(id)
          this.presence = next
          setCollabPresence(next)
        }
        // a second member showed up: the room needs a snapshot to start from
        if (peers > 1) this.publish()
      }),
    )
    this.unsubs.push(
      this.deps.bridge.onAwareness((room, raw) => {
        if (room !== this.path) return
        this.applyPresence(raw)
      }),
    )
    const reply = await this.deps.bridge.join(path, '', this.clientId)
    this.peers = reply.peers
    if (reply.update) this.deps.apply(reply.update)
    if (reply.peers > 1) this.publish()
  }

  /** leave the room */
  async close(): Promise<void> {
    const path = this.path
    this.path = null
    this.peers = 1
    if (this.timer) {
      clearTimeout(this.timer)
      this.timer = null
    }
    for (const unsub of this.unsubs) unsub()
    this.unsubs = []
    if (path) this.deps.bridge.leave(path)
  }

  /** the workbook changed: publish the new state, if somebody is listening */
  publish(): void {
    if (!this.path || this.peers < 2) return
    if (this.timer) clearTimeout(this.timer)
    this.timer = setTimeout(() => {
      this.timer = null
      this.publishNow()
    }, this.deps.debounceMs ?? DEFAULT_DEBOUNCE_MS)
  }

  /** publish the current state right away */
  publishNow(): void {
    const path = this.path
    if (!path || this.peers < 2) return
    const snapshot = this.deps.snapshot()
    const max = this.deps.maxSnapshotBytes ?? DEFAULT_MAX_SNAPSHOT_BYTES
    if (snapshot.length > max) return
    this.deps.bridge.update(path, snapshot)
  }

  /** publish the local presence */
  publishPresence(): void {
    const path = this.path
    if (!path || this.peers < 2) return
    this.deps.bridge.awareness(path, this.deps.presence())
  }

  /** the local member's awareness id — the presence carries it */
  get awarenessId(): number {
    return this.clientId
  }

  private applyPresence(raw: string): void {
    let presence: SheetsPresence
    try {
      presence = JSON.parse(raw) as SheetsPresence
    } catch {
      return
    }
    if (!presence || typeof presence.id !== 'number') return
    const next = new Map(this.presence)
    next.set(presence.id, {
      id: presence.id,
      name: presence.name,
      location: presence.sheetName
        ? `${presence.sheetName}${presence.range ? `!${presence.range}` : ''}`
        : undefined,
      sheetId: presence.sheetId,
      sheetName: presence.sheetName,
      range: presence.range,
    })
    this.presence = next
    setCollabPresence(next)
  }
}
