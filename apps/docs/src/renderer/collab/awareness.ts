import { Awareness } from 'y-protocols/awareness'
import type * as Y from 'yjs'

/**
 * The awareness the editor's cursor plugin watches.
 *
 * y-protocols' awareness starts a timer that prunes peers that stopped
 * sending, and an editor that never joins a room must not leave a timer
 * behind — the editor's own tests count timers and abort on a loop. The real
 * awareness is therefore created when a room opens; until then this facade
 * answers the cursor plugin with an empty world.
 */
export class CollabAwareness {
  private real: Awareness | null = null
  private readonly listeners: Array<(changes: AwarenessChanges, origin: unknown) => void> = []

  /** the real awareness, once a room opened */
  get instance(): Awareness | null {
    return this.real
  }

  /** the room's document is the awareness' home: one participant, one id */
  open(doc: Y.Doc): Awareness {
    if (this.real) return this.real
    const awareness = new Awareness(doc)
    this.real = awareness
    for (const listener of this.listeners) awareness.on('update', listener)
    return awareness
  }

  /** drop the awareness: the session left the room */
  close(): void {
    this.real?.setLocalState(null)
    this.real = null
  }

  // --- the surface y-prosemirror's cursor plugin uses ---

  get clientID(): number {
    return this.real?.clientID ?? 0
  }

  getStates(): Map<number, Record<string, unknown>> {
    return (this.real?.getStates() ?? new Map()) as Map<number, Record<string, unknown>>
  }

  getLocalState(): Record<string, unknown> | null {
    return (this.real?.getLocalState() ?? null) as Record<string, unknown> | null
  }

  setLocalStateField(field: string, value: unknown): void {
    this.real?.setLocalStateField(field, value)
  }

  on(event: 'update', listener: (changes: AwarenessChanges, origin: unknown) => void): void {
    if (event !== 'update') return
    this.listeners.push(listener)
    this.real?.on('update', listener)
  }

  off(event: 'update', listener: (changes: AwarenessChanges, origin: unknown) => void): void {
    if (event !== 'update') return
    const index = this.listeners.indexOf(listener)
    if (index >= 0) this.listeners.splice(index, 1)
    this.real?.off('update', listener)
  }
}

export interface AwarenessChanges {
  added: number[]
  updated: number[]
  removed: number[]
}
