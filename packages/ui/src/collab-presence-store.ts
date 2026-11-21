/**
 * The presence store of a shared document.
 *
 * The status bar renders the members of the room; it is a different component
 * from the one that owns the session, so the presence travels through this
 * tiny external store instead of prop threading.
 */

export interface CollabPresenceState {
  /** the member's awareness id — the colour is derived from it */
  id: number
  /** the member's display name */
  name: string
  /** where the member is right now (a sheet name, a slide name) */
  location?: string | undefined
  /** the sheet the member is on, when the app has sheets */
  sheetId?: string | undefined
  /** the sheet's name */
  sheetName?: string | undefined
  /** the range the member has selected, in A1 notation */
  range?: string | undefined
}

let state: ReadonlyMap<number, CollabPresenceState> = new Map()
const listeners = new Set<() => void>()

/** the members of the room, keyed by their awareness id */
export function currentCollabPresence(): ReadonlyMap<number, CollabPresenceState> {
  return state
}

/** subscribe the chrome to the presence store */
export function subscribeCollabPresence(listener: () => void): () => void {
  listeners.add(listener)
  return () => {
    listeners.delete(listener)
  }
}

/** replace the presence state and notify the subscribers */
export function setCollabPresence(next: ReadonlyMap<number, CollabPresenceState>): void {
  state = next
  for (const listener of listeners) listener()
}
