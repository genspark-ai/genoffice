/**
 * Collaboration lifecycle — join/leave/sync/presence. The session map
 * (`COLLAB_SESSIONS`) and presence map (`PRESENCE`) live in
 * `common/state.ts`; this module just wires the channels.
 */
import { COLLAB_SESSIONS, PRESENCE, registerHandle } from '../common/index.js'

export function registerCollabSessionHandlers(): void {
  registerHandle('collab:join', (_event: unknown, args: unknown) => {
    const { docId, userId, userName } = args as { docId: string; userId: string; userName?: string }
    const sessionId = `${docId}:${userId}`

    if (!COLLAB_SESSIONS.has(docId)) {
      COLLAB_SESSIONS.set(docId, {
        docId,
        users: new Set(),
        lastActivity: Date.now(),
        locks: new Map(),
        cursors: new Map(),
        changes: [],
      })
    }

    const session = COLLAB_SESSIONS.get(docId)!
    session.users.add(userId)
    session.lastActivity = Date.now()

    return { sessionId, users: [...session.users], docId, userCount: session.users.size }
  })

  registerHandle('collab:leave', (_event: unknown, args: unknown) => {
    const { docId, userId } = args as { docId: string; userId: string }
    const session = COLLAB_SESSIONS.get(docId)
    if (session) {
      session.users.delete(userId)
      if (session.users.size === 0) {
        COLLAB_SESSIONS.delete(docId)
      }
    }
    return { ok: true }
  })

  registerHandle('collab:sync', (_event: unknown, args: unknown) => {
    const { docId, userId } = args as { docId: string; changes: unknown; userId: string }
    const session = COLLAB_SESSIONS.get(docId)
    if (session) {
      session.lastActivity = Date.now()
      return {
        ok: true,
        acknowledged: true,
        users: [...session.users],
        timestamp: Date.now(),
        serverTime: Date.now(),
      }
    }
    return { ok: false, error: 'Session not found' }
  })

  registerHandle('collab:presence-update', (_event: unknown, args: unknown) => {
    const { docId, userId, userName, status, cursor } = args as {
      docId: string
      userId: string
      userName?: string
      status?: 'active' | 'idle' | 'away'
      cursor?: { x: number; y: number; selection?: { start: number; end: number } }
    }

    if (!PRESENCE.has(docId)) {
      PRESENCE.set(docId, new Map())
    }

    const docPresence = PRESENCE.get(docId)!
    const colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7', '#DDA0DD', '#98D8C8', '#F7DC6F']
    const colorIndex = userId.charCodeAt(0) % colors.length

    docPresence.set(userId, {
      userId,
      userName: userName || userId,
      status: status || 'active',
      lastSeen: Date.now(),
      cursor,
      color: colors[colorIndex],
    })

    return { ok: true }
  })

  registerHandle('collab:presence-list', (_event: unknown, args: unknown) => {
    const { docId } = args as { docId: string }
    const docPresence = PRESENCE.get(docId)
    if (!docPresence) return []

    const now = Date.now()
    for (const [uid, p] of docPresence) {
      if (now - p.lastSeen > 60000) {
        docPresence.delete(uid)
      }
    }

    return [...docPresence.values()]
  })
}
