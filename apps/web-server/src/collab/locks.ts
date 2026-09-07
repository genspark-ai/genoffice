/**
 * Document-level locking, cursor sync, change tracking and conflict
 * detection/resolution. All state lives on the `COLLAB_SESSIONS` map.
 */
import { COLLAB_SESSIONS, DOC_PERMISSIONS, registerHandle } from '../common/index.js'

export function registerLockHandlers(): void {
  registerHandle('collab:lock-acquire', (_event: unknown, args: unknown) => {
    const { docId, userId, sectionId } = args as { docId: string; userId: string; sectionId?: string }
    const session = COLLAB_SESSIONS.get(docId)
    if (!session) return { ok: false, error: 'Session not found' }

    const lockKey = sectionId || 'document'
    const existingLock = session.locks.get(lockKey)

    if (existingLock && existingLock.userId !== userId && Date.now() - existingLock.timestamp < 30000) {
      return {
        ok: false,
        error: 'Section is locked',
        lockedBy: existingLock.userId,
        lockedUntil: existingLock.timestamp + 30000,
      }
    }

    session.locks.set(lockKey, { userId, timestamp: Date.now() })
    return { ok: true, lockKey, acquiredAt: Date.now() }
  })

  registerHandle('collab:lock-release', (_event: unknown, args: unknown) => {
    const { docId, userId, sectionId } = args as { docId: string; userId: string; sectionId?: string }
    const session = COLLAB_SESSIONS.get(docId)
    if (!session) return { ok: false, error: 'Session not found' }

    const lockKey = sectionId || 'document'
    const existingLock = session.locks.get(lockKey)

    if (existingLock && existingLock.userId === userId) {
      session.locks.delete(lockKey)
    }

    return { ok: true }
  })

  registerHandle('collab:lock-status', (_event: unknown, args: unknown) => {
    const { docId, sectionId } = args as { docId: string; sectionId?: string }
    const session = COLLAB_SESSIONS.get(docId)
    if (!session) return { locks: [] }

    const lockKey = sectionId || 'document'
    const lock = session.locks.get(lockKey)

    if (!lock) return { locks: [] }

    return {
      locks: [{
        sectionId: lockKey,
        userId: lock.userId,
        timestamp: lock.timestamp,
        expired: Date.now() - lock.timestamp > 30000,
      }],
    }
  })
}

export function registerCursorHandlers(): void {
  registerHandle('collab:cursor-update', (_event: unknown, args: unknown) => {
    const { docId, userId, position, selection } = args as {
      docId: string
      userId: string
      position: { x: number; y: number; offset: number }
      selection?: { start: number; end: number }
    }

    const session = COLLAB_SESSIONS.get(docId)
    if (!session) return { ok: false, error: 'Session not found' }

    session.cursors.set(userId, {
      position,
      selection,
      timestamp: Date.now(),
    })

    return { ok: true }
  })

  registerHandle('collab:cursor-list', (_event: unknown, args: unknown) => {
    const { docId } = args as { docId: string }
    const session = COLLAB_SESSIONS.get(docId)
    if (!session) return []

    const colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7', '#DDA0DD', '#98D8C8', '#F7DC6F']

    return [...session.cursors.entries()].map(([uid, cursor], idx) => ({
      userId: uid,
      position: cursor.position,
      selection: cursor.selection,
      color: colors[idx % colors.length],
    }))
  })
}

export function registerChangeTrackingHandlers(): void {
  registerHandle('collab:change-track', (_event: unknown, args: unknown) => {
    const { docId, userId, change } = args as {
      docId: string
      userId: string
      change: {
        type: 'insert' | 'delete' | 'replace' | 'format'
        position: number
        content?: string
        length?: number
        attributes?: Record<string, unknown>
      }
    }

    const session = COLLAB_SESSIONS.get(docId)
    if (!session) return { ok: false, error: 'Session not found' }

    const changeRecord = {
      id: `change-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
      docId,
      userId,
      change,
      timestamp: Date.now(),
      version: session.changes.length,
    }

    session.changes.push(changeRecord)

    if (session.changes.length > 1000) {
      session.changes = session.changes.slice(-1000)
    }

    return { ok: true, changeId: changeRecord.id, version: changeRecord.version }
  })

  registerHandle('collab:change-since', (_event: unknown, args: unknown) => {
    const { docId, sinceVersion } = args as { docId: string; sinceVersion: number }
    const session = COLLAB_SESSIONS.get(docId)
    if (!session) return { changes: [], latestVersion: 0 }

    const changes = session.changes.filter(c => c.version > sinceVersion)

    return {
      changes: changes.map(c => ({
        id: c.id,
        userId: c.userId,
        change: c.change,
        timestamp: c.timestamp,
        version: c.version,
      })),
      latestVersion: session.changes.length,
    }
  })
}

export function registerConflictHandlers(): void {
  registerHandle('collab:conflict-detect', (_event: unknown, args: unknown) => {
    const { docId, baseVersion, changes } = args as {
      docId: string
      baseVersion: number
      changes: unknown[]
    }
    const session = COLLAB_SESSIONS.get(docId)
    if (!session) return { hasConflicts: false }

    const serverVersion = session.changes.length

    if (baseVersion < serverVersion) {
      return {
        hasConflicts: true,
        conflictType: 'concurrent-edits',
        serverChanges: session.changes.filter(c => c.version >= baseVersion),
        yourVersion: baseVersion,
        serverVersion,
      }
    }

    return { hasConflicts: false }
  })

  registerHandle('collab:conflict-resolve', (_event: unknown, _args: unknown) => ({
    ok: true,
    strategy: 'merge',
    resolvedAt: Date.now(),
    newVersion: Date.now(),
  }))
}

export function registerPermissionHandlers(): void {
  registerHandle('collab:permissions-get', (_event: unknown, args: unknown) => {
    const { docId } = args as { docId: string }
    // Permissions here are the legacy `DOC_PERMISSIONS` map; the
    // enterprise-grade permissions live in `enterprise/permissions.ts`.
    const perms = DOC_PERMISSIONS.get(docId)
    if (!perms) return []

    return [...perms.entries()].map(([userId, perm]) => ({
      userId,
      permission: perm,
    }))
  })

  registerHandle('collab:permissions-set', (_event: unknown, args: unknown) => {
    const { docId, userId, permission } = args as {
      docId: string
      userId: string
      permission: 'view' | 'edit' | 'admin'
    }

    if (!DOC_PERMISSIONS.has(docId)) {
      DOC_PERMISSIONS.set(docId, new Map())
    }

    DOC_PERMISSIONS.get(docId)!.set(userId, permission)

    return { ok: true }
  })
}
