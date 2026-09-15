/**
 * Collab module entry — wires session/lock/cursor/change/conflict/
 * permissions/history/comments/templates channels into the shared
 * registry. State lives in `common/state.ts`.
 */
import {
  registerChangeTrackingHandlers,
  registerConflictHandlers,
  registerCursorHandlers,
  registerLockHandlers,
  registerPermissionHandlers,
} from './locks'
import {
  registerCommentHandlers,
  registerHistoryHandlers,
  registerTemplateHandlers,
} from './history-comments-templates'
import { registerCollabSessionHandlers } from './sessions'

export function registerCollabHandlers(): void {
  registerCollabSessionHandlers()
  registerLockHandlers()
  registerCursorHandlers()
  registerChangeTrackingHandlers()
  registerConflictHandlers()
  registerPermissionHandlers()
  registerHistoryHandlers()
  registerCommentHandlers()
  registerTemplateHandlers()
}
