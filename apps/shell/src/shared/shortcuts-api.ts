/** every action a keyboard shortcut can trigger in the shell */
export type ShortcutId =
  | 'openConversationPicker'
  | 'openFileSearch'
  | 'newConversation'
  | 'goBack'
  | 'goForward'
  | 'newTab'
  | 'reopenClosedTab'
  | 'previousTab'
  | 'nextTab'
  | 'openSettings'
  | 'openRecycleBin'

/**
 * Shortcuts pressed while an editor tab (a sibling WebContentsView) has focus
 * never reach the shell window's DOM. The main process watches
 * `before-input-event` on those views, maps the chord with the same bindings
 * the shell uses, and forwards the action id to the shell.
 */
export const SHORTCUT_CHANNELS = {
  /** main → shell: an action fired while an editor view had focus */
  invoke: 'shortcuts:invoke',
  /** shell → main: the current bindings (action id → chord strings) */
  sync: 'shortcuts:sync',
} as const

export interface ShortcutsApi {
  /** an action was triggered from outside the shell DOM; returns an unsubscribe */
  onInvoke(handler: (id: ShortcutId) => void): () => void
  /** publish the current bindings so the main process can match chords on editor views */
  syncBindings(map: Record<string, string[]>): void
}
