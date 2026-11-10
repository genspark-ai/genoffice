import { ipcMain } from 'electron'
import type { WebContents } from 'electron'
import { chordFromEvent, normalizeChord } from '../shared/chords'
import { SHORTCUT_CHANNELS } from '../shared/shortcuts-api'
import type { ShortcutId } from '../shared/shortcuts-api'

/**
 * Editor tabs are separate WebContentsViews: key presses inside them never reach the shell
 * window, so tab shortcuts (Ctrl+T, Ctrl+Shift+T, Alt+Up/Down…) would go dead while editing.
 * This watches `before-input-event` on each editor view, matches the chord against the
 * bindings the shell publishes, and forwards the action to the shell.
 *
 * Only actions that don't collide with editor shortcuts are forwarded. Ctrl+K (insert link) and
 * Ctrl+P (print) stay with the editor; the picker/search shortcuts work from Home.
 */
const FORWARDED = new Set<ShortcutId>([
  'newTab',
  'reopenClosedTab',
  'previousTab',
  'nextTab',
  'openSettings',
  'openRecycleBin',
])

const CYCLING = new Set<ShortcutId>(['previousTab', 'nextTab'])

/** chord → action, rebuilt whenever the shell publishes its bindings */
let chordToAction = new Map<string, ShortcutId>()
let initialised = false

function rebuild(raw: unknown): void {
  const next = new Map<string, ShortcutId>()
  if (raw && typeof raw === 'object') {
    for (const [id, chords] of Object.entries(raw as Record<string, unknown>)) {
      if (!FORWARDED.has(id as ShortcutId) || !Array.isArray(chords)) continue
      for (const chord of chords) {
        const norm = typeof chord === 'string' ? normalizeChord(chord) : null
        if (norm && !next.has(norm)) next.set(norm, id as ShortcutId)
      }
    }
  }
  chordToAction = next
}

/** call once at startup */
export function initShortcutForwarding(): void {
  if (initialised) return
  initialised = true
  ipcMain.on(SHORTCUT_CHANNELS.sync, (_event, map: unknown) => rebuild(map))
}

/**
 * Attach to an editor view's webContents when the view is created.
 * `getShell` returns the shell window's webContents (looked up lazily: it may be recreated).
 */
export function forwardShortcuts(view: WebContents, getShell: () => WebContents | null): void {
  view.on('before-input-event', (event, input) => {
    if (input.type !== 'keyDown') return
    const chord = chordFromEvent(
      {
        key: input.key,
        code: input.code,
        ctrlKey: input.control,
        metaKey: input.meta,
        altKey: input.alt,
        shiftKey: input.shift,
      },
      process.platform === 'darwin',
    )
    if (!chord) return
    const action = chordToAction.get(chord)
    if (!action) return
    event.preventDefault()
    // held keys: only tab cycling repeats; "new tab" must not open a tab per auto-repeat
    if (input.isAutoRepeat && !CYCLING.has(action)) return
    const target = getShell()
    if (target && !target.isDestroyed()) target.send(SHORTCUT_CHANNELS.invoke, action)
  })
}
