import { useEffect } from 'react'
import {
  SHORTCUTS_CHANGED_EVENT,
  chordFromEvent,
  getShortcutMap,
  isShortcutRecording,
  resolveShortcut,
  runShortcut,
} from './shortcuts'
import type { ShortcutMap } from './shortcuts'
import type { ShortcutsApi } from '../../shared/shortcuts-api'

declare global {
  interface Window {
    /** present in Electron (preload); absent in the browser mock */
    aiOfficeShortcuts?: ShortcutsApi
  }
}

/** dialogs that own the keyboard: shortcuts stay quiet underneath them */
const BLOCKING_OVERLAYS = '.set-overlay, .modal-overlay, .onb-overlay'

/**
 * One window-level keydown listener for every rebindable shortcut. Handlers
 * register through `onShortcut` (Home, TabBar, AccountEntry); this only maps a
 * key press to an action and offers it to them.
 *
 * Also accepts shortcuts forwarded by the main process (`aiOfficeShortcuts`),
 * which is how they keep working while an editor tab has keyboard focus.
 */
export function useGlobalShortcuts(): void {
  useEffect(() => {
    let map: ShortcutMap = getShortcutMap()
    // the main process matches chords on editor views, so it needs the live bindings
    const publish = (): void => window.aiOfficeShortcuts?.syncBindings(getShortcutMap())
    const refresh = (): void => {
      map = getShortcutMap()
      publish()
    }
    publish()
    window.addEventListener(SHORTCUTS_CHANGED_EVENT, refresh)

    const onKeyDown = (e: KeyboardEvent): void => {
      if (e.defaultPrevented || isShortcutRecording()) return
      if (document.querySelector(BLOCKING_OVERLAYS)) return
      const chord = chordFromEvent(e)
      if (!chord) return
      const id = resolveShortcut(map, chord)
      if (!id) return
      // auto-repeat only makes sense for cycling (Alt+Up/Down); never re-fire "new tab" etc.
      if (e.repeat && id !== 'previousTab' && id !== 'nextTab') {
        e.preventDefault()
        return
      }
      if (runShortcut(id)) {
        e.preventDefault()
        e.stopPropagation()
      }
    }
    window.addEventListener('keydown', onKeyDown, true)

    const offForwarded = window.aiOfficeShortcuts?.onInvoke((id) => {
      if (document.querySelector(BLOCKING_OVERLAYS)) return
      runShortcut(id)
    })

    return () => {
      window.removeEventListener(SHORTCUTS_CHANGED_EVENT, refresh)
      window.removeEventListener('keydown', onKeyDown, true)
      offForwarded?.()
    }
  }, [])
}
