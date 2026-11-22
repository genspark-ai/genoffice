import { useSyncExternalStore } from 'react'
import type { ActionCommandId } from '../shared/commands'
import type { KeyOverrides } from '../shared/keybindings'
import { buildKeyTable, chordFromEvent, keyVerdict, type KeyTable } from './keybindings'

type Runner = (command: ActionCommandId, payload?: string) => void

let overrides: KeyOverrides = {}
let table: KeyTable | null = null
let runner: Runner | null = null
const listeners = new Set<() => void>()

function apply(next: KeyOverrides): void {
  overrides = next
  table = Object.keys(next).length > 0 ? buildKeyTable(next) : null
  for (const listener of listeners) listener()
}

function subscribe(listener: () => void): () => void {
  listeners.add(listener)
  return () => listeners.delete(listener)
}

/** the user's keyboard assignments, re-rendering when any window changes them */
export function useKeyOverrides(): KeyOverrides {
  return useSyncExternalStore(subscribe, () => overrides)
}

export function saveKeyOverrides(next: KeyOverrides): void {
  apply(next)
  void window.desktop?.setKeyOverrides?.(next)
}

export function setKeyLayerRunner(run: Runner | null): void {
  runner = run
}

/**
 * A dialog (the shortcut box of Customize Keyboard among them) and a text
 * field outside the document keep their keys: a comment box submits on the
 * chord Page Break has in the document.
 */
export function ownsItsKeys(target: Element): boolean {
  if (target.closest('.modal-backdrop')) return true
  const field = target.closest(
    'input, textarea, select, [contenteditable]:not([contenteditable="false"])',
  )
  return !!field && !field.closest('.ProseMirror, td[contenteditable]')
}

function onKeyDown(event: KeyboardEvent): void {
  if (!table || event.isComposing) return
  if (event.target instanceof Element && ownsItsKeys(event.target)) return
  const chord = chordFromEvent(event)
  if (!chord) return
  const verdict = keyVerdict(table, chord)
  if (verdict === 'pass') return
  event.preventDefault()
  event.stopImmediatePropagation()
  if (verdict !== 'swallow') runner?.(verdict.run.cmd, verdict.run.payload)
}

/**
 * Installed once, before anything else listens: a customized chord has to be
 * decided ahead of the editor keymaps and the window's built-in handlers.
 */
export function installKeyLayer(): void {
  window.addEventListener('keydown', onKeyDown, true)
  void window.desktop
    ?.getKeyOverrides?.()
    .then(apply)
    .catch(() => undefined)
  window.desktop?.onKeyOverridesChanged?.(apply)
}
