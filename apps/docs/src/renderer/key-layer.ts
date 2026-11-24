import { useSyncExternalStore } from 'react'
import type { ActionCommandId } from '../shared/commands'
import type { KeyOverrides } from '../shared/keybindings'
import { docKeysFromXml, docKeysXml, layerDocKeys, type DocKey } from './doc-keymaps'
import { copyTextKey } from './editor/copy-text'
import { extendModeKey, extendModeOn } from './editor/extend-selection'
import {
  buildKeyTable,
  chordFromEvent,
  documentKeyCommand,
  keyVerdict,
  type KeyTable,
} from './keybindings'

type Runner = (command: ActionCommandId, payload?: string) => void

/** the machine's assignments, as Customize Keyboard stores them */
let machine: KeyOverrides = {}
/** the key maps of the open document, and the part they were read from */
let docKeys: readonly DocKey[] = []
let docKeysSource: string | undefined
/** what the keys answer to: the machine's assignments under the document's */
let overrides: KeyOverrides = {}
let table: KeyTable | null = null
let runner: Runner | null = null
let onDocKeysEdited: (() => void) | null = null
/** the menu bar may hold keys of a document; true at first, as a reloaded page cannot know what it sent before */
let menuHasDocKeys = true
const listeners = new Set<() => void>()

function rebuild(): void {
  const layered = docKeys.length > 0
  overrides = layered ? layerDocKeys(machine, docKeys) : machine
  table = Object.keys(overrides).length > 0 ? buildKeyTable(overrides) : null
  if (layered || menuHasDocKeys)
    window.desktop?.setDocumentKeyOverrides?.(layered ? overrides : null)
  menuHasDocKeys = layered
  for (const listener of listeners) listener()
}

function apply(next: KeyOverrides): void {
  machine = next
  rebuild()
}

function subscribe(listener: () => void): () => void {
  listeners.add(listener)
  return () => listeners.delete(listener)
}

/** the keyboard assignments in force (the machine's and the open document's), re-rendering when they change */
export function useKeyOverrides(): KeyOverrides {
  return useSyncExternalStore(subscribe, () => overrides)
}

/** the machine's own assignments, which Customize Keyboard edits under "This computer" */
export function useMachineKeyOverrides(): KeyOverrides {
  return useSyncExternalStore(subscribe, () => machine)
}

export function useDocumentKeys(): readonly DocKey[] {
  return useSyncExternalStore(subscribe, () => docKeys)
}

/** the key maps of a document just opened; `xml` is its Customizations part, if it has one */
export function loadDocumentKeys(xml: string | undefined): void {
  docKeysSource = xml
  docKeys = docKeysFromXml(xml)
  rebuild()
}

/** the user changed the document's key maps: they apply at once and are written with the next save */
export function saveDocumentKeys(next: readonly DocKey[]): void {
  docKeys = next
  rebuild()
  onDocKeysEdited?.()
}

export function setDocumentKeysEditedListener(listener: (() => void) | null): void {
  onDocKeysEdited = listener
}

/** the Customizations part to save; undefined when there is none to write */
export function documentKeysXml(): string | undefined {
  return docKeysXml(docKeysSource, docKeys) ?? undefined
}

/** false for a document whose Customizations part is not one this app can rewrite */
export function documentKeysEditable(): boolean {
  return docKeysSource === undefined || docKeysXml(docKeysSource, []) !== null
}

/** the assignments as they stand now, for a key handler outside React */
export function currentKeyOverrides(): KeyOverrides {
  return overrides
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

/** commands that act on the caret of the editor the key was pressed in */
const CARET_COMMANDS: readonly ActionCommandId[] = [
  'caret-move',
  'delete-word',
  'edit-text',
  'edit-clear-contents',
]

function onKeyDown(event: KeyboardEvent): void {
  if (event.isComposing) return
  const target = event.target instanceof Element ? event.target : null
  if (target && ownsItsKeys(target)) return
  if (copyTextKey(event)) {
    event.preventDefault()
    event.stopImmediatePropagation()
    return
  }
  const chord = chordFromEvent(event)
  const verdict = table && chord ? keyVerdict(table, chord) : 'pass'
  if (verdict === 'pass' && extendModeOn()) {
    const move = chord ? documentKeyCommand(overrides, chord) : null
    const caretMove = move?.cmd === 'caret-move' ? (move.payload ?? null) : null
    if (extendModeKey(event, caretMove, (payload) => runner?.('caret-move', payload))) {
      event.preventDefault()
      event.stopImmediatePropagation()
      return
    }
  }
  if (!chord) return
  let run = verdict === 'pass' || verdict === 'swallow' ? null : verdict.run
  if (verdict === 'pass') {
    run = documentKeyCommand(overrides, chord)
    if (!run) return
    if (CARET_COMMANDS.includes(run.cmd) && !target?.closest('.ProseMirror')) return
  }
  event.preventDefault()
  event.stopImmediatePropagation()
  if (run) runner?.(run.cmd, run.payload)
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
