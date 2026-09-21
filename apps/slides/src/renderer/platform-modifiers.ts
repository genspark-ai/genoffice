/**
 * PowerPoint's canvas modifiers per platform: toggle-in-selection is Ctrl+click on
 * Windows/Linux and Cmd+click on mac (Ctrl+click is the mac system right-click);
 * duplicate-at-drop is Ctrl+drag on Windows/Linux and Option+drag on mac. Shift adds
 * on both.
 */
export interface ModifierKeys {
  shiftKey: boolean
  ctrlKey: boolean
  metaKey: boolean
  altKey: boolean
}

export const isMac = /mac/i.test(typeof navigator === 'undefined' ? '' : navigator.platform)

export function isToggleModifier(e: ModifierKeys, mac = isMac): boolean {
  return e.shiftKey || (mac ? e.metaKey : e.ctrlKey)
}

export function isDuplicateDragModifier(e: ModifierKeys, mac = isMac): boolean {
  return mac ? e.altKey : e.ctrlKey
}

/** Plain click replaces the selection; a toggle-modifier click adds or removes the element. */
export function nextSelection(prev: string[], id: string | null, additive: boolean): string[] {
  if (id == null) return []
  if (additive) return prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]
  return [id]
}
