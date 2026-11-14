type SaveKeyEvent = Pick<
  KeyboardEvent,
  'key' | 'code' | 'metaKey' | 'ctrlKey' | 'altKey' | 'getModifierState'
>

export function isSaveShortcut(event: SaveKeyEvent): boolean {
  if (!(event.metaKey || event.ctrlKey) || event.altKey || event.getModifierState('AltGraph')) {
    return false
  }
  if (/^[a-z]$/i.test(event.key)) return event.key.toLowerCase() === 's'
  return /^[\u0080-\u{10ffff}]$/u.test(event.key) && event.code === 'KeyS'
}
