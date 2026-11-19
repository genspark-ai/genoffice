/**
 * Native <input type="color"> never fires a change event when the user picks the exact
 * color the input already holds (e.g. adding a black stroke while the input's default is
 * already black, or applying the same custom text color to a second element). Call this
 * right before the picker opens: it flips the lowest bit of the blue channel, so any
 * confirmed selection — including the previously shown color — differs from the input's
 * value and fires change. The 1/255 difference is imperceptible in the picker preview.
 */
export function armColorInput(input: HTMLInputElement): void {
  const m = /^#([0-9a-fA-F]{6})$/.exec(input.value)
  if (!m) return
  const rgb = parseInt(m[1]!, 16) ^ 1
  input.value = `#${rgb.toString(16).padStart(6, '0')}`
}

/** Model color (#RRGGBB, optional AA) → <input type="color"> value; null when not a hex color. */
export function toPickerHex(color: string | undefined): string | null {
  if (!color) return null
  const m = /^#?([0-9a-fA-F]{6})/.exec(color)
  return m ? `#${m[1]!.toLowerCase()}` : null
}
