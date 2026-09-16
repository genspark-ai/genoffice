/**
 * Length fields accept a unit suffix ("2.54cm", "1in", "10mm", "12pt",
 * "96px", "914400emu") anywhere the model speaks EMU. Normalized in place
 * before validation, so every op keeps seeing integers; a bare number is
 * still EMU. px is CSS px (96 per inch), the unit of the page-spec pipeline.
 * Length fields are the `x/y/cx/cy/dx/dy` geometry keys, every `…Emu` key
 * (widthEmu, wEmu, hEmu, colWidthsEmu, bulletHangEmu, ...) and inset sides.
 */
const EMU_PER: Record<string, number> = {
  emu: 1,
  in: 914400,
  cm: 360000,
  mm: 36000,
  pt: 12700,
  pc: 152400,
  px: 9525,
}

const LENGTH = /^\s*(-?\d+(?:\.\d+)?)\s*(emu|in|cm|mm|pt|pc|px)\s*$/i

/** Bare geometry keys; every other length field is named `…Emu`, and insets are `{l, t, r, b}`. */
const GEOMETRY_KEYS = new Set(['x', 'y', 'cx', 'cy', 'dx', 'dy'])
const INSET_KEYS = new Set(['l', 't', 'r', 'b'])

function isLengthKey(key: string, parent?: string): boolean {
  if (key.endsWith('Emu') || GEOMETRY_KEYS.has(key)) return true
  return INSET_KEYS.has(key) && (parent === 'insets' || parent === 'insetsEmu')
}

/** "2.54cm" → 2286000; undefined when the string is not a length. */
export function parseLength(value: string): number | undefined {
  const m = LENGTH.exec(value)
  if (!m) return undefined
  return Math.round(Number(m[1]) * EMU_PER[m[2]!.toLowerCase()]!)
}

export function normalizeLengthUnits(value: unknown, key?: string, parent?: string): unknown {
  if (typeof value === 'string') {
    if (key === undefined || !isLengthKey(key, parent)) return value
    return parseLength(value) ?? value
  }
  if (Array.isArray(value)) {
    for (let i = 0; i < value.length; i++) value[i] = normalizeLengthUnits(value[i], key, parent)
    return value
  }
  // picture/media bytes ride along as Uint8Array; Object.keys on those is one key per byte
  if (value && typeof value === 'object' && !ArrayBuffer.isView(value)) {
    const record = value as Record<string, unknown>
    for (const k of Object.keys(record)) record[k] = normalizeLengthUnits(record[k], k, key)
    return value
  }
  return value
}
