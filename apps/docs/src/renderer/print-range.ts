/**
 * Parse a Word-style print range ("1,3,5-8") into sorted 0-based page indices.
 * Accepts CJK/ASCII separators and en-dashes; returns null on any invalid or
 * out-of-bounds part (same semantics as the slides print dialog).
 */
export function parsePrintRange(text: string, max: number): number[] | null {
  const out = new Set<number>()
  const parts = text.split(/[,，、;；\s]+/).filter(Boolean)
  if (parts.length === 0) return null
  for (const part of parts) {
    const m = /^(\d+)\s*[-–]\s*(\d+)$|^(\d+)$/.exec(part)
    if (!m) return null
    const a = Number(m[1] ?? m[3])
    const b = Number(m[2] ?? m[3])
    if (a < 1 || b > max || a > b) return null
    for (let i = a; i <= b; i++) out.add(i - 1)
  }
  return [...out].sort((x, y) => x - y)
}
