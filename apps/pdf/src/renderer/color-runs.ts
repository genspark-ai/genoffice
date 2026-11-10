/** Selection-level text-edit styles, renderer side: a draft keeps one encoded style
    key per code unit of its value ('' = the draft's base style); commits turn that
    into the compact [start,end) ranges TextEditInput.styleRuns carries. The run
    helpers below treat keys as opaque strings — equality is all they need — so they
    serve plain colors and full styles alike. Mapping between text forms (draft value
    ↔ wrapped/committed newText ↔ reopened blockSource) aligns on non-whitespace
    chars — wrapping and line joining only rearrange whitespace. */

export interface ColorRun {
  start: number
  end: number
  /** Encoded style key (see encodeStyle); historically a bare CSS hex like '#d32f2f' */
  color: string
}

/** Selection-level style overrides of one char; every field absent = inherit the
    draft-level value (which in turn inherits the document's original run). */
export interface CharStyle {
  /** CSS hex like '#d32f2f' */
  color?: string
  /** EDIT_FONTS id */
  font?: string
  /** Font size in PDF pt */
  size?: number
  /** Explicit on/off; absent = inherit the draft toggle */
  bold?: boolean
  italic?: boolean
}

/** Canonical key: '' = base; else 'color|font|size|bold|italic' with '' fields
    inheriting and bold/italic '1'/'0' explicit on/off. A bare-color style encodes
    with trailing '|'s so distinct styles never collide. Fields are validated
    with the same gates decodeStyle/patchStyle use, so a crafted value passed
    straight to encode (bypassing patchStyle) can never inject a '|' field
    separator or a non-decimal size into a stored key. */
export function encodeStyle(s: CharStyle): string {
  const tri = (v: boolean | undefined) => (v === undefined ? '' : v ? '1' : '0')
  const color = typeof s.color === 'string' && /^#[0-9a-f]{6}$/i.test(s.color) ? s.color : ''
  const font = typeof s.font === 'string' && isStyleFontId(s.font) ? s.font : ''
  const size =
    typeof s.size === 'number' &&
    Number.isFinite(s.size) &&
    s.size > 0 &&
    s.size <= MAX_STYLE_FONT_SIZE
      ? String(s.size)
      : ''
  const key = [color, font, size, tri(s.bold), tri(s.italic)].join('|')
  return key === '||||' ? '' : key
}

/**
 * Font ids are EDIT_FONTS entries (see apps/pdf/src/shared/ipc.ts: arial,
 * times, courier) but the check stays charset-based rather than a closed
 * allowlist so a future id degrades to inherit instead of needing a lockstep
 * update here. The charset excludes '|' so a crafted id can never inject an
 * extra field and shift size/bold/italic on the next decode.
 */
export function isStyleFontId(font: string): boolean {
  return /^[A-Za-z0-9_-]{1,32}$/.test(font)
}

/** Largest font size the layout path can handle (pt); larger values blow up
    Math.round(size * 100) to Infinity and break sz attributes. */
export const MAX_STYLE_FONT_SIZE = 1000

/** Decimal point sizes only: rejects hex (0x10), exponents (1e2), whitespace
    padding and anything Number() would coerce but Word would never emit. */
export function parseStyleFontSize(raw: string): number | undefined {
  const text = raw.trim()
  if (!/^\d+(\.\d+)?$/.test(text)) return undefined
  const n = Number(text)
  if (!Number.isFinite(n) || n <= 0 || n > MAX_STYLE_FONT_SIZE) return undefined
  return n
}

export function decodeStyle(key: string): CharStyle {
  if (!key) return {}
  const [color = '', font = '', size = '', bold = '', italic = ''] = key.split('|')
  const out: CharStyle = {}
  if (/^#[0-9a-f]{6}$/i.test(color)) out.color = color
  if (font && isStyleFontId(font)) out.font = font
  if (size) {
    const n = parseStyleFontSize(size)
    if (n !== undefined) out.size = n
  }
  if (bold === '1') out.bold = true
  else if (bold === '0') out.bold = false
  if (italic === '1') out.italic = true
  else if (italic === '0') out.italic = false
  return out
}

/** Merge a partial style into an encoded key: undefined fields keep their value,
    null clears the field back to inherit. Invalid values (non-hex color,
    off-charset font, out-of-range size, non-boolean toggle) are ignored so a
    crafted patch can never inject field separators or Infinity sizes. */
export function patchStyle(
  key: string,
  patch: { [K in keyof CharStyle]?: CharStyle[K] | null },
): string {
  const s = decodeStyle(key)
  for (const k of ['color', 'font', 'size', 'bold', 'italic'] as const) {
    const v = patch[k]
    if (v === undefined) continue
    if (v === null) {
      delete s[k]
      continue
    }
    if (k === 'color') {
      if (typeof v !== 'string' || !/^#[0-9a-f]{6}$/i.test(v)) continue
      s.color = v
    } else if (k === 'font') {
      if (typeof v !== 'string' || !isStyleFontId(v)) continue
      s.font = v
    } else if (k === 'size') {
      if (typeof v !== 'number' || !Number.isFinite(v) || v <= 0 || v > MAX_STYLE_FONT_SIZE)
        continue
      s.size = v
    } else if (k === 'bold' || k === 'italic') {
      if (typeof v !== 'boolean') continue
      s[k] = v
    }
  }
  return encodeStyle(s)
}

/** Carry per-char colors across a textarea value change: the edit is localized to the
    span between the old/new values' common prefix and suffix (typing, paste, cut, IME
    commits and native undo all arrive this way). Inserted chars inherit the color at
    the left boundary — how run styling behaves in every rich editor. */
export function spliceCharColors(
  oldValue: string,
  colors: readonly string[],
  newValue: string,
): string[] {
  let p = 0
  const maxP = Math.min(oldValue.length, newValue.length)
  while (p < maxP && oldValue[p] === newValue[p]) p++
  let s = 0
  const maxS = maxP - p
  while (s < maxS && oldValue[oldValue.length - 1 - s] === newValue[newValue.length - 1 - s]) s++
  const inherited = p > 0 ? (colors[p - 1] ?? '') : ''
  const insertedLen = newValue.length - p - s
  return [
    ...colors.slice(0, p),
    ...Array<string>(insertedLen).fill(inherited),
    ...colors.slice(oldValue.length - s),
  ]
}

/** Map per-char colors from one text form to another that differs only in whitespace
    layout (wrapText / joinBlockLines): non-whitespace chars pair up in order,
    whitespace inherits the preceding char's color. */
export function mapCharColors(
  srcText: string,
  srcColors: readonly string[],
  dstText: string,
): string[] {
  const isWs = (ch: string) => /\s/.test(ch)
  const out: string[] = []
  let si = 0
  for (let di = 0; di < dstText.length; di++) {
    if (isWs(dstText[di]!)) {
      out.push(out[di - 1] ?? '')
      continue
    }
    while (si < srcText.length && isWs(srcText[si]!)) si++
    out.push(si < srcText.length ? (srcColors[si] ?? '') : '')
    si++
  }
  return out
}

/** Compact non-base ranges; adjacent equal colors merge */
export function colorsToRuns(colors: readonly string[]): ColorRun[] {
  const runs: ColorRun[] = []
  for (let i = 0; i < colors.length; i++) {
    const c = colors[i]!
    if (!c) continue
    const last = runs[runs.length - 1]
    if (last && last.end === i && last.color === c) last.end = i + 1
    else runs.push({ start: i, end: i + 1, color: c })
  }
  return runs
}

export function runsToColors(len: number, runs: readonly ColorRun[]): string[] {
  const out = Array<string>(len).fill('')
  for (const r of runs) {
    for (let i = Math.max(0, r.start); i < Math.min(len, r.end); i++) out[i] = r.color
  }
  return out
}

export const colorRunsEqual = (
  a: readonly ColorRun[] | undefined,
  b: readonly ColorRun[] | undefined,
): boolean => {
  const an = a ?? []
  const bn = b ?? []
  return (
    an.length === bn.length &&
    an.every((r, i) => r.start === bn[i]!.start && r.end === bn[i]!.end && r.color === bn[i]!.color)
  )
}

/** Consecutive same-color spans over the full text (base color spans included, color '')
    — what the editor backdrop and the pending-edit preview render */
export function colorSegments(
  text: string,
  colors: readonly string[],
): { text: string; color: string }[] {
  const segs: { text: string; color: string }[] = []
  for (let i = 0; i < text.length; i++) {
    const c = colors[i] ?? ''
    const last = segs[segs.length - 1]
    if (last && last.color === c) last.text += text[i]!
    else segs.push({ text: text[i]!, color: c })
  }
  return segs.length ? segs : [{ text, color: '' }]
}
