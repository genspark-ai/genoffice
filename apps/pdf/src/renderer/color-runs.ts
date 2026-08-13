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
    with trailing '|'s so distinct styles never collide. */
export function encodeStyle(s: CharStyle): string {
  const tri = (v: boolean | undefined) => (v === undefined ? '' : v ? '1' : '0')
  const key = [
    s.color ?? '',
    s.font ?? '',
    s.size !== undefined ? String(s.size) : '',
    tri(s.bold),
    tri(s.italic),
  ].join('|')
  return key === '||||' ? '' : key
}

export function decodeStyle(key: string): CharStyle {
  if (!key) return {}
  const [color = '', font = '', size = '', bold = '', italic = ''] = key.split('|')
  const out: CharStyle = {}
  if (color) out.color = color
  if (font) out.font = font
  if (size) out.size = Number(size)
  if (bold) out.bold = bold === '1'
  if (italic) out.italic = italic === '1'
  return out
}

/** Merge a partial style into an encoded key: undefined fields keep their value,
    null clears the field back to inherit. */
export function patchStyle(
  key: string,
  patch: { [K in keyof CharStyle]?: CharStyle[K] | null },
): string {
  const s = decodeStyle(key)
  for (const k of ['color', 'font', 'size', 'bold', 'italic'] as const) {
    const v = patch[k]
    if (v === undefined) continue
    if (v === null) delete s[k]
    else (s as Record<string, unknown>)[k] = v
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
