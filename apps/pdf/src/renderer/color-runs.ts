/** Selection-level text-edit colors, renderer side: a draft keeps one color per code
    unit of its value ('' = the draft's base color); commits turn that into the compact
    [start,end) ranges TextEditInput.colorRuns carries. Mapping between text forms
    (draft value ↔ wrapped/committed newText ↔ reopened blockSource) aligns on
    non-whitespace chars — wrapping and line joining only rearrange whitespace. */

export interface ColorRun {
  start: number
  end: number
  /** CSS hex like '#d32f2f' */
  color: string
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
