/** Greedy line breaking for paragraph (block) text edits: Latin breaks at word
    boundaries, CJK after any character, with basic kinsoku (no closing punctuation
    at a line start, no opening bracket at a line end). Measurement goes through
    canvas measureText in the same CSS family the editor previews with — renderer
    metrics are the project's source of truth for text width (hb wasm mismeasures
    AAT faces). */

import { foldRadicals } from '../shared/radicals'

const CJK = /[⺀-〿぀-ヿㇰ-ㇿ㐀-䶿一-鿿豈-﫿＀-￯가-힯]/

/** Must not start a line: closing brackets/quotes, CJK and Latin sentence punctuation */
const NO_LINE_START = new Set('，。、；：！？）】」』〉》〕…‥·—～ヽヾゝゞ々ー%％℃,.;:!?)]}"\'’”°')
/** Must not end a line: opening brackets/quotes */
const NO_LINE_END = new Set('（【「『〈《〔([{“‘')

/** Breakable units: whitespace runs, single CJK chars, Latin words (with trailing
    punctuation attached so kinsoku can push it as one piece) */
function tokenize(text: string): string[] {
  const units: string[] = []
  let word = ''
  const flush = () => {
    if (word) units.push(word)
    word = ''
  }
  for (const ch of text) {
    if (/\s/.test(ch)) {
      flush()
      if (units.length > 0 && units[units.length - 1] === ' ') continue
      units.push(' ')
    } else if (CJK.test(ch) || NO_LINE_START.has(ch) || NO_LINE_END.has(ch)) {
      flush()
      units.push(ch)
    } else {
      word += ch
    }
  }
  flush()
  return units
}

let ctx: CanvasRenderingContext2D | null = null
const widthCache = new Map<string, number>()

/** Text width in PDF pt for the given effective font size and CSS family;
    cssStyle is a CSS font-shorthand prefix like 'bold' / 'italic' / 'italic bold' */
export function measurePt(
  text: string,
  fontSizePt: number,
  cssFamily: string,
  cssStyle = '',
): number {
  if (!ctx) ctx = document.createElement('canvas').getContext('2d')
  if (!ctx) return text.length * fontSizePt * 0.5
  const font = `${cssStyle} 100px ${cssFamily}`.trim()
  if (ctx.font !== font) ctx.font = font
  let w = 0
  for (const ch of text) {
    const key = `${cssStyle}\u0000${cssFamily}\u0000${ch}`
    let cw = widthCache.get(key)
    if (cw === undefined) {
      cw = ctx.measureText(ch).width
      widthCache.set(key, cw)
    }
    w += cw
  }
  return (w / 100) * fontSizePt
}

/**
 * Wrap one paragraph (no '\n' inside) to the given width. Returns at least one line.
 * A unit longer than the whole width is hard-broken by character.
 */
export function wrapText(
  text: string,
  widthPt: number,
  fontSizePt: number,
  cssFamily: string,
  cssStyle = '',
): string[] {
  const w = (s: string) => measurePt(s, fontSizePt, cssFamily, cssStyle)
  const units = tokenize(text.trim())
  const lines: string[] = []
  let line = ''
  const push = () => {
    if (line.trim()) lines.push(line.trim())
    line = ''
  }
  for (let i = 0; i < units.length; i++) {
    let u = units[i]!
    if (u === ' ' && !line) continue
    if (line && w(line + u) > widthPt && u !== ' ') {
      if (NO_LINE_START.has(u)) {
        // Pull the last unit down with the punctuation instead of starting a line with it
        const m = /^(.*?)(\S)$/.exec(line)
        if (m && m[1]!.trim()) {
          line = m[1]!
          push()
          line = m[2]!
        } else push()
      } else {
        // Never leave an opening bracket dangling at the line end
        const last = line.trimEnd()
        if (last && NO_LINE_END.has(last[last.length - 1]!)) {
          line = last.slice(0, -1)
          push()
          line = last[last.length - 1]!
        } else push()
      }
    }
    // Hard-break a single unit wider than the block (URLs, very narrow blocks)
    while (w(u) > widthPt && u.length > 1) {
      let k = 1
      while (k < u.length && w(line + u.slice(0, k + 1)) <= widthPt) k++
      line += u.slice(0, k)
      push()
      u = u.slice(k)
    }
    line += u
  }
  push()
  return lines.length > 0 ? lines : ['']
}

/** pdf.js sometimes extracts CJK through a font cmap that yields radical-block
    codepoints (U+2E80-2FDF, e.g. U+2F83 for U+81EA, U+2EE6 for U+9E1F) where the page
    really draws unified ideographs — and PDFium extracts those. Fold each radical to
    its unified equivalent (visually identical) so block drafts edit and write back
    real ideographs: retained radical variants between two folded edits would
    otherwise reach the rebuild font as undrawable characters. NFKC alone is not
    enough — the Radicals Supplement (U+2E80-2EFF) has no decompositions — so this
    goes through the shared RADICAL_EQUIV table first. */
export function unifyRadicals(s: string): string {
  // Gate covers everything foldRadicals maps (radical blocks AND CJK Strokes
  // U+31C0-31E3) — the engine folds strokes too, so both sides must agree
  return /[\u2e80-\u2fdf\u31c0-\u31e3]/.test(s)
    ? foldRadicals(s).replace(/[\u2e80-\u2fdf]/g, (ch) => ch.normalize('NFKC'))
    : s
}

/** Non-whitespace code units of `s` (radical-folded) with their original indices —
    splice matching must ignore synthesized/joined spaces and radical variants.
    Folds are pushed unit by unit (an astral fold becomes two surrogate entries
    sharing one original index), so an unfolded needle aligns with an
    already-folded haystack that carries the surrogates as separate units. */
function nonSpaceMap(s: string): { chars: string[]; idx: number[] } {
  const chars: string[] = []
  const idx: number[] = []
  for (let i = 0; i < s.length; i++) {
    if (/\s/.test(s[i]!)) continue
    const folded = unifyRadicals(s[i]!)
    for (let k = 0; k < folded.length; k++) {
      chars.push(folded[k]!)
      idx.push(i)
    }
  }
  return { chars, idx }
}

const indexOfSub = (hay: string[], needle: string[], from: number): number => {
  if (needle.length === 0) return -1
  outer: for (let i = Math.max(0, from); i + needle.length <= hay.length; i++) {
    for (let j = 0; j < needle.length; j++) if (hay[i + j] !== needle[j]) continue outer
    return i
  }
  return -1
}

/** Fold pending line edits into a block's joined paragraph text: each edit's
    oldText is located whitespace-insensitively (line joins differ in synthesized
    spaces) and replaced by its newText as committed. `hint` is the edit's rough
    non-space offset within the block (from the edited line's position), used to
    prefer the right occurrence when a paragraph repeats itself. Returns null when
    any oldText cannot be found or two edits overlap — the caller then keeps the
    unfolded behavior. When `outRanges` is given, it receives the [start,end)
    range each edit's (radical-folded) newText occupies inside the returned
    string, in input-edit order — callers use this to carry the edits'
    selection styles onto the folded paragraph. */
export function spliceBlockText(
  blockText: string,
  edits: { oldText: string; newText: string; hint?: number }[],
  outRanges?: [number, number][],
): string | null {
  const hay = nonSpaceMap(blockText)
  const ranges: { start: number; end: number; newText: string; editIdx: number }[] = []
  for (const [editIdx, e] of edits.entries()) {
    const needle = nonSpaceMap(e.oldText).chars
    // From the hinted line start first; a crossing-span oldText that begins on the
    // previous line (or a stale hint) falls back to the first occurrence anywhere
    let at = indexOfSub(hay.chars, needle, e.hint ?? 0)
    if (at < 0) at = indexOfSub(hay.chars, needle, 0)
    if (at < 0) return null
    ranges.push({
      start: hay.idx[at]!,
      end: hay.idx[at + needle.length - 1]! + 1,
      newText: unifyRadicals(e.newText),
      editIdx,
    })
  }
  ranges.sort((a, b) => a.start - b.start)
  for (let i = 1; i < ranges.length; i++) {
    if (ranges[i]!.start < ranges[i - 1]!.end) return null
  }
  let out = ''
  let pos = 0
  for (const r of ranges) {
    out += blockText.slice(pos, r.start)
    if (outRanges) outRanges[r.editIdx] = [out.length, out.length + r.newText.length]
    out += r.newText
    pos = r.end
  }
  return out + blockText.slice(pos)
}

/** Map a [start,end) code-unit range of one visual line's raw text to the offsets of
    the same characters inside the block's joined paragraph text. Whitespace differs
    between the two (line joins synthesize/drop spaces, the DOM line synthesizes gap
    spaces), so alignment goes through the folded non-space sequences, like
    spliceBlockText. A range covering only whitespace collapses to a caret. null =
    the line cannot be located inside the block text. */
export function mapLineRangeToBlock(
  blockText: string,
  lineText: string,
  start: number,
  end: number,
): [number, number] | null {
  const hay = nonSpaceMap(blockText)
  const needle = nonSpaceMap(lineText)
  if (needle.chars.length === 0) return null
  const at = indexOfSub(hay.chars, needle.chars, 0)
  if (at < 0) return null
  // Non-space ranks of the line covered by [start,end)
  let k1 = 0
  while (k1 < needle.idx.length && needle.idx[k1]! < start) k1++
  let k2 = k1
  while (k2 < needle.idx.length && needle.idx[k2]! < end) k2++
  if (k1 >= k2) {
    const caret =
      k1 < needle.idx.length ? hay.idx[at + k1]! : hay.idx[at + needle.chars.length - 1]! + 1
    return [caret, caret]
  }
  return [hay.idx[at + k1]!, hay.idx[at + k2 - 1]! + 1]
}

/** Join a block's extracted lines back into one logical paragraph string: a space at
    Latin boundaries, nothing between CJK (their hard breaks carry no space) */
export function joinBlockLines(lines: string[]): string {
  let out = ''
  for (const line of lines) {
    const t = line.trim()
    if (!t) continue
    if (out) {
      const a = out[out.length - 1]!
      const b = t[0]!
      out += CJK.test(a) && CJK.test(b) ? '' : ' '
    }
    out += t
  }
  return unifyRadicals(out)
}
