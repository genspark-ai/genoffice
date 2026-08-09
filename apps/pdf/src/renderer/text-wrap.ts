/** Greedy line breaking for paragraph (block) text edits: Latin breaks at word
    boundaries, CJK after any character, with basic kinsoku (no closing punctuation
    at a line start, no opening bracket at a line end). Measurement goes through
    canvas measureText in the same CSS family the editor previews with — renderer
    metrics are the project's source of truth for text width (hb wasm mismeasures
    AAT faces). */

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
  return out
}
