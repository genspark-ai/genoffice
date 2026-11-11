/** MS-PPT StyleTextPropAtom (RT_StyleTextPropAtom, 0x0FA1) run tables.
 *
 *  Layout (MS-PPT 2.9): zero or more TextPFRun records covering the text's
 *  characters, then zero or more TextCFRun records covering them again. Each
 *  run opens with count (u32); a TextPFRun also has indentLevel (u16) and a
 *  TextPFException, a TextCFRun a TextCFException. Both exceptions begin with
 *  a 4-byte mask whose set bits say which optional fields follow, in the
 *  structure's fixed order. Unknown/variable-size fields stop the parse
 *  cleanly rather than desynchronising the reader. */

export interface TextCharFormat {
  /** characters this run covers, in UTF-16 code units of the source text */
  count: number
  bold: boolean
  italic: boolean
  underline: boolean
  /** points, 1..4000, or null when the run does not set a size */
  sizePt: number | null
  /** RRGGBB, or null for scheme/indexed colors */
  color: string | null
}

export interface TextParagraphFormat {
  count: number
  level: number
  /** DrawingML alignment name, or null */
  align: 'l' | 'ctr' | 'r' | 'just' | null
}

export interface TextStyle {
  paragraphs: TextParagraphFormat[]
  characters: TextCharFormat[]
}

/* CFMasks bits (MS-PPT 2.9.15) */
const CF_BOLD = 1 << 0
const CF_ITALIC = 1 << 1
const CF_UNDERLINE = 1 << 2
const CF_SHADOW = 1 << 4
const CF_FEHINT = 1 << 5
const CF_KUMI = 1 << 7
const CF_EMBOSS = 1 << 9
const CF_HAS_STYLE = 0xf << 10
const CF_TYPEFACE = 1 << 16
const CF_SIZE = 1 << 17
const CF_COLOR = 1 << 18
const CF_POSITION = 1 << 19
const CF_OLD_EA = 1 << 21
const CF_ANSI = 1 << 22
const CF_SYMBOL = 1 << 23

/* PFMasks bits (MS-PPT 2.9.21) */
const PF_BULLET_FLAGS = 0b1111 // hasBullet / bulletHasFont / bulletHasColor / bulletHasSize
const PF_BULLET_FONT = 1 << 4
const PF_BULLET_COLOR = 1 << 5
const PF_BULLET_SIZE = 1 << 6
const PF_BULLET_CHAR = 1 << 7
const PF_LEFT_MARGIN = 1 << 8
const PF_INDENT = 1 << 10
const PF_ALIGN = 1 << 11
const PF_LINE_SPACING = 1 << 12
const PF_SPACE_BEFORE = 1 << 13
const PF_SPACE_AFTER = 1 << 14
const PF_DEFAULT_TAB = 1 << 15
const PF_FONT_ALIGN = 1 << 16
const PF_WRAP_FLAGS = 0b111 << 17
const PF_TAB_STOPS = 1 << 20
const PF_TEXT_DIRECTION = 1 << 21

const MAX_RUNS = 100_000

const ALIGNMENTS: Array<TextParagraphFormat['align']> = ['l', 'ctr', 'r', 'just']

/** ColorIndexStruct: red, green, blue, palette index. Index 0 is an RGB value. */
function colorOf(view: DataView, at: number): string | null {
  if (at + 4 > view.byteLength) return null
  const r = view.getUint8(at)
  const g = view.getUint8(at + 1)
  const b = view.getUint8(at + 2)
  const index = view.getUint8(at + 3)
  if (index !== 0 && r === 0 && g === 0 && b === 0) return null
  return [r, g, b]
    .map((c) => c.toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase()
}

/**
 * Parse the run tables. `charCount` is the text's length in UTF-16 code units
 * (the atom's recInstance normally carries it). A malformed tail is dropped,
 * not thrown: a hostile deck must still convert with the text it can read.
 */
export function parseStyleTextProp(payload: Uint8Array, charCount: number): TextStyle {
  const view = new DataView(payload.buffer, payload.byteOffset, payload.byteLength)
  const paragraphs: TextParagraphFormat[] = []
  const characters: TextCharFormat[] = []
  let at = 0
  const need = (bytes: number): boolean => at + bytes <= view.byteLength

  // TextPFRun records, until their counts cover the text
  let covered = 0
  while (covered < charCount && need(10) && paragraphs.length < MAX_RUNS) {
    const count = view.getUint32(at, true)
    at += 4
    const level = view.getUint16(at, true)
    at += 2
    if (!need(4)) break
    const mask = view.getUint32(at, true)
    at += 4
    const has = (bit: number): boolean => (mask & bit) !== 0
    // TextPFException field order; tabStops is variable-sized, so its presence
    // ends the parse instead of guessing the length.
    const variable = has(PF_TAB_STOPS)
    let align: TextParagraphFormat['align'] = null
    const step = (bytes: number, read?: (offset: number) => void): void => {
      if (at + bytes > view.byteLength) return
      read?.(at)
      at += bytes
    }
    if (has(PF_BULLET_FLAGS)) step(2)
    if (has(PF_BULLET_CHAR)) step(2)
    if (has(PF_BULLET_FONT)) step(2)
    if (has(PF_BULLET_SIZE)) step(2)
    if (has(PF_BULLET_COLOR)) step(4)
    if (has(PF_ALIGN))
      step(2, (o) => {
        align = ALIGNMENTS[view.getUint16(o, true)] ?? null
      })
    if (has(PF_LINE_SPACING)) step(2)
    if (has(PF_SPACE_BEFORE)) step(2)
    if (has(PF_SPACE_AFTER)) step(2)
    if (has(PF_LEFT_MARGIN)) step(2)
    if (has(PF_INDENT)) step(2)
    if (has(PF_DEFAULT_TAB)) step(2)
    if (has(PF_FONT_ALIGN)) step(2)
    if (has(PF_WRAP_FLAGS)) step(2)
    if (has(PF_TEXT_DIRECTION)) step(2)
    paragraphs.push({ count, level, align })
    covered += count
    if (count === 0 || variable) break
  }

  // TextCFRun records follow the paragraph runs
  covered = 0
  while (covered < charCount && need(8) && characters.length < MAX_RUNS) {
    const count = view.getUint32(at, true)
    at += 4
    if (!need(4)) break
    const mask = view.getUint32(at, true)
    at += 4
    const has = (bit: number): boolean => (mask & bit) !== 0
    let bold = false
    let italic = false
    let underline = false
    let sizePt: number | null = null
    let color: string | null = null
    const step = (bytes: number, read?: (offset: number) => void): void => {
      if (at + bytes > view.byteLength) return
      read?.(at)
      at += bytes
    }
    // TextCFException field order (MS-PPT 2.9.14)
    if (
      has(
        CF_BOLD |
          CF_ITALIC |
          CF_UNDERLINE |
          CF_SHADOW |
          CF_FEHINT |
          CF_KUMI |
          CF_EMBOSS |
          CF_HAS_STYLE,
      )
    ) {
      step(2, (o) => {
        const style = view.getUint16(o, true)
        bold = (style & 0x01) !== 0
        italic = (style & 0x02) !== 0
        underline = (style & 0x04) !== 0
      })
    }
    if (has(CF_TYPEFACE)) step(2)
    if (has(CF_OLD_EA)) step(2)
    if (has(CF_ANSI)) step(2)
    if (has(CF_SYMBOL)) step(2)
    if (has(CF_SIZE))
      step(2, (o) => {
        const size = view.getUint16(o, true)
        if (size >= 1 && size <= 4000) sizePt = size
      })
    if (has(CF_COLOR))
      step(4, (o) => {
        color = colorOf(view, o)
      })
    if (has(CF_POSITION)) step(2)
    // pp10ext / newEATypeface / csTypeface / pp11ext MUST be zero
    characters.push({ count, bold, italic, underline, sizePt, color })
    covered += count
    if (count === 0) break
  }

  return { paragraphs, characters }
}

/** Split normalized text into formatted runs, preserving the source order. */
export function applyCharacterRuns(
  text: string,
  runs: readonly TextCharFormat[],
): Array<{ text: string; format: TextCharFormat | null }> {
  if (runs.length === 0) return [{ text, format: null }]
  const out: Array<{ text: string; format: TextCharFormat | null }> = []
  const units = [...text]
  let at = 0
  for (const run of runs) {
    const take = Math.max(0, Math.min(run.count, units.length - at))
    if (take > 0) out.push({ text: units.slice(at, at + take).join(''), format: run })
    at += take
  }
  if (at < units.length) out.push({ text: units.slice(at).join(''), format: null })
  return out
}
