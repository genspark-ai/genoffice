import { type ColorScheme, colorIndexColor } from './escher'
import { PPT_TYPES, type PptRecord, payloadOf, walkRecordsLenient } from './records'

/**
 * PowerPoint text formatting: the paragraph/character exception runs of a
 * StyleTextPropAtom, the per-level defaults of a TxMasterStyleAtom, and the
 * FontCollection the fontRef indexes resolve against.
 */

export interface ParagraphProps {
  level: number
  bullet?: boolean
  bulletChar?: string
  bulletFontRef?: number
  bulletColor?: number
  bulletSize?: number
  /** fBulletHasFont / fBulletHasColor / fBulletHasSize: false means the bullet follows the text */
  bulletHasFont?: boolean
  bulletHasColor?: boolean
  bulletHasSize?: boolean
  /** 0 left, 1 center, 2 right, 3 justify, 4 distributed */
  align?: number
  /** > 0 percent of a line, < 0 master units (1/8 pt) */
  lineSpacing?: number
  spaceBefore?: number
  spaceAfter?: number
  /** master units */
  leftMargin?: number
  indent?: number
  /** tab stops, master units */
  tabs?: TabStop[]
}

export interface CharProps {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  shadow?: boolean
  fontRef?: number
  eaFontRef?: number
  /** points */
  size?: number
  /** a ColorIndexStruct */
  color?: number
  /** percent baseline shift */
  position?: number
}

export interface ParagraphRun extends ParagraphProps {
  count: number
}

export interface CharRun extends CharProps {
  count: number
}

export interface StyleRuns {
  paragraphs: ParagraphRun[]
  chars: CharRun[]
}

/** a PP9 auto-numbered bullet: TextAutoNumberSchemeEnum and the first number */
export interface AutoNumber {
  scheme: number
  startAt: number
}

/** ST_TextAutonumberScheme names by TextAutoNumberSchemeEnum value */
export const AUTO_NUMBER_SCHEMES = [
  'alphaLcPeriod',
  'alphaUcPeriod',
  'arabicParenR',
  'arabicPeriod',
  'romanLcParenBoth',
  'romanLcParenR',
  'romanLcPeriod',
  'romanUcPeriod',
  'alphaLcParenBoth',
  'alphaLcParenR',
  'alphaUcParenBoth',
  'alphaUcParenR',
  'arabicParenBoth',
  'arabicPlain',
  'romanUcParenBoth',
  'romanUcParenR',
  'ea1ChsPlain',
  'ea1ChsPeriod',
  'circleNumDbPlain',
  'circleNumWdWhitePlain',
  'circleNumWdBlackPlain',
  'ea1ChtPlain',
  'ea1ChtPeriod',
  'arabic1Minus',
  'arabic2Minus',
  'hebrew2Minus',
  'ea1JpnKorPlain',
  'ea1JpnKorPeriod',
  'arabicDbPlain',
  'arabicDbPeriod',
  'thaiAlphaPeriod',
  'thaiAlphaParenR',
  'thaiAlphaParenBoth',
  'thaiNumPeriod',
  'thaiNumParenR',
  'thaiNumParenBoth',
  'hindiAlphaPeriod',
  'hindiNumPeriod',
  'hindiAlpha1Period',
  'hindiNumParenR',
] as const

function viewOf(bytes: Uint8Array): DataView {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
}

class Cursor {
  at = 0
  constructor(
    readonly bytes: Uint8Array,
    readonly v = viewOf(bytes),
  ) {}
  get ok(): boolean {
    return this.at <= this.bytes.length
  }
  u16(): number {
    const value = this.at + 2 <= this.bytes.length ? this.v.getUint16(this.at, true) : 0
    this.at += 2
    return value
  }
  i16(): number {
    const value = this.at + 2 <= this.bytes.length ? this.v.getInt16(this.at, true) : 0
    this.at += 2
    return value
  }
  u32(): number {
    const value = this.at + 4 <= this.bytes.length ? this.v.getUint32(this.at, true) : 0
    this.at += 4
    return value
  }
}

/** TextPFException masks → fields, in the order the bytes lay them out */
function readParagraphProps(c: Cursor, into: ParagraphProps): void {
  const mask = c.u32()
  if (mask & 0x000f) {
    const flags = c.u16()
    if (mask & 0x0001) into.bullet = (flags & 0x0001) !== 0
    if (mask & 0x0002) into.bulletHasFont = (flags & 0x0002) !== 0
    if (mask & 0x0004) into.bulletHasColor = (flags & 0x0004) !== 0
    if (mask & 0x0008) into.bulletHasSize = (flags & 0x0008) !== 0
  }
  if (mask & 0x0080) into.bulletChar = String.fromCharCode(c.u16())
  if (mask & 0x0010) into.bulletFontRef = c.u16()
  if (mask & 0x0040) into.bulletSize = c.i16()
  if (mask & 0x0020) into.bulletColor = c.u32()
  if (mask & 0x0800) into.align = c.u16()
  if (mask & 0x1000) into.lineSpacing = c.i16()
  if (mask & 0x2000) into.spaceBefore = c.i16()
  if (mask & 0x4000) into.spaceAfter = c.i16()
  if (mask & 0x0100) into.leftMargin = c.i16()
  if (mask & 0x0400) into.indent = c.i16()
  if (mask & 0x8000) c.u16() // defaultTabSize
  if (mask & 0x100000) into.tabs = readTabStops(c)
  if (mask & 0x10000) c.u16() // fontAlign
  if (mask & 0xe0000) c.u16() // wrap flags
  if (mask & 0x200000) c.u16() // text direction
}

/** TextCFException masks → fields */
function readCharProps(c: Cursor, into: CharProps): void {
  const mask = c.u32()
  if (mask & 0x3eb7) {
    const style = c.u16()
    if (mask & 0x0001) into.bold = (style & 0x0001) !== 0
    if (mask & 0x0002) into.italic = (style & 0x0002) !== 0
    if (mask & 0x0004) into.underline = (style & 0x0004) !== 0
    if (mask & 0x0010) into.shadow = (style & 0x0010) !== 0
  }
  if (mask & 0x00010000) into.fontRef = c.u16()
  if (mask & 0x00200000) into.eaFontRef = c.u16()
  if (mask & 0x00400000) c.u16() // ansi font
  if (mask & 0x00800000) c.u16() // symbol font
  if (mask & 0x00020000) into.size = c.u16()
  if (mask & 0x00040000) into.color = c.u32()
  if (mask & 0x00080000) into.position = c.i16()
}

/** A StyleTextPropAtom over a text of `textLength` characters. */
export function parseStyleRuns(payload: Uint8Array, textLength: number): StyleRuns {
  const c = new Cursor(payload)
  const paragraphs: ParagraphRun[] = []
  const chars: CharRun[] = []
  let covered = 0
  // the runs cover the text plus its terminating paragraph mark
  const total = textLength + 1
  while (covered < total && c.at + 10 <= payload.length) {
    const count = c.u32()
    const level = c.u16()
    const run: ParagraphRun = { count, level }
    readParagraphProps(c, run)
    paragraphs.push(run)
    covered += count
    if (count === 0) break
  }
  covered = 0
  while (covered < total && c.at + 8 <= payload.length) {
    const count = c.u32()
    const run: CharRun = { count }
    readCharProps(c, run)
    chars.push(run)
    covered += count
    if (count === 0) break
  }
  return { paragraphs, chars }
}

/**
 * A StyleTextProp9Atom: one entry per paragraph run of the StyleTextPropAtom,
 * each a TextPFException9 + TextCFException9 + TextSIException. Only the
 * auto-number scheme is kept; the rest is walked to stay aligned.
 */
export function parseStyleTextProp9(payload: Uint8Array): Array<AutoNumber | null> {
  const c = new Cursor(payload)
  const out: Array<AutoNumber | null> = []
  while (c.at + 12 <= payload.length) {
    const pfMask = c.u32()
    if (pfMask & 0x800000) c.u16() // bulletBlipRef
    const hasAutoNumber = pfMask & 0x2000000 ? c.u16() : undefined
    let scheme: AutoNumber | null = null
    if (pfMask & 0x1000000) scheme = { scheme: c.u16(), startAt: c.u16() }
    const numbered = hasAutoNumber === undefined ? scheme !== null : hasAutoNumber !== 0
    out.push(numbered ? (scheme ?? { scheme: 3, startAt: 1 }) : null)
    const cfMask = c.u32()
    if (cfMask & 0x100000) c.u32() // pp10runid
    const siMask = c.u32()
    if (siMask & 0x0001) c.u16() // spellInfo
    if (siMask & 0x0002) c.u16() // lid
    if (siMask & 0x0004) c.u16() // altLid
    if (siMask & 0x0040) c.u16() // bidi
    if (siMask & 0x0020) c.u32() // pp10runid
    if (siMask & 0x0200) c.at += 4 + 4 * c.u32() // smartTags
  }
  return out
}

export interface LevelStyle {
  paragraph: ParagraphProps
  chars: CharProps
}

/** one indent level of a TextRulerAtom, in master units */
export interface RulerLevel {
  leftMargin?: number
  indent?: number
}

/** a tab stop: position in master units; 0 left, 1 centre, 2 right, 3 decimal */
export interface TabStop {
  position: number
  type: number
}

export interface TextRuler {
  levels: RulerLevel[]
  tabs: TabStop[]
}

function readTabStops(c: Cursor): TabStop[] {
  const count = Math.min(c.u16(), 0x1000)
  const tabs: TabStop[] = []
  for (let index = 0; index < count && c.at + 4 <= c.bytes.length; index++) {
    tabs.push({ position: c.i16(), type: c.u16() })
  }
  return tabs
}

/**
 * A TextRulerAtom: the tab stops and, per indent level, the text position and
 * the first-line/bullet position the ruler sets for this text.
 */
export function parseTextRuler(payload: Uint8Array): TextRuler {
  const c = new Cursor(payload)
  const mask = c.u32()
  if (mask & 0x0002) c.u16()
  if (mask & 0x0001) c.u16()
  const tabs = mask & 0x0004 ? readTabStops(c) : []
  const levels: RulerLevel[] = []
  for (let level = 0; level < 5; level++) {
    const entry: RulerLevel = {}
    if (mask & (0x0008 << level)) entry.leftMargin = c.i16()
    if (mask & (0x0100 << level)) entry.indent = c.i16()
    levels.push(entry)
  }
  return { levels, tabs }
}

/**
 * The centre-body / centre-title / half- and quarter-body styles (instances
 * 5-8) state only what differs from the body (or title) style they refine.
 */
export function layerLevels(base: LevelStyle[], over: LevelStyle[] | undefined): LevelStyle[] {
  if (!over || over.length === 0) return base
  const count = Math.max(base.length, over.length)
  const out: LevelStyle[] = []
  for (let level = 0; level < count; level++) {
    const under = base[level] ?? base[base.length - 1]
    const top = over[level] ?? over[over.length - 1]
    out.push({
      paragraph: { ...under?.paragraph, ...stripUndefined(top?.paragraph ?? {}), level },
      chars: { ...under?.chars, ...stripUndefined(top?.chars ?? {}) },
    })
  }
  return out
}

/**
 * A TxMasterStyleAtom: per indent level, the default paragraph and character
 * properties of one text type. Levels for the body-like types (instance ≥ 5)
 * open with their own 16-bit level number.
 */
export function parseMasterStyle(payload: Uint8Array, instance: number): LevelStyle[] {
  const c = new Cursor(payload)
  const levels = Math.min(c.u16(), 5)
  const out: LevelStyle[] = []
  for (let level = 0; level < levels && c.at < payload.length; level++) {
    if (instance >= 5) c.u16()
    const paragraph: ParagraphProps = { level }
    readParagraphProps(c, paragraph)
    const chars: CharProps = {}
    readCharProps(c, chars)
    // a level states only what differs from the level above it
    const above = out[out.length - 1]
    out.push(
      above
        ? {
            paragraph: { ...above.paragraph, ...stripUndefined(paragraph), level },
            chars: { ...above.chars, ...stripUndefined(chars) },
          }
        : { paragraph, chars },
    )
  }
  return out
}

/** The FontCollection's FontEntityAtoms: fontRef → face name. */
export function parseFontCollection(stream: Uint8Array, collection: PptRecord): string[] {
  const names: string[] = []
  for (const record of walkRecordsLenient(
    stream,
    collection.offset,
    collection.offset + collection.length,
  )) {
    if (record.type !== PPT_TYPES.fontEntityAtom) continue
    const payload = payloadOf(stream, record)
    const raw = Buffer.from(
      payload.buffer,
      payload.byteOffset,
      Math.min(64, payload.length),
    ).toString('utf16le')
    names.push(raw.split('\0')[0] ?? '')
  }
  return names
}

/** The text-type defaults a run falls back to: title-like types share the title style. */
export function masterStyleIndex(textKind: number): number {
  switch (textKind) {
    case 0:
    case 6:
      return 0
    case 1:
    case 5:
    case 7:
    case 8:
      return 1
    case 2:
      return 2
    default:
      return 4
  }
}

/** the master levels of a text type: instances 5-8 layered over the style they refine */
export function levelsForKind(
  textKind: number,
  styles: (instance: number) => LevelStyle[] | undefined,
): LevelStyle[] {
  const base = styles(masterStyleIndex(textKind)) ?? []
  return textKind >= 5 && textKind <= 8 ? layerLevels(base, styles(textKind)) : base
}

/**
 * The levels a text resolves against. A slide text box reads the document's default
 * text style; the master's other-text style (DrawingML's otherStyle) is only for
 * shapes on the master itself.
 */
export function levelsForText(
  textKind: number,
  onMaster: boolean,
  masterStyle: (instance: number) => LevelStyle[] | undefined,
  docStyle: (instance: number) => LevelStyle[] | undefined,
): { levels: LevelStyle[]; base: LevelStyle[] } {
  const base = docStyle(4) ?? masterStyle(4) ?? []
  if (textKind === 4 && !onMaster) return { levels: base, base }
  return {
    levels: levelsForKind(textKind, (instance) => masterStyle(instance) ?? docStyle(instance)),
    base,
  }
}

export interface ResolvedRun {
  text: string
  bold: boolean
  italic: boolean
  underline: boolean
  /** points */
  size: number
  /** RRGGBB */
  color: string
  /** RRGGBB of the text shadow, when the run casts one */
  shadow?: string
  font?: string
  eaFont?: string
  /** percent baseline shift */
  position: number
  /** a dynamic field the text stands for (slide number, date, header, footer) */
  field?: 'slidenum' | 'datetime1' | 'header' | 'footer'
}

export interface ResolvedParagraph {
  level: number
  align: number
  bullet: boolean
  bulletChar: string
  bulletFont?: string
  bulletColor?: string
  bulletSize?: number
  lineSpacing?: number
  spaceBefore?: number
  spaceAfter?: number
  leftMargin: number
  indent: number
  tabs?: TabStop[]
  /** runs; a '\u000b' run boundary is a soft line break */
  runs: ResolvedRun[]
  /** points: the size of the paragraph mark (what an empty paragraph measures) */
  endSize: number
  autoNumber?: AutoNumber
}

export interface TextDefaults {
  /** master style per indent level for this text's type */
  levels: LevelStyle[]
  /** the "other" text style the levels fall back to */
  base: LevelStyle[]
  fonts: string[]
  scheme: ColorScheme
  /** the text's own ruler: its margins sit between the master levels and the paragraph runs */
  ruler?: TextRuler
  autoNumbers?: Array<AutoNumber | null>
}

const DEFAULT_SIZE = 18

/**
 * Lay the exception runs over the raw text: paragraphs split on '\r', each
 * paragraph takes the paragraph run covering its first character, and each
 * character run becomes one output run (merged across equal properties).
 */
export function resolveText(
  raw: string,
  runs: StyleRuns | null,
  defaults: TextDefaults,
): ResolvedParagraph[] {
  const out: ResolvedParagraph[] = []
  const paragraphs = raw.split(/\r\n|\r|\n/)
  // trailing empty paragraph from a terminating '\r' is not a paragraph of its own
  if (paragraphs.length > 1 && paragraphs[paragraphs.length - 1] === '') paragraphs.pop()
  let paraAt = 0
  let paraRunIndex = 0
  let paraRunLeft = runs?.paragraphs[0]?.count ?? Number.POSITIVE_INFINITY
  let charRunIndex = 0
  let charRunLeft = runs?.chars[0]?.count ?? Number.POSITIVE_INFINITY
  for (const text of paragraphs) {
    const paraRun = runs?.paragraphs[paraRunIndex]
    const level = Math.min(paraRun?.level ?? 0, 4)
    const levelStyle = defaults.levels[level] ?? defaults.levels[0]
    const baseStyle = defaults.base[level] ?? defaults.base[0]
    const pp = {
      ...baseStyle?.paragraph,
      ...levelStyle?.paragraph,
      ...stripUndefined(defaults.ruler?.levels[level] ?? {}),
      ...stripUndefined(paraRun ?? {}),
    }
    const tabs = pp.tabs?.length ? pp.tabs : defaults.ruler?.tabs
    const paragraph: ResolvedParagraph = {
      level,
      align: pp.align ?? 0,
      bullet: pp.bullet ?? false,
      bulletChar: pp.bulletChar ?? '•',
      bulletFont:
        pp.bulletHasFont !== false && pp.bulletFontRef !== undefined
          ? defaults.fonts[pp.bulletFontRef]
          : undefined,
      bulletColor:
        pp.bulletHasColor !== false && pp.bulletColor !== undefined
          ? colorIndexColor(pp.bulletColor, defaults.scheme)
          : undefined,
      bulletSize: pp.bulletHasSize !== false ? pp.bulletSize : undefined,
      lineSpacing: pp.lineSpacing,
      spaceBefore: pp.spaceBefore,
      spaceAfter: pp.spaceAfter,
      leftMargin: pp.leftMargin ?? 0,
      indent: pp.indent ?? 0,
      ...(tabs && tabs.length > 0 ? { tabs } : {}),
      ...(defaults.autoNumbers?.[paraRunIndex]
        ? { autoNumber: defaults.autoNumbers[paraRunIndex]! }
        : {}),
      runs: [],
      endSize: DEFAULT_SIZE,
    }
    // characters of this paragraph plus its paragraph mark
    const length = text.length + 1
    let consumed = 0
    while (consumed < length) {
      const charRun = runs?.chars[charRunIndex]
      const take = Math.min(charRunLeft, length - consumed)
      const slice = text.slice(consumed, Math.min(consumed + take, text.length))
      const cp = { ...baseStyle?.chars, ...levelStyle?.chars, ...stripUndefined(charRun ?? {}) }
      // the run covering the paragraph mark sizes the mark itself
      if (consumed + take >= length)
        paragraph.endSize = cp.size && cp.size > 0 ? cp.size : DEFAULT_SIZE
      if (slice.length > 0) {
        const run: ResolvedRun = {
          text: slice,
          bold: cp.bold ?? false,
          italic: cp.italic ?? false,
          underline: cp.underline ?? false,
          size: cp.size && cp.size > 0 ? cp.size : DEFAULT_SIZE,
          color:
            (cp.color !== undefined ? colorIndexColor(cp.color, defaults.scheme) : undefined) ??
            defaults.scheme.colors[1] ??
            '000000',
          shadow: cp.shadow ? (defaults.scheme.colors[2] ?? '808080') : undefined,
          font: cp.fontRef !== undefined ? defaults.fonts[cp.fontRef] : undefined,
          eaFont: cp.eaFontRef !== undefined ? defaults.fonts[cp.eaFontRef] : undefined,
          position: cp.position ?? 0,
        }
        const last = paragraph.runs[paragraph.runs.length - 1]
        if (last && sameRunProps(last, run)) last.text += run.text
        else paragraph.runs.push(run)
      }
      consumed += take
      charRunLeft -= take
      if (charRunLeft <= 0) {
        charRunIndex++
        charRunLeft = runs?.chars[charRunIndex]?.count ?? Number.POSITIVE_INFINITY
      }
    }
    paraAt += length
    paraRunLeft -= length
    if (paraRunLeft <= 0) {
      paraRunIndex++
      paraRunLeft = runs?.paragraphs[paraRunIndex]?.count ?? Number.POSITIVE_INFINITY
    }
    out.push(paragraph)
  }
  void paraAt
  return out
}

function sameRunProps(a: ResolvedRun, b: ResolvedRun): boolean {
  return (
    a.bold === b.bold &&
    a.italic === b.italic &&
    a.underline === b.underline &&
    a.size === b.size &&
    a.color === b.color &&
    a.shadow === b.shadow &&
    a.font === b.font &&
    a.eaFont === b.eaFont &&
    a.position === b.position
  )
}

function stripUndefined<T extends object>(value: T): Partial<T> {
  const out: Partial<T> = {}
  for (const [key, entry] of Object.entries(value)) {
    if (entry !== undefined) (out as Record<string, unknown>)[key] = entry
  }
  return out
}
