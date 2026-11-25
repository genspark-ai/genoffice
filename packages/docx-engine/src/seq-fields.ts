import { fieldNumberText, isFieldNumberFormat } from './field-number-format'
import { formatSeqNumber, SEQ_NUMBER_FORMATS, type SeqNumberFormat } from './generate'

/** a field whose result is a number counted over the document: SEQ, and the STYLEREF a caption takes its chapter from */
export interface XmlSeqField {
  kind: 'SEQ' | 'STYLEREF'
  /** from the field's opening `w:fldChar` or `w:fldSimple` tag to the end of its closing one */
  range: [number, number]
  instr: string
  result: string
  /** the text of each `w:t` of the result, as [start, end) offsets into the XML */
  spans: Array<[number, number]>
  /** the ` w:dirty="…"` attribute of the field's opening tag, when it has one */
  dirtyAttr: [number, number] | null
  /** `w:fldLock` on the opening tag: the result stands as the file has it */
  locked: boolean
}

const FIELD_TOKEN =
  /<w:fldChar\b([^>]*)>|<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>|<w:t\b[^>]*>([^<]*)<\/w:t>|<w:fldSimple\b([^>]*)>|<\/w:fldSimple>/g

const decodeXml = (text: string): string =>
  text
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')

const LOCKED = /\sw:fldLock="(?:1|true)"/

export const NUMBERING_FIELD = /(?:SEQ|STYLEREF)\s/i

const kindOf = (instr: string): XmlSeqField['kind'] | null => {
  const word = /^\s*(SEQ|STYLEREF)\s/i.exec(instr)?.[1].toUpperCase()
  return word === 'SEQ' || word === 'STYLEREF' ? word : null
}

interface Frame {
  instr: string
  separated: boolean
  simple: boolean
  spans: Array<[number, number]>
  result: string
  dirtyAttr: [number, number] | null
  locked: boolean
}

/** the SEQ and STYLEREF fields of a block's XML in document order, nested ones included */
export function seqFieldsOfXml(xml: string): XmlSeqField[] {
  if (!NUMBERING_FIELD.test(xml)) return []
  const out: Array<XmlSeqField & { at: number }> = []
  const stack: Frame[] = []
  const starts: number[] = []
  const dirtyOf = (attrs: string, tagStart: number, tag: string): [number, number] | null => {
    const m = /\sw:dirty="[^"]*"/.exec(attrs)
    if (!m) return null
    const from = tagStart + tag.indexOf(attrs) + m.index
    return [from, from + m[0].length]
  }
  const close = (end: number) => {
    const frame = stack.pop()
    const at = starts.pop() ?? 0
    if (!frame) return
    const instr = decodeXml(frame.instr)
    const kind = kindOf(instr)
    if (kind)
      out.push({
        kind,
        range: [at, end],
        instr: instr.trim(),
        result: decodeXml(frame.result),
        spans: frame.spans,
        dirtyAttr: frame.dirtyAttr,
        locked: frame.locked,
        at,
      })
  }
  for (const m of xml.matchAll(FIELD_TOKEN)) {
    const top = stack[stack.length - 1]
    if (m[1] !== undefined) {
      const type = /w:fldCharType="(\w+)"/.exec(m[1])?.[1]
      if (type === 'begin') {
        starts.push(m.index)
        stack.push({
          instr: '',
          separated: false,
          simple: false,
          spans: [],
          result: '',
          dirtyAttr: dirtyOf(m[1], m.index, m[0]),
          locked: LOCKED.test(m[1]),
        })
      } else if (type === 'separate') {
        if (top && !top.simple) top.separated = true
      } else if (type === 'end') {
        if (top && !top.simple) close(m.index + m[0].length)
      }
    } else if (m[2] !== undefined) {
      if (top && !top.separated) top.instr += m[2]
    } else if (m[3] !== undefined) {
      if (top?.separated) {
        const start = m.index + m[0].length - '</w:t>'.length - m[3].length
        top.spans.push([start, start + m[3].length])
        top.result += m[3]
      }
    } else if (m[4] !== undefined) {
      starts.push(m.index)
      stack.push({
        instr: /\bw:instr="([^"]*)"/.exec(m[4])?.[1] ?? '',
        separated: true,
        simple: true,
        spans: [],
        result: '',
        dirtyAttr: dirtyOf(m[4], m.index, m[0]),
        locked: LOCKED.test(m[4]),
      })
    } else if (top?.simple) close(m.index + m[0].length)
  }
  return out.sort((a, b) => a.at - b.at).map(({ at: _at, ...field }) => field)
}

const SEQ_SWITCH = /\\([A-Za-z*#@])(?:\s*(?:"([^"]*)"|([^\s\\]+)))?/g

export interface SeqInstruction {
  id: string
  restart: number | null
  /** `\c`: the number before it again */
  repeat: boolean
  /** `\h` without a format: the field shows nothing */
  hidden: boolean
  /** one of the formats the caption dialogs offer; null when the instruction asks for another */
  format: SeqNumberFormat | null
  /** another `\*` number format that is computed here: the switch's argument as written */
  formatName: string | null
  /** `\s N`: the count starts again after each heading of level N or above it */
  headingLevel: number | null
  /** a switch whose numbering is not computed here */
  unsupported: boolean
}

export function parseSeqInstruction(instr: string): SeqInstruction | null {
  const body = /^\s*SEQ\s+([\s\S]*)$/i.exec(instr)?.[1]
  if (body === undefined) return null
  const id = /^(?:"([^"]+)"|([^\s\\]+))/.exec(body)
  if (!id) return null
  const out: SeqInstruction = {
    id: id[1] ?? id[2],
    restart: null,
    repeat: false,
    hidden: false,
    format: 'ARABIC',
    formatName: null,
    headingLevel: null,
    unsupported: false,
  }
  let formatted = false
  for (const m of body.slice(id[0].length).matchAll(SEQ_SWITCH)) {
    const arg = m[2] ?? m[3]
    switch (m[1]) {
      case 'r':
        out.restart = Number(arg)
        if (!Number.isInteger(out.restart)) out.unsupported = true
        break
      case 's':
        out.headingLevel = Number(arg)
        if (!Number.isInteger(out.headingLevel) || out.headingLevel < 1 || out.headingLevel > 9)
          out.unsupported = true
        break
      case 'c':
        out.repeat = true
        break
      case 'h':
        out.hidden = true
        break
      case 'n':
        break
      case '*': {
        if (arg === undefined || /^(?:MERGEFORMAT|CHARFORMAT)$/i.test(arg)) break
        formatted = true
        const known = SEQ_NUMBER_FORMATS.find((f) => f === arg)
        out.format = known ?? (/^arabic$/i.test(arg) ? 'ARABIC' : null)
        out.formatName = out.format === null && isFieldNumberFormat(arg) ? arg : null
        break
      }
      case '#':
        formatted = true
        out.format = null
        out.formatName = null
        break
      default:
        out.unsupported = true
    }
  }
  if (out.hidden && formatted) out.hidden = false
  return out
}

export class SeqCounter {
  private last = new Map<string, number>()
  private headings = 0
  /** by level: how many headings had passed when the last one of that level or above it did */
  private lastHeading: number[] = []
  private seen = new Map<string, number>()

  heading(level: number): void {
    this.headings++
    for (let l = level; l <= 9; l++) this.lastHeading[l] = this.headings
  }

  /** the field's number, and the text it shows (null: not ours to write) */
  next(seq: SeqInstruction): { value: number; text: string | null } {
    const restarted =
      seq.headingLevel !== null &&
      (this.lastHeading[seq.headingLevel] ?? 0) > (this.seen.get(seq.id) ?? 0)
    this.seen.set(seq.id, this.headings)
    const before = restarted ? undefined : this.last.get(seq.id)
    const value =
      seq.restart !== null ? seq.restart : seq.repeat ? (before ?? 1) : (before ?? 0) + 1
    this.last.set(seq.id, value)
    const text = seq.hidden
      ? null
      : seq.format !== null
        ? formatSeqNumber(value, seq.format)
        : seq.formatName
          ? fieldNumberText(value, seq.formatName)
          : null
    return { value, text }
  }
}

export type NumberingEvent =
  | { kind: 'paragraph'; styleId: string; level: number | undefined }
  | { kind: 'field'; field: XmlSeqField; index: number }

const PARAGRAPH_STYLE = /<w:pStyle\b[^>]*\bw:val="([^"]*)"[^>]*>/g

/**
 * A block's styled paragraphs and its SEQ / STYLEREF fields, in document order. `levels`
 * gives the heading level of a style id; without it no paragraph is reported.
 */
export function numberingEventsOfXml(
  xml: string,
  levels?: ReadonlyMap<string, number>,
): NumberingEvent[] {
  const out: Array<NumberingEvent & { at: number }> = []
  if (levels) {
    // the style a tracked change replaced is not the paragraph's
    const replaced = xml.includes('<w:pPrChange')
      ? [...xml.matchAll(/<w:pPrChange\b[\s\S]*?<\/w:pPrChange>/g)].map(
          (m) => [m.index, m.index + m[0].length] as const,
        )
      : []
    for (const m of xml.matchAll(PARAGRAPH_STYLE)) {
      if (replaced.some(([from, to]) => m.index > from && m.index < to)) continue
      const styleId = decodeXml(m[1])
      out.push({ kind: 'paragraph', styleId, level: levels.get(styleId), at: m.index })
    }
  }
  seqFieldsOfXml(xml).forEach((field, index) =>
    out.push({ kind: 'field', field, index, at: field.range[0] }),
  )
  return out.sort((a, b) => a.at - b.at).map(({ at: _at, ...event }) => event as NumberingEvent)
}

/**
 * Identifiers whose cached numbers are not what counting the fields in order gives.
 * `levels` (style id -> heading level) lets `\s` restarts be counted; without it a
 * label using `\s` is one of them.
 */
export function unreliableSeqLabels(
  xmls: Iterable<string | null | undefined>,
  levels?: ReadonlyMap<string, number>,
): Set<string> {
  const bad = new Set<string>()
  const counter = new SeqCounter()
  for (const xml of xmls) {
    if (!xml) continue
    for (const event of numberingEventsOfXml(xml, levels)) {
      if (event.kind === 'paragraph') {
        if (event.level) counter.heading(event.level)
        continue
      }
      const seq = parseSeqInstruction(event.field.instr)
      if (!seq) continue
      if (seq.unsupported || (seq.headingLevel !== null && !levels)) {
        bad.add(seq.id)
        continue
      }
      const { text } = counter.next(seq)
      if (text !== null && text !== event.field.result.trim()) bad.add(seq.id)
    }
  }
  return bad
}

export interface StyleRefInstruction {
  /** `STYLEREF 2`: the built-in heading of that level */
  level: number | null
  /** the style's name, when it is not a heading level */
  style: string | null
  /** the paragraph's text, its number as the list shows it (`\n`), or the number alone (`\s`, `\n \t`) */
  show: 'text' | 'number' | 'bare'
  /** `\r` / `\w`: the number with the levels above it, as far as the field's own place needs or all of them */
  context: 'relative' | 'full' | null
  /** `\p`: where the paragraph is from the field, alone or after its number */
  aboveBelow: boolean
  /** `\* Upper`, `Lower`, `FirstCap` or `Caps`, left for the caller to apply to the result */
  textCase: string | null
  /** `\t` beside `\r` or `\w`: the number in context without the words around it */
  bareContext: boolean
  /** a switch whose result is not computed here (any `\*` conversion but the four cases) */
  unsupported: boolean
}

export function parseStyleRefInstruction(instr: string): StyleRefInstruction | null {
  const body = /^\s*STYLEREF\s+([\s\S]*)$/i.exec(instr)?.[1]
  if (body === undefined) return null
  const name = /^(?:"([^"]+)"|([^\s\\]+))/.exec(body)
  if (!name) return null
  const style = (name[1] ?? name[2]).trim()
  const level = /^[1-9]$/.test(style) ? Number(style) : null
  const out: StyleRefInstruction = {
    level,
    style: level === null ? style : null,
    show: 'text',
    context: null,
    aboveBelow: false,
    textCase: null,
    bareContext: false,
    unsupported: false,
  }
  let stripped = false
  for (const m of body.slice(name[0].length).matchAll(SEQ_SWITCH)) {
    switch (m[1]) {
      case 'n':
        if (out.show === 'text') out.show = 'number'
        break
      case 's':
        out.show = 'bare'
        break
      case 't':
        stripped = true
        break
      case 'r':
        out.context ??= 'relative'
        break
      case 'w':
        out.context = 'full'
        break
      case 'p':
        out.aboveBelow = true
        break
      // the page is searched from its bottom only by a field in a header or footer: in the
      // text of the document the search runs back from the field either way
      case 'l':
        break
      case '*': {
        const arg = m[2] ?? m[3] ?? ''
        if (/^(?:Upper|Lower|FirstCap|Caps)$/i.test(arg)) out.textCase = arg
        else if (!/^(?:MERGEFORMAT|CHARFORMAT)$/i.test(arg)) out.unsupported = true
        break
      }
      default:
        out.unsupported = true
    }
  }
  if (stripped && out.show === 'number') out.show = 'bare'
  if (stripped && out.context) out.bareContext = true
  return out
}

/** a paragraph as STYLEREF sees it */
export interface StyledParagraph {
  styleId: string | null
  styleName?: string | null
  headingLevel?: number | null
  text: string
  /** its list number as shown ("Chapter 2."); null or absent when it has none */
  marker?: string | null
  /** the numbers of `marker` with the words around them dropped ("2") */
  number?: string | null
  /** what it shows cannot be told here: a field pointing at it keeps its result */
  unknown?: boolean
}

/** what a STYLEREF's `\r`, `\w` and `\p` need beyond the paragraphs themselves */
export interface StyleRefSurroundings {
  /** the paragraph's number in that context; null when it cannot be told */
  numberOf?: (
    paragraph: StyledParagraph,
    context: 'relative' | 'full',
    /** `\t`: without the words around the numbers; null where the levels do not read as one number */
    bare: boolean,
  ) => string | null
  /** `\p` in the language of the interface */
  words?: { above: string; below: string }
}

/** what Word leaves in a STYLEREF that finds no paragraph of its style */
export const STYLEREF_NOT_FOUND = 'Error! No text of specified style in document.'

const sameName = (a: string | null | undefined, b: string): boolean =>
  !!a && a.toLowerCase() === b.toLowerCase()

/**
 * The result of a STYLEREF standing after the first `before` paragraphs: the nearest
 * paragraph of its style going back, else the first one going on. A paragraph without
 * a number gives "0", as Word saves it. Null: not computed here.
 */
export function styleRefResult(
  ref: StyleRefInstruction,
  paragraphs: readonly StyledParagraph[],
  before: number,
  around: StyleRefSurroundings = {},
): string | null {
  if (ref.unsupported) return null
  const matches = (p: StyledParagraph): boolean =>
    ref.level !== null
      ? p.headingLevel === ref.level
      : sameName(p.styleName, ref.style!) || sameName(p.styleId, ref.style!)
  let at = -1
  for (let i = Math.min(before, paragraphs.length) - 1; i >= 0 && at < 0; i--)
    if (matches(paragraphs[i])) at = i
  for (let i = Math.max(before, 0); i < paragraphs.length && at < 0; i++)
    if (matches(paragraphs[i])) at = i
  if (at < 0) return STYLEREF_NOT_FOUND
  const found = paragraphs[at]
  if (found.unknown) return null
  const word = ref.aboveBelow ? around.words?.[at < before ? 'above' : 'below'] : undefined
  if (ref.aboveBelow && word === undefined) return null
  const placed = (number: string): string => (word === undefined ? number : `${number} ${word}`)
  if (ref.context) {
    if (!found.marker?.trim()) return placed('0')
    const number = around.numberOf?.(found, ref.context, ref.bareContext) ?? null
    return number === null ? null : placed(number)
  }
  // `\p` alone says where the paragraph is and nothing of what it holds
  if (ref.show === 'text') return word ?? found.text
  if (ref.show === 'bare') return placed(found.number || '0')
  return placed(found.marker?.trim().replace(/\.+$/, '') || '0')
}

/** `results[i]` becomes the i-th field's result (null leaves it); a rewritten field loses its `w:dirty` */
export function rewriteSeqResultsXml(xml: string, results: ReadonlyArray<string | null>): string {
  const edits: Array<{ from: number; to: number; text: string }> = []
  seqFieldsOfXml(xml).forEach((field, i) => {
    const text = results[i]
    if (text === null || text === undefined || field.spans.length === 0 || field.locked) return
    field.spans.forEach(([from, to], n) =>
      edits.push({ from, to, text: n === 0 ? escapeText(text) : '' }),
    )
    if (field.dirtyAttr) edits.push({ from: field.dirtyAttr[0], to: field.dirtyAttr[1], text: '' })
  })
  let out = xml
  for (const edit of edits.sort((a, b) => b.from - a.from))
    out = out.slice(0, edit.from) + edit.text + out.slice(edit.to)
  return out
}

const escapeText = (text: string): string =>
  text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
