import { formatSeqNumber, SEQ_NUMBER_FORMATS, type SeqNumberFormat } from './generate'

export interface XmlSeqField {
  instr: string
  result: string
  /** the text of each `w:t` of the result, as [start, end) offsets into the XML */
  spans: Array<[number, number]>
  /** the ` w:dirty="…"` attribute of the field's opening tag, when it has one */
  dirtyAttr: [number, number] | null
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

const isSeq = (instr: string): boolean => /^\s*SEQ\s/i.test(instr)

interface Frame {
  instr: string
  separated: boolean
  simple: boolean
  spans: Array<[number, number]>
  result: string
  dirtyAttr: [number, number] | null
}

/** the SEQ fields of a block's XML in document order, nested ones included */
export function seqFieldsOfXml(xml: string): XmlSeqField[] {
  if (!/SEQ\s/i.test(xml)) return []
  const out: Array<XmlSeqField & { at: number }> = []
  const stack: Frame[] = []
  const starts: number[] = []
  const dirtyOf = (attrs: string, tagStart: number, tag: string): [number, number] | null => {
    const m = /\sw:dirty="[^"]*"/.exec(attrs)
    if (!m) return null
    const from = tagStart + tag.indexOf(attrs) + m.index
    return [from, from + m[0].length]
  }
  const close = () => {
    const frame = stack.pop()
    const at = starts.pop() ?? 0
    if (!frame) return
    const instr = decodeXml(frame.instr)
    if (isSeq(instr))
      out.push({
        instr: instr.trim(),
        result: decodeXml(frame.result),
        spans: frame.spans,
        dirtyAttr: frame.dirtyAttr,
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
        })
      } else if (type === 'separate') {
        if (top && !top.simple) top.separated = true
      } else if (type === 'end') {
        if (top && !top.simple) close()
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
      })
    } else if (top?.simple) close()
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
  /** null when the instruction asks for a number format this cannot write */
  format: SeqNumberFormat | null
  /** a switch whose numbering is not computed here (`\s`: restart at a heading level) */
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
        break
      }
      case '#':
        formatted = true
        out.format = null
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

  /** the field's number, and the text it shows (null: not ours to write) */
  next(seq: SeqInstruction): { value: number; text: string | null } {
    const before = this.last.get(seq.id)
    const value =
      seq.restart !== null ? seq.restart : seq.repeat ? (before ?? 1) : (before ?? 0) + 1
    this.last.set(seq.id, value)
    const text = seq.hidden || seq.format === null ? null : formatSeqNumber(value, seq.format)
    return { value, text }
  }
}

/** identifiers whose cached numbers are not what counting the fields in order gives */
export function unreliableSeqLabels(xmls: Iterable<string | null | undefined>): Set<string> {
  const bad = new Set<string>()
  const counter = new SeqCounter()
  for (const xml of xmls) {
    if (!xml) continue
    for (const field of seqFieldsOfXml(xml)) {
      const seq = parseSeqInstruction(field.instr)
      if (!seq) continue
      if (seq.unsupported) {
        bad.add(seq.id)
        continue
      }
      const { text } = counter.next(seq)
      if (text !== null && text !== field.result.trim()) bad.add(seq.id)
    }
  }
  return bad
}

/** `results[i]` becomes the i-th SEQ field's result (null leaves it); a rewritten field loses its `w:dirty` */
export function rewriteSeqResultsXml(xml: string, results: ReadonlyArray<string | null>): string {
  const edits: Array<{ from: number; to: number; text: string }> = []
  seqFieldsOfXml(xml).forEach((field, i) => {
    const text = results[i]
    if (text === null || text === undefined || field.spans.length === 0) return
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
