import { onlyXeFields } from './parse-props'
import { decodeEntities } from './parse-xml-text'
import { escapeXmlText } from './xml-utils'

/** the fields Word lists under Edit > Links: content brought in from another file */
export const LINK_FIELD_RE = /^\s*(INCLUDETEXT|INCLUDEPICTURE|LINK)(?=\s|$)/i

export type LinkFieldKind = 'INCLUDETEXT' | 'INCLUDEPICTURE' | 'LINK'

/**
 * The name Paste link gives the range it links to. Only bookmarks of this
 * name are read as a range (Run.rangeMark), so the range holds when its
 * paragraphs are edited; every other bookmark is read as it always was.
 */
export const LINK_SOURCE_PREFIX = 'LinkSource'

export interface LinkFieldInfo {
  kind: LinkFieldKind
  /** the file the field names, as written (escaped backslashes undone) */
  source: string
  /** INCLUDETEXT: the bookmark; LINK: the item inside the source */
  item: string
  /** LINK: the application class (`Excel.Sheet.12`) */
  app: string
  /** LINK's `\a` switch: Word refreshes it without being asked */
  automatic: boolean
}

const ARG_RE = /"((?:[^"\\]|\\.)*)"|(\S+)/g
const unescapeArg = (quoted: string | undefined, bare: string): string =>
  quoted !== undefined ? quoted.replace(/\\(.)/g, '$1') : bare.replace(/\\\\/g, '\\')

/** a switch, and whether the word after it belongs to it */
function switchOf(word: string, kind: LinkFieldKind): { takesArg: boolean } | null {
  if (!/^\\[a-zA-Z*#@!]$/.test(word)) return null
  const letter = word[1].toLowerCase()
  if ('*#@'.includes(letter)) return { takesArg: true }
  return { takesArg: kind === 'LINK' ? letter === 'f' : 'ctx'.includes(letter) }
}

/** reads a link field's instruction; null when it is not one */
export function parseLinkInstruction(instr: string): LinkFieldInfo | null {
  const kind = LINK_FIELD_RE.exec(instr)?.[1]?.toUpperCase() as LinkFieldKind | undefined
  if (!kind) return null
  const rest = instr.replace(LINK_FIELD_RE, '')
  const args: string[] = []
  let automatic = false
  ARG_RE.lastIndex = 0
  let m: RegExpExecArray | null
  while ((m = ARG_RE.exec(rest)) !== null) {
    const sw = m[1] === undefined ? switchOf(m[2], kind) : null
    if (!sw) {
      args.push(unescapeArg(m[1], m[2]))
      continue
    }
    if (kind === 'LINK' && m[2].toLowerCase() === '\\a') automatic = true
    if (sw.takesArg) ARG_RE.exec(rest)
  }
  if (kind === 'LINK')
    return { kind, app: args[0] ?? '', source: args[1] ?? '', item: args[2] ?? '', automatic }
  return { kind, app: '', source: args[0] ?? '', item: args[1] ?? '', automatic: false }
}

const quoteArg = (value: string): string => `"${value.replace(/\\/g, '\\\\').replace(/"/g, '')}"`

/** an INCLUDETEXT instruction for a file and, when given, one of its bookmarks */
export function includeTextInstruction(source: string, bookmark?: string | null): string {
  return `INCLUDETEXT ${quoteArg(source)}${bookmark ? ` ${bookmark}` : ''}`
}

/**
 * A LINK instruction for a range of a Word document, as Paste link writes it:
 * `format` is the result's kind, `r` formatted text or `t` unformatted text.
 */
export function wordLinkInstruction(source: string, bookmark: string, format: 'r' | 't'): string {
  return `LINK Word.Document.12 ${quoteArg(source)} ${quoteArg(bookmark)} \\${format}`
}

/** a LINK field to a range of a Word document: its source is read like INCLUDETEXT's */
export function isWordDocumentLink(info: LinkFieldInfo): boolean {
  return info.kind === 'LINK' && /^Word\.Document\b/i.test(info.app) && /\.docx$/i.test(info.source)
}

/** the result of a LINK field is kept as text without formatting (`\t`) */
export const linkResultIsPlain = (instr: string): boolean => /\s\\t(?=\s|$)/i.test(instr)

/** the instruction naming another file; everything else it says stays */
export function withLinkSource(instr: string, source: string): string {
  const info = parseLinkInstruction(instr)
  if (!info) return instr
  const head = LINK_FIELD_RE.exec(instr)![0]
  const rest = instr.slice(head.length)
  // LINK names its application first
  let skip = info.kind === 'LINK' ? 1 : 0
  ARG_RE.lastIndex = 0
  let m: RegExpExecArray | null
  while ((m = ARG_RE.exec(rest)) !== null) {
    const sw = m[1] === undefined ? switchOf(m[2], info.kind) : null
    if (sw) {
      if (sw.takesArg) ARG_RE.exec(rest)
      continue
    }
    if (skip-- > 0) continue
    return head + rest.slice(0, m.index) + quoteArg(source) + rest.slice(m.index + m[0].length)
  }
  return `${head}${rest.trimEnd()} ${quoteArg(source)} `
}

/** LINK's `\a` switch on or off */
export function withLinkAutomatic(instr: string, automatic: boolean): string {
  const without = instr.replace(/\s\\a(?=\s|$)/gi, '')
  if (!automatic) return without
  return `${without.trimEnd()} \\a${/\s$/.test(instr) ? ' ' : ''}`
}

interface Span {
  start: number
  end: number
}

interface FoundField {
  instr: string
  lock: boolean
  beginPara: number
  /** the begin run through the separate run */
  begin: Span
  endPara: number
  end: Span
}

interface Frame {
  candidate: boolean
  para: number
  beginRun: Span
  lock: boolean
  instr: string
  separated: boolean
  separate?: Span
}

const TOKEN_RE = /<w:fldChar\b[^>]*>|<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/g
const WRAPPER_OPEN_RE =
  /<w:(?:ins|del|moveFrom|moveTo|hyperlink|sdt|smartTag|fldSimple|customXml)\b[^>]*(?<!\/)>/g
const WRAPPER_CLOSE_RE =
  /<\/w:(?:ins|del|moveFrom|moveTo|hyperlink|sdt|smartTag|fldSimple|customXml)>/g
/** content whose paragraph is read by another path than the run reader */
const HOST_BLOCKED_RE =
  /<w:drawing|<w:pict[\s>]|<w:object|<m:oMath|<w:sectPr|<mc:AlternateContent|<w:txbxContent|<w:delInstrText|<w:cellIns|<w:cellDel/

function runSpan(xml: string, at: number): Span | null {
  const start = Math.max(xml.lastIndexOf('<w:r>', at), xml.lastIndexOf('<w:r ', at))
  const close = xml.indexOf('</w:r>', at)
  if (start === -1 || close === -1) return null
  return { start, end: close + '</w:r>'.length }
}

/** the run holds one field character and nothing a reader would show */
function bareFieldRun(xml: string, span: Span): boolean {
  const inner = xml
    .slice(xml.indexOf('>', span.start) + 1, span.end - '</w:r>'.length)
    .replace(/<w:rPr>[\s\S]*?<\/w:rPr>|<w:rPr\/>/, '')
  return /^\s*<w:fldChar\b[^>]*\/>\s*$/.test(inner)
}

function directChild(xml: string, at: number): boolean {
  const before = xml.slice(0, at)
  return (
    (before.match(WRAPPER_OPEN_RE)?.length ?? 0) === (before.match(WRAPPER_CLOSE_RE)?.length ?? 0)
  )
}

/** only instruction runs stand between the begin run and the separate run */
function onlyInstructionRuns(xml: string): boolean {
  const rest = xml.replace(/<w:r\b[^>]*>([\s\S]*?)<\/w:r>/g, (run, inner: string) =>
    /^\s*<w:instrText\b[^>]*>[\s\S]*?<\/w:instrText>\s*$/.test(
      inner.replace(/<w:rPr>[\s\S]*?<\/w:rPr>|<w:rPr\/>/, ''),
    )
      ? ''
      : run,
  )
  return rest.trim() === ''
}

function findLinkFields(
  elements: readonly { name: string; start: number; end: number }[],
  documentXml: string,
): FoundField[] {
  const found: FoundField[] = []
  const stack: Frame[] = []
  for (const el of elements) {
    const xml = documentXml.slice(el.start, el.end)
    if (!xml.includes('<w:fldChar')) {
      if (el.name !== 'w:p' && xml.includes('<w:instrText'))
        for (const f of stack) f.candidate = false
      continue
    }
    const hostable = el.name === 'w:p' && !HOST_BLOCKED_RE.test(xml)
    TOKEN_RE.lastIndex = 0
    let m: RegExpExecArray | null
    while ((m = TOKEN_RE.exec(xml)) !== null) {
      const top = stack[stack.length - 1]
      if (m[0].startsWith('<w:instrText')) {
        if (top && !top.separated) top.instr += decodeEntities(m[1])
        continue
      }
      const type = /\bw:fldCharType\s*=\s*["'](begin|separate|end)["']/.exec(m[0])?.[1]
      const span = runSpan(xml, m.index)
      const plain =
        hostable && span !== null && bareFieldRun(xml, span) && directChild(xml, span.start)
      if (type === 'begin') {
        // a field begun inside another's instruction is part of that instruction
        if (top && !top.separated) top.candidate = false
        stack.push({
          candidate: plain && stack.length === 0,
          para: el.start,
          beginRun: span ?? { start: 0, end: 0 },
          lock: /\bw:fldLock\s*=\s*["'](?:1|true)["']/.test(m[0]),
          instr: '',
          separated: false,
        })
      } else if (type === 'separate') {
        if (!top || top.separated) continue
        top.separated = true
        top.candidate &&=
          plain &&
          top.para === el.start &&
          onlyInstructionRuns(xml.slice(top.beginRun.end, span!.start))
        if (top.candidate) top.separate = span!
      } else if (type === 'end') {
        const frame = stack.pop()
        if (!frame?.candidate || !frame.separate || !plain || !LINK_FIELD_RE.test(frame.instr))
          continue
        found.push({
          instr: frame.instr,
          lock: frame.lock,
          beginPara: frame.para,
          begin: { start: frame.beginRun.start, end: frame.separate.end },
          endPara: el.start,
          end: span!,
        })
      }
    }
  }
  return found
}

/** a link Paste link made: its result is text, wherever its two ends stand */
function pastedLink(instr: string): boolean {
  const info = parseLinkInstruction(instr)
  return !!info && isWordDocumentLink(info) && info.item.startsWith(LINK_SOURCE_PREFIX)
}

const markField = (instr: string, lock = false): string =>
  `<w:r><w:fldChar w:fldCharType="begin"${lock ? ' w:fldLock="true"' : ''}/></w:r>` +
  `<w:r><w:instrText xml:space="preserve">${escapeXmlText(instr)}</w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

/**
 * The body paragraphs that open or close a link field, rewritten so each end
 * reads as a marker: keyed by the paragraph's offset in the document. A field
 * is left as it is when either of its paragraphs could not stay editable.
 */
export function linkFieldRewrites(
  elements: readonly { name: string; start: number; end: number }[],
  documentXml: string,
): Map<number, string> {
  if (!/INCLUDETEXT|INCLUDEPICTURE|LINK/i.test(documentXml)) return new Map()
  const ends = new Map(elements.map((el) => [el.start, el.end]))
  const hasText = (para: number): boolean =>
    /<w:t(?:\s[^>]*)?>[^<]*\S/.test(documentXml.slice(para, ends.get(para)))
  // a picture or object field alone in its paragraph keeps the placeholder it is shown as
  let fields = findLinkFields(elements, documentXml).filter(
    (f) =>
      /^\s*INCLUDETEXT/i.test(f.instr) ||
      pastedLink(f.instr) ||
      (hasText(f.beginPara) && hasText(f.endPara)),
  )
  for (;;) {
    const edits = new Map<number, { span: Span; xml: string; field: FoundField }[]>()
    const add = (para: number, span: Span, xml: string, field: FoundField) =>
      edits.set(para, [...(edits.get(para) ?? []), { span, xml, field }])
    for (const f of fields) {
      add(f.beginPara, f.begin, markField(`GENOFFICE.LINKBEGIN ${f.instr}`, f.lock), f)
      add(f.endPara, f.end, markField('GENOFFICE.LINKEND'), f)
    }
    const out = new Map<number, string>()
    const failed = new Set<FoundField>()
    for (const [para, list] of edits) {
      let xml = documentXml.slice(para, ends.get(para))
      for (const edit of [...list].sort((a, b) => b.span.start - a.span.start))
        xml = xml.slice(0, edit.span.start) + edit.xml + xml.slice(edit.span.end)
      if (onlyXeFields(xml)) out.set(para, xml)
      else for (const edit of list) failed.add(edit.field)
    }
    if (failed.size === 0) return out
    fields = fields.filter((f) => !failed.has(f))
  }
}

export interface BlockLinkField {
  /** which of the block's link fields, in document order */
  index: number
  instr: string
  locked: boolean
}

interface BlockFieldSpan extends BlockLinkField {
  beginTag: Span
  /** the begin run through the separate run */
  code: Span
  endRun: Span
}

function blockFieldSpans(xml: string): BlockFieldSpan[] {
  const out: BlockFieldSpan[] = []
  const stack: (Omit<BlockFieldSpan, 'index' | 'endRun' | 'code'> & {
    beginRun: Span | null
    separate: Span | null
    separated: boolean
  })[] = []
  TOKEN_RE.lastIndex = 0
  let m: RegExpExecArray | null
  while ((m = TOKEN_RE.exec(xml)) !== null) {
    const top = stack[stack.length - 1]
    if (m[0].startsWith('<w:instrText')) {
      if (top && !top.separated) top.instr += decodeEntities(m[1])
      continue
    }
    const type = /\bw:fldCharType\s*=\s*["'](begin|separate|end)["']/.exec(m[0])?.[1]
    const span = runSpan(xml, m.index)
    const bare = span !== null && bareFieldRun(xml, span)
    if (type === 'begin') {
      stack.push({
        instr: '',
        locked: /\bw:fldLock\s*=\s*["'](?:1|true)["']/.test(m[0]),
        beginTag: { start: m.index, end: m.index + m[0].length },
        beginRun: bare ? span : null,
        separate: null,
        separated: false,
      })
    } else if (type === 'separate') {
      if (!top || top.separated) continue
      top.separated = true
      top.separate = bare ? span : null
    } else if (type === 'end') {
      const frame = stack.pop()
      if (!frame || !LINK_FIELD_RE.test(frame.instr)) continue
      if (!frame.beginRun || !frame.separate || !bare) continue
      out.push({
        index: 0,
        instr: frame.instr,
        locked: frame.locked,
        beginTag: frame.beginTag,
        code: { start: frame.beginRun.start, end: frame.separate.end },
        endRun: span!,
      })
    }
  }
  return out
    .sort((a, b) => a.code.start - b.code.start)
    .map((field, index) => ({ ...field, index }))
}

/** the link fields that begin and end inside one block's XML */
export function blockLinkFields(xml: string): BlockLinkField[] {
  if (!xml.includes('<w:fldChar') || !/INCLUDETEXT|INCLUDEPICTURE|LINK/i.test(xml)) return []
  return blockFieldSpans(xml).map(({ index, instr, locked }) => ({ index, instr, locked }))
}

/** Break Link on a block's field: its code and field characters go, its result stays */
export function breakBlockLinkField(xml: string, index: number): string | null {
  const field = blockFieldSpans(xml)[index]
  if (!field) return null
  return (
    xml.slice(0, field.code.start) +
    xml.slice(field.code.end, field.endRun.start) +
    xml.slice(field.endRun.end)
  )
}

/** a block's link field with another instruction and lock state */
export function editBlockLinkField(
  xml: string,
  index: number,
  change: { instr?: string; locked?: boolean },
): string | null {
  const field = blockFieldSpans(xml)[index]
  if (!field) return null
  const locked = change.locked ?? field.locked
  const tag = xml
    .slice(field.beginTag.start, field.beginTag.end)
    .replace(/\s+w:fldLock\s*=\s*["'][^"']*["']/, '')
    .replace(/\s*\/?>$/, (close) => `${locked ? ' w:fldLock="true"' : ''}${close.trim()}`)
  let code = xml.slice(field.beginTag.end, field.code.end)
  if (change.instr !== undefined) {
    let first = true
    code = code.replace(
      /(<w:instrText\b[^>]*>)[\s\S]*?(<\/w:instrText>)/g,
      (_, open: string, close: string) => {
        const text = first ? escapeXmlText(change.instr!) : ''
        first = false
        return `${open.includes('xml:space') ? open : open.replace('>', ' xml:space="preserve">')}${text}${close}`
      },
    )
  }
  return xml.slice(0, field.beginTag.start) + tag + code + xml.slice(field.code.end)
}

/** the outside file the first picture of a block is linked to, when its blip names one */
export function pictureLinkTarget(
  xml: string,
  rels: ReadonlyMap<string, { target: string; targetMode?: string }>,
): string | null {
  const blip = /<a:blip\b[^>]*>/.exec(xml)?.[0]
  const rId = blip && /\br:link\s*=\s*["']([^"']+)["']/.exec(blip)?.[1]
  const rel = rId ? rels.get(rId) : undefined
  return rel?.targetMode === 'External' && rel.target ? rel.target : null
}

const LINK_TOKEN = 'aidocs-link:'
const LINK_TOKEN_RE = /\br:link="aidocs-link:([^"]*)"/g

/**
 * What a picture's `r:link` holds between an edit and the save when the link
 * is new or names another file: the target itself, which the save turns into
 * an external relationship of the part the picture is in.
 */
export const pictureLinkToken = (target: string): string => LINK_TOKEN + encodeURIComponent(target)

/** the file a picture's drawing was linked to in this session and not saved yet; null when it has none */
export function pendingPictureLink(xml: string): string | null {
  const token = new RegExp(LINK_TOKEN_RE.source).exec(xml)?.[1]
  if (token === undefined) return null
  try {
    return decodeURIComponent(token)
  } catch {
    return null
  }
}

/** every pending picture link of the XML as the relationship `relate` gives its target */
export function resolvePictureLinkTokens(xml: string, relate: (target: string) => string): string {
  if (!xml.includes(LINK_TOKEN)) return xml
  return xml.replace(LINK_TOKEN_RE, (whole, token: string) => {
    try {
      return `r:link="${relate(decodeURIComponent(token))}"`
    } catch {
      return whole
    }
  })
}

/** the first picture of a block linked to the relationship given, or to nothing */
export function withPictureLink(xml: string, rId: string | null): string {
  const blip = /<a:blip\b[^>]*>/.exec(xml)
  if (!blip) return xml
  let tag = blip[0].replace(/\s+r:link\s*=\s*["'][^"']*["']/, '')
  if (rId !== null) tag = tag.replace(/^<a:blip\b/, `<a:blip r:link="${rId}"`)
  return xml.slice(0, blip.index) + tag + xml.slice(blip.index + blip[0].length)
}
