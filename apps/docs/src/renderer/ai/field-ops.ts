import type { Mark, Node as PmNode } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'
import type { Op, OpDef, OpResult, RunEnv, Target, TopBlock } from './ops'
import { posAfterText } from './after-text'

/**
 * Word field ops: insert an inline field with a locally computed (or given)
 * result, bookmark a paragraph for REF / PAGEREF, and refresh the fields the
 * editor can compute. What only Word can compute (page numbers, document
 * properties) is written with the begin fldChar marked dirty, so Word
 * refreshes it on open.
 */

export interface FieldOpHelpers {
  validateShape(op: Op, def: OpDef, where: string): string | null
  matchTarget(doc: PmNode, target: Target, sel: RunEnv['sel']): TopBlock[]
  markChanged(tr: Transaction, pos: number, ctx: RunEnv['ctx']): void
}

const TEXT_BLOCKS = new Set(['docParagraph', 'docHeading', 'docListItem'])
const BLOCKED: Record<string, string> = {
  TOC: 'use insertToc',
  INDEX: 'index fields are not supported',
  XE: 'index entries are not supported',
  HYPERLINK: 'use setFont with link',
  INCLUDEPICTURE: 'use insert_image',
  INCLUDETEXT: 'use insert_content',
  EMBED: 'embedded objects are not supported',
  FORMCHECKBOX: 'form fields are not supported',
  FORMTEXT: 'form fields are not supported',
  FORMDROPDOWN: 'form fields are not supported',
  ADDIN: 'add-in fields are not supported',
}
/** result defaults to "1" and Word recomputes on open */
const COUNTER_FIELDS = new Set([
  'PAGE',
  'NUMPAGES',
  'SECTIONPAGES',
  'SECTION',
  'NUMWORDS',
  'NUMCHARS',
  'PAGEREF',
])
const DATE_FIELDS = new Set(['DATE', 'TIME', 'CREATEDATE', 'SAVEDATE', 'PRINTDATE'])
/** marks that never carry over onto an inserted field result */
const NO_INHERIT = new Set([
  'instrField',
  'refField',
  'comment',
  'ins',
  'del',
  'link',
  'docSym',
  'ctrlCheckbox',
  'revisionOriginal',
  'rprChange',
])
const BOOKMARK_NAME = /^[A-Za-z_][A-Za-z0-9_]{0,39}$/
const SCOPES = ['all', 'SEQ', 'REF', 'DATE'] as const

const MONTHS = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December',
]
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']

interface Resolved {
  instr: string
  result: string
  dirty: boolean
  /** REF fields become refField marks (the editor's cross-reference run) */
  refName?: string
}

function tokens(args: string): string[] {
  return tokensWithQuotes(args).map((t) => t.text)
}

interface Tok {
  text: string
  quoted: boolean
}

function tokensWithQuotes(args: string): Tok[] {
  return [...args.matchAll(/"([^"]*)"|(\S+)/g)].map((m) =>
    m[1] !== undefined ? { text: m[1], quoted: true } : { text: m[2]!, quoted: false },
  )
}

/** value of a `\x` switch: `\@ "yyyy"` → yyyy, `\* ROMAN` → ROMAN */
function switchOf(args: string, letter: string): string | undefined {
  const m = new RegExp(`\\\\${letter === '*' ? '\\*' : letter}\\s*(?:"([^"]*)"|(\\S+))`).exec(args)
  return m ? (m[1] ?? m[2]) : undefined
}

function hasSwitch(args: string, letter: string): boolean {
  return new RegExp(`\\\\${letter}(?=\\s|$)`).test(args)
}

function toRoman(n: number): string {
  const table: Array<[number, string]> = [
    [1000, 'M'],
    [900, 'CM'],
    [500, 'D'],
    [400, 'CD'],
    [100, 'C'],
    [90, 'XC'],
    [50, 'L'],
    [40, 'XL'],
    [10, 'X'],
    [9, 'IX'],
    [5, 'V'],
    [4, 'IV'],
    [1, 'I'],
  ]
  let out = ''
  for (const [value, glyph] of table) {
    while (n >= value) {
      out += glyph
      n -= value
    }
  }
  return out
}

function toAlpha(n: number): string {
  let out = ''
  while (n > 0) {
    n--
    out = String.fromCharCode(65 + (n % 26)) + out
    n = Math.floor(n / 26)
  }
  return out
}

export function formatSeqNumber(n: number, format: string | undefined): string {
  switch (format) {
    case 'ROMAN':
      return toRoman(n)
    case 'roman':
      return toRoman(n).toLowerCase()
    case 'ALPHABETIC':
      return toAlpha(n)
    case 'alphabetic':
      return toAlpha(n).toLowerCase()
    default:
      return String(n)
  }
}

const pad = (n: number, width: number) => String(n).padStart(width, '0')

/** Word date-time picture (yyyy, MMMM, dddd, HH, h, am/pm, 'literal') → text */
export function formatFieldDate(date: Date, picture: string): string {
  const h12 = date.getHours() % 12 || 12
  return picture.replace(
    /yyyy|yy|MMMM|MMM|MM|M|dddd|ddd|dd|d|HH|H|hh|h|mm|m|ss|s|am\/pm|AM\/PM|'[^']*'/g,
    (tok) => {
      switch (tok) {
        case 'yyyy':
          return String(date.getFullYear())
        case 'yy':
          return pad(date.getFullYear() % 100, 2)
        case 'MMMM':
          return MONTHS[date.getMonth()]!
        case 'MMM':
          return MONTHS[date.getMonth()]!.slice(0, 3)
        case 'MM':
          return pad(date.getMonth() + 1, 2)
        case 'M':
          return String(date.getMonth() + 1)
        case 'dddd':
          return DAYS[date.getDay()]!
        case 'ddd':
          return DAYS[date.getDay()]!.slice(0, 3)
        case 'dd':
          return pad(date.getDate(), 2)
        case 'd':
          return String(date.getDate())
        case 'HH':
          return pad(date.getHours(), 2)
        case 'H':
          return String(date.getHours())
        case 'hh':
          return pad(h12, 2)
        case 'h':
          return String(h12)
        case 'mm':
          return pad(date.getMinutes(), 2)
        case 'm':
          return String(date.getMinutes())
        case 'ss':
          return pad(date.getSeconds(), 2)
        case 's':
          return String(date.getSeconds())
        case 'am/pm':
          return date.getHours() < 12 ? 'am' : 'pm'
        case 'AM/PM':
          return date.getHours() < 12 ? 'AM' : 'PM'
        default:
          return tok.slice(1, -1)
      }
    },
  )
}

const DEFAULT_PICTURE: Record<string, string> = { TIME: 'h:mm am/pm' }

function dateResult(type: string, args: string, now: Date): string {
  return formatFieldDate(now, switchOf(args, '@') ?? DEFAULT_PICTURE[type] ?? 'M/d/yyyy')
}

const REF_HEAD_RE = /^\s*(?:REF\s+)?(?:"[^"]+"|[^\s\\]+)/
/** only \h and \* formatting switches leave the result as the bookmark text we can compute here */
function plainRefSwitches(switches: string): boolean {
  return switches.replace(/\\h\b|\\\*\s+\S+/g, '').trim() === ''
}

/**
 * IF a op b "then" "else" with literal operands; undefined when it needs Word.
 * Measured against Word: unquoted numbers compare as numbers, quoted operands as
 * case-sensitive text in alphabetical order, and with = / <> the second operand
 * may hold ? and * wildcards.
 */
export function evaluateIf(args: string): string | undefined {
  const parts = tokensWithQuotes(args)
  if (parts.length < 5) return undefined
  const [ta, top, tb, tyes, tno] = parts as [Tok, Tok, Tok, Tok, Tok]
  const [a, op, b, yes, no] = [ta.text, top.text, tb.text, tyes.text, tno.text]
  const na = Number(a)
  const nb = Number(b)
  const numeric =
    !ta.quoted && !tb.quoted && a !== '' && b !== '' && !Number.isNaN(na) && !Number.isNaN(nb)
  const wildcard = (op === '=' || op === '<>') && /[?*]/.test(b)
  const cmp = numeric ? na - nb : a.localeCompare(b)
  const matches = wildcard
    ? new RegExp(
        `^${b
          .replace(/[.+^${}()|[\]\\]/g, '\\$&')
          .replace(/\*/g, '.*')
          .replace(/\?/g, '.')}$`,
      ).test(a)
    : cmp === 0
  let hit: boolean
  switch (op) {
    case '=':
      hit = matches
      break
    case '<>':
      hit = !matches
      break
    case '<':
      hit = cmp < 0
      break
    case '>':
      hit = cmp > 0
      break
    case '<=':
      hit = cmp <= 0
      break
    case '>=':
      hit = cmp >= 0
      break
    default:
      return undefined
  }
  return hit ? yes : no
}

interface BookmarkHit {
  node: PmNode
  pos: number
}

function bookmarkNames(node: PmNode, attr: 'bookmarks' | 'hiddenBookmarks'): string[] {
  const value = node.attrs?.[attr]
  return Array.isArray(value) ? (value as string[]) : []
}

function findBookmark(doc: PmNode, name: string): BookmarkHit | null {
  let hit: BookmarkHit | null = null
  doc.descendants((node, pos) => {
    if (hit) return false
    if (
      bookmarkNames(node, 'bookmarks').includes(name) ||
      bookmarkNames(node, 'hiddenBookmarks').includes(name)
    ) {
      hit = { node, pos }
      return false
    }
    return true
  })
  return hit
}

function userBookmarks(doc: PmNode): string[] {
  const names: string[] = []
  doc.descendants((node) => {
    names.push(...bookmarkNames(node, 'bookmarks'))
    return true
  })
  return names
}

/** bookmark text as shown, minus deleted runs and REF results that point back at `excludeRef` (a REF inside its own bookmark would otherwise grow on every refresh) */
function liveTextOf(node: PmNode, excludeRef?: string): string {
  let out = ''
  node.descendants((child) => {
    if (!child.isText) return true
    const skip = child.marks.some(
      (m) =>
        m.type.name === 'del' ||
        (excludeRef !== undefined && m.type.name === 'refField' && m.attrs.name === excludeRef),
    )
    if (!skip) out += child.text ?? ''
    return true
  })
  return out
}

function markOf(node: PmNode, name: string): Mark | undefined {
  return node.marks.find((m) => m.type.name === name)
}

function seqIdentifier(args: string): string | undefined {
  return tokens(args.replace(/\\\S+(?:\s+"[^"]*"|\s+\S+)?/g, ' '))[0]
}

function keywordOf(instr: string): string {
  return /^\s*([A-Za-z]+)/.exec(instr)?.[1]?.toUpperCase() ?? ''
}

function argsOf(instr: string): string {
  return instr.replace(/^\s*[A-Za-z]+\s*/, '')
}

/** SEQ fields with this identifier the editor can see before `beforePos` */
function seqCountBefore(doc: PmNode, id: string, beforePos: number): number {
  return fieldRuns(doc).filter((run) => {
    if (run.pos >= beforePos || !run.instr) return false
    const instr = String(run.instr.attrs.instr)
    return keywordOf(instr) === 'SEQ' && seqIdentifier(argsOf(instr)) === id
  }).length
}

function resolveField(op: Op, doc: PmNode, pos: number, now = new Date()): Resolved {
  const type = String(op.type).toUpperCase()
  const args = typeof op.args === 'string' ? op.args.trim() : ''
  const given = typeof op.result === 'string' ? op.result : undefined
  const instr = args ? `${type} ${args}` : type
  const needsResult = (): string => {
    if (given !== undefined) return given
    throw new Error(
      `insertField: ${type} cannot be computed here; pass result (the text Word shows until it updates the field)`,
    )
  }
  if (type === 'SEQ') {
    const id = seqIdentifier(args)
    if (!id)
      throw new Error('insertField: SEQ needs an identifier in args, e.g. "Figure \\* ARABIC"')
    const restart = switchOf(args, 'r')
    const number = restart !== undefined ? Number(restart) : seqCountBefore(doc, id, pos) + 1
    return {
      instr,
      result: given ?? formatSeqNumber(number, switchOf(args, '*')),
      dirty: true,
    }
  }
  if (type === 'REF' || type === 'PAGEREF') {
    const name = tokens(args)[0]
    if (!name || name.startsWith('\\'))
      throw new Error(`insertField: ${type} needs a bookmark name in args`)
    const target = findBookmark(doc, name)
    if (!target) {
      const known = userBookmarks(doc)
      throw new Error(
        `insertField: no bookmark named "${name}"${known.length ? `; bookmarks in the document: ${known.slice(0, 20).join(', ')}` : '; the document has no bookmarks (insertBookmark adds one)'}`,
      )
    }
    if (type === 'PAGEREF') return { instr, result: given ?? '1', dirty: true }
    const plainRef = plainRefSwitches(args.replace(REF_HEAD_RE, ''))
    const text = given ?? (liveTextOf(target.node, name).trim() || name)
    return plainRef
      ? { instr, result: text, dirty: false, refName: name }
      : { instr, result: text, dirty: true }
  }
  if (DATE_FIELDS.has(type)) {
    const computable = type === 'DATE' || type === 'TIME'
    return {
      instr,
      result: given ?? dateResult(type, args, now),
      dirty: !computable,
    }
  }
  if (type === 'MERGEFIELD') {
    const name = tokens(args)[0]
    if (!name) throw new Error('insertField: MERGEFIELD needs the merge field name in args')
    return { instr, result: given ?? `«${name}»`, dirty: false }
  }
  if (type === 'IF') {
    const value = evaluateIf(args)
    if (value !== undefined && given === undefined) return { instr, result: value, dirty: false }
    return { instr, result: needsResult(), dirty: true }
  }
  if (COUNTER_FIELDS.has(type)) return { instr, result: given ?? '1', dirty: true }
  return { instr, result: needsResult(), dirty: true }
}

/** position inside the block where the field goes: end, start, or right after `afterText` */
function insertionPos(b: TopBlock, op: Op): number {
  const contentStart = b.pos + 1
  const contentEnd = b.pos + b.node.nodeSize - 1
  if (typeof op.afterText !== 'string') return op.position === 'start' ? contentStart : contentEnd
  const where = posAfterText(b.node, contentStart, op.afterText, `block ${b.index}`)
  if ('error' in where) throw new Error(`insertField: ${where.error}`)
  return where.pos
}

function inheritedMarks(doc: PmNode, pos: number): readonly Mark[] {
  const $pos = doc.resolve(pos)
  const neighbour = $pos.nodeBefore ?? $pos.nodeAfter
  if (!neighbour?.isText) return []
  return neighbour.marks.filter((m) => !NO_INHERIT.has(m.type.name))
}

function singleTextBlock(op: Op, env: RunEnv, h: FieldOpHelpers, name: string): TopBlock {
  const blocks = h.matchTarget(env.tr.doc, op.target as Target, env.sel)
  if (blocks.length !== 1) {
    throw new Error(
      `${name}: target must match exactly one paragraph (matched ${blocks.length}); use blockIndexes`,
    )
  }
  const b = blocks[0]!
  if (!TEXT_BLOCKS.has(b.node.type.name)) {
    throw new Error(
      `${name}: block ${b.index} is ${b.node.type.name === 'docProtected' ? 'protected' : `a ${b.node.type.name}`}; target a paragraph, heading or list item`,
    )
  }
  return b
}

function runInsertField(op: Op, env: RunEnv, h: FieldOpHelpers): OpResult {
  const { tr, schema } = env
  const b = singleTextBlock(op, env, h, 'insertField')
  const pos = insertionPos(b, op)
  const r = resolveField(op, tr.doc, pos)
  const fieldMark = r.refName
    ? schema.marks.refField!.create({ name: r.refName, instr: ` ${r.instr.trim()} `, dirty: false })
    : schema.marks.instrField!.create({ instr: r.instr, dirty: r.dirty })
  // an empty result is legal for a field but not for a text node: Word's own convention is a space
  tr.insert(pos, schema.text(r.result || ' ', [...inheritedMarks(tr.doc, pos), fieldMark]))
  h.markChanged(tr, b.pos, env.ctx)
  return {
    op: 'insertField',
    matched: 1,
    changed: 1,
    skippedProtected: 0,
    detail: `{ ${r.instr} } = "${r.result}"${r.dirty ? ' (Word recomputes on open)' : ''}`,
  }
}

function runInsertBookmark(op: Op, env: RunEnv, h: FieldOpHelpers): OpResult {
  const { tr } = env
  const name = String(op.name)
  const b = singleTextBlock(op, env, h, 'insertBookmark')
  if (findBookmark(tr.doc, name))
    throw new Error(`insertBookmark: bookmark "${name}" already exists`)
  tr.setNodeMarkup(b.pos, undefined, {
    ...b.node.attrs,
    bookmarks: [...bookmarkNames(b.node, 'bookmarks'), name],
  })
  h.markChanged(tr, b.pos, env.ctx)
  return { op: 'insertBookmark', matched: 1, changed: 1, skippedProtected: 0, detail: name }
}

interface FieldRun {
  pos: number
  end: number
  text: string
  marks: readonly Mark[]
  /** the result spans several text nodes (formatting applied to part of it) */
  split: boolean
  instr: Mark | undefined
  ref: Mark | undefined
}

/** one entry per field: adjacent text nodes carrying an equal field mark are the same field */
function fieldRuns(doc: PmNode): FieldRun[] {
  const out: FieldRun[] = []
  doc.descendants((node, pos) => {
    if (!node.isText) return true
    const instr = markOf(node, 'instrField')
    const ref = markOf(node, 'refField')
    const mark = instr ?? ref
    if (!mark) return true
    const prev = out[out.length - 1]
    if (prev && prev.end === pos && (prev.instr ?? prev.ref)!.eq(mark)) {
      prev.end += node.nodeSize
      prev.text += node.text ?? ''
      prev.split = true
    } else {
      out.push({
        pos,
        end: pos + node.nodeSize,
        text: node.text ?? '',
        marks: node.marks,
        split: false,
        instr,
        ref,
      })
    }
    return true
  })
  return out
}

function runUpdateFields(op: Op, env: RunEnv, h: FieldOpHelpers, now = new Date()): OpResult {
  const { tr, schema } = env
  const scope = (op.scope as (typeof SCOPES)[number] | undefined) ?? 'all'
  const runs = fieldRuns(tr.doc)
  const seq = new Map<string, number>()
  let recomputed = 0
  let dirtied = 0
  let missing = 0
  const retag = (marks: readonly Mark[], dirty: boolean | null) =>
    marks.map((m) =>
      (m.type.name === 'instrField' || m.type.name === 'refField') && dirty !== null
        ? m.type.create({ ...m.attrs, dirty })
        : m,
    )
  const rewrite = (run: FieldRun, value: string, dirty: boolean | null) => {
    const text = value || ' '
    const from = tr.mapping.map(run.pos)
    const to = tr.mapping.map(run.end)
    const wasDirty = (run.instr ?? run.ref)!.attrs.dirty === true
    if (text === run.text) {
      if (dirty === null || dirty === wasDirty) return
      // same result: flip the flag on every node so formatting applied to part of it stays put
      const nodes: Array<{ pos: number; node: PmNode }> = []
      tr.doc.nodesBetween(from, to, (node, pos) => {
        if (node.isText) nodes.push({ pos, node })
      })
      for (const { pos, node } of nodes)
        tr.replaceWith(pos, pos + node.nodeSize, schema.text(node.text!, retag(node.marks, dirty)))
    } else {
      tr.replaceWith(from, to, schema.text(text, retag(run.marks, dirty)))
    }
    h.markChanged(tr, tr.doc.resolve(from).before(1), env.ctx)
  }
  for (const run of runs) {
    if (run.ref && !run.instr) {
      if (scope !== 'all' && scope !== 'REF') continue
      const name = String(run.ref.attrs.name)
      const instr = run.ref.attrs.instr as string | null
      if (instr && !plainRefSwitches(instr.replace(REF_HEAD_RE, ''))) {
        // \p / \n / \r … results are Word's: keep the cached text, make sure Word refreshes it
        if (run.ref.attrs.dirty !== true) {
          rewrite(run, run.text, true)
          dirtied++
        }
        continue
      }
      const target = findBookmark(tr.doc, name)
      if (!target) {
        missing++
        continue
      }
      rewrite(run, liveTextOf(target.node, name).trim() || run.text, false)
      recomputed++
      continue
    }
    const instr = String(run.instr!.attrs.instr)
    const keyword = keywordOf(instr)
    const args = argsOf(instr)
    if (keyword === 'SEQ') {
      const id = seqIdentifier(args)
      if (!id) continue
      const restart = switchOf(args, 'r')
      const value =
        restart !== undefined
          ? Number(restart)
          : hasSwitch(args, 'c')
            ? (seq.get(id) ?? 1)
            : (seq.get(id) ?? 0) + 1
      seq.set(id, value)
      if (scope !== 'all' && scope !== 'SEQ') continue
      rewrite(run, formatSeqNumber(value, switchOf(args, '*')), true)
      recomputed++
    } else if (keyword === 'DATE' || keyword === 'TIME') {
      if (scope !== 'all' && scope !== 'DATE') continue
      rewrite(run, dateResult(keyword, args, now), false)
      recomputed++
    } else if (
      scope === 'all' &&
      keyword !== 'MERGEFIELD' &&
      !/^(?:ADDIN|FORMCHECKBOX)$/.test(keyword) &&
      run.instr!.attrs.dirty !== true &&
      !run.instr!.attrs.beginXml
    ) {
      rewrite(run, run.text, true)
      dirtied++
    }
  }
  const parts = [`${recomputed} recomputed`]
  if (dirtied) parts.push(`${dirtied} marked for Word to recompute on open`)
  if (missing) parts.push(`${missing} REF field(s) point at a missing bookmark`)
  return {
    op: 'updateFields',
    matched: runs.length,
    changed: recomputed + dirtied,
    skippedProtected: 0,
    detail: parts.join(', '),
  }
}

export function fieldOpDefs(h: FieldOpHelpers): OpDef[] {
  const insertField: OpDef = {
    name: 'insertField',
    signature:
      '{ op: "insertField", target, type: "SEQ"|"REF"|"PAGEREF"|"DATE"|"TIME"|"MERGEFIELD"|"IF"|"PAGE"|"NUMPAGES"|"AUTHOR"|..., args?: string, result?: string, position?: "start"|"end", afterText?: string }  // insert a Word field into exactly one paragraph, at its end unless position or afterText says otherwise; args is the instruction after the keyword ({ SEQ Figure \\* ARABIC } = type "SEQ", args "Figure \\* ARABIC"). Computed here: SEQ (visible fields with that identifier so far + 1), REF/PAGEREF (the bookmarked paragraph; insertBookmark creates one), DATE/TIME (now, \\@ picture), MERGEFIELD («Name»), IF with literal operands; PAGE/NUMPAGES show "1"; anything else needs result, the text shown until Word updates it. Fields only Word can compute are marked dirty and refresh when the file opens in Word',
    keys: ['type', 'args', 'result', 'position', 'afterText'],
    target: 'required',
    validate(op, where) {
      const shape = h.validateShape(op, insertField, where)
      if (shape) return shape
      if (typeof op.type !== 'string' || !/^[A-Za-z]+$/.test(op.type))
        return `${where}: type must be a field keyword such as SEQ, REF or DATE`
      const blocked = BLOCKED[op.type.toUpperCase()]
      if (blocked)
        return `${where}: ${op.type.toUpperCase()} fields cannot be inserted this way; ${blocked}`
      if (op.args !== undefined && typeof op.args !== 'string')
        return `${where}: args must be a string`
      if (op.result !== undefined && (typeof op.result !== 'string' || op.result === ''))
        return `${where}: result must be a non-empty string`
      if (op.position !== undefined && op.position !== 'start' && op.position !== 'end')
        return `${where}: position must be "start" or "end"`
      if (op.afterText !== undefined && (typeof op.afterText !== 'string' || op.afterText === ''))
        return `${where}: afterText must be a non-empty string`
      if (op.afterText !== undefined && op.position !== undefined)
        return `${where}: give either position or afterText, not both`
      return null
    },
    apply: (op, env) => runInsertField(op, env, h),
  }
  const insertBookmark: OpDef = {
    name: 'insertBookmark',
    signature:
      '{ op: "insertBookmark", target, name }  // bookmark exactly one paragraph (name: letters, digits, underscore, up to 40 characters, unique) so REF / PAGEREF fields can point at it',
    keys: ['name'],
    target: 'required',
    validate(op, where) {
      const shape = h.validateShape(op, insertBookmark, where)
      if (shape) return shape
      if (typeof op.name !== 'string' || !BOOKMARK_NAME.test(op.name))
        return `${where}: name must start with a letter or underscore and use only letters, digits and underscores (max 40)`
      return null
    },
    apply: (op, env) => runInsertBookmark(op, env, h),
  }
  const updateFields: OpDef = {
    name: 'updateFields',
    signature:
      '{ op: "updateFields", scope?: "all"|"SEQ"|"REF"|"DATE" }  // recompute the inline fields the editor can: SEQ numbering in document order, REF results from their bookmarks, DATE/TIME to now; with scope "all" every other field (PAGE, NUMPAGES, AUTHOR...) is marked dirty so Word recomputes it on open. Table-of-contents blocks are untouched (insertToc rebuilds one)',
    keys: ['scope'],
    target: 'none',
    validate(op, where) {
      const shape = h.validateShape(op, updateFields, where)
      if (shape) return shape
      if (op.scope !== undefined && !SCOPES.includes(op.scope as (typeof SCOPES)[number]))
        return `${where}: scope must be one of ${SCOPES.join(', ')}`
      return null
    },
    apply: (op, env) => runUpdateFields(op, env, h),
  }
  return [insertField, insertBookmark, updateFields]
}
