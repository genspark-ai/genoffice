/**
 * Word's AutoCorrect / AutoFormat-as-you-type: the replace table, capitals,
 * smart quotes, "--" to dashes, automatic lists, fractions, superscript
 * ordinals, *bold* and _italic_, and on Enter border lines and +--+ tables.
 *
 * Only real typed text reaches handleTextInput, so paste, drop, AI edits and
 * IME composition never trigger a rule. Each correction records an undo entry
 * that one Backspace right afterwards reverts, as in Word.
 */
import { Extension, type Editor } from '@tiptap/core'
import {
  DOMParser as PmDOMParser,
  type Fragment,
  type Mark,
  type Node as PmNode,
  type ResolvedPos,
} from '@tiptap/pm/model'
import {
  Plugin,
  PluginKey,
  TextSelection,
  type EditorState,
  type Transaction,
} from '@tiptap/pm/state'
import { keymap } from '@tiptap/pm/keymap'
import { closeHistory } from '@tiptap/pm/history'
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view'
import type { StyleInfo, StyleUpsert } from '@genoffice/docx-engine'
import { t } from '../i18n/locale'
import { headingStyleId } from '../style-gallery'
import type { AutoTextEntry } from './auto-text'
import { dateStyleOnEnter, definedStyleOnEnter } from './autoformat-styles'
import { effectiveSizeHalfPoints } from './text-style-resolve'
import {
  autocorrectEnabled,
  autocorrectLists,
  setAutocorrectList,
  setAutocorrectPref,
  type AutocorrectRule,
  type ReplaceEntry,
} from '../autocorrect-pref'
import { spellcheckEnabled } from '../spellcheck-pref'

export interface AutocorrectListHooks {
  /** numId for a typed "- " / "* " bullet (document definition or a new one) */
  bulletNumId: () => string | null
  /** numId for a typed "1." / "a)" / "(1)" list, starting at the typed value */
  numberedNumId: (numFmt: string, pattern: string, start: number) => string | null
  /** a one-row table in place of the caret's typed +---+---+ line, as one undo step */
  insertTable?: (cols: number) => void
  /** the saved AutoText entries the AutoComplete tip offers */
  autoText?: () => AutoTextEntry[]
  /** puts an entry of more than one paragraph in at the selection */
  insertAutoText?: (entry: AutoTextEntry) => void
  /** Tools > AutoCorrect Options…, from the smart button's menu */
  openOptions?: () => void
  /** a style definition to write with the document: one a rule applies and the document lacks */
  defineStyle?: (up: StyleUpsert) => void
}

export interface AutocorrectStorage {
  hooks: AutocorrectListHooks | null
}

interface Undoable {
  transform: Transaction
  from: number
  to: number
  text: string
}

export const autocorrectKey = new PluginKey<Undoable | null>('autocorrect')

declare module '@tiptap/core' {
  interface Storage {
    autocorrect: AutocorrectStorage
  }
}

interface Ctx {
  state: EditorState
  $from: ResolvedPos
  from: number
  /** the text just typed, not yet in the document */
  text: string
  /** paragraph text before the caret; inline atoms show as U+FFFC */
  before: string
  hooks: AutocorrectListHooks | null
}

type Rule = (ctx: Ctx) => Transaction | null

const caretMarks = (ctx: Ctx): readonly Mark[] => ctx.state.storedMarks ?? ctx.$from.marks()

/** insert the typed text with the marks the browser would have given it */
function insertTyped(tr: Transaction, ctx: Ctx, at: number): Transaction {
  return tr.insert(at, ctx.state.schema.text(ctx.text, caretMarks(ctx)))
}

/** replace the last `count` characters before the caret with `text` (the typed text is dropped) */
function replaceTail(ctx: Ctx, count: number, text: string): Transaction {
  const tr = ctx.state.tr
  tr.replaceWith(ctx.from - count, ctx.from, ctx.state.schema.text(text, caretMarks(ctx)))
  return tr
}

// ---- smart quotes ----

const QUOTE_OPENS_AFTER = /[\s([{<\u2018\u201c\u00ab\u2013\u2014\u3000]/

const smartQuotes: Rule = (ctx) => {
  if (ctx.text !== '"' && ctx.text !== "'") return null
  const prev = ctx.before.slice(-1)
  const opening = prev === '' || QUOTE_OPENS_AFTER.test(prev)
  const glyph = ctx.text === '"' ? (opening ? '\u201c' : '\u201d') : opening ? '\u2018' : '\u2019'
  return ctx.state.tr.insert(ctx.from, ctx.state.schema.text(glyph, caretMarks(ctx)))
}

// ---- hyphens to dashes ----

// "word--word " becomes an em dash, "word - word " an en dash (both fire on the space
// that completes the second word, like Word)
const EM_DASH = /[^\s-]--([^\s-]+)$/
const EN_DASH = /\S - (\S+)$/

const dashes: Rule = (ctx) => {
  if (ctx.text !== ' ') return null
  const em = EM_DASH.exec(ctx.before)
  if (em) {
    const at = ctx.from - em[1].length - 2
    const tr = ctx.state.tr.replaceWith(
      at,
      at + 2,
      ctx.state.schema.text('\u2014', ctx.$from.marks()),
    )
    return insertTyped(tr, ctx, tr.mapping.map(ctx.from))
  }
  const en = EN_DASH.exec(ctx.before)
  if (en) {
    const at = ctx.from - en[1].length - 2
    const tr = ctx.state.tr.replaceWith(
      at,
      at + 1,
      ctx.state.schema.text('\u2013', ctx.$from.marks()),
    )
    return insertTyped(tr, ctx, tr.mapping.map(ctx.from))
  }
  return null
}

// ---- automatic lists ----

interface ListSpec {
  kind: 'bullet' | 'ordered'
  numFmt: string
  pattern: string
  start: number
}

const CJK_NUMERALS = '\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341'

/** the list Word would start for this paragraph text (typed before the space) */
export function parseListPrefix(text: string): ListSpec | null {
  if (/^[-*\u2022]$/.test(text)) return { kind: 'bullet', numFmt: 'bullet', pattern: '', start: 1 }
  let m = /^(\d{1,4})([.)\u3001])$/.exec(text)
  if (m) {
    const start = Number(m[1])
    return start >= 1 ? { kind: 'ordered', numFmt: 'decimal', pattern: `%1${m[2]}`, start } : null
  }
  m = /^\((\d{1,4})\)$/.exec(text)
  if (m) {
    const start = Number(m[1])
    return start >= 1 ? { kind: 'ordered', numFmt: 'decimal', pattern: '(%1)', start } : null
  }
  m = /^([a-z])([.)])$/.exec(text)
  if (m) {
    if (m[1] === 'i')
      return { kind: 'ordered', numFmt: 'lowerRoman', pattern: `%1${m[2]}`, start: 1 }
    const start = m[1].charCodeAt(0) - 96
    return { kind: 'ordered', numFmt: 'lowerLetter', pattern: `%1${m[2]}`, start }
  }
  m = /^([A-Z])([.)])$/.exec(text)
  if (m) {
    if (m[1] === 'I')
      return { kind: 'ordered', numFmt: 'upperRoman', pattern: `%1${m[2]}`, start: 1 }
    const start = m[1].charCodeAt(0) - 64
    return { kind: 'ordered', numFmt: 'upperLetter', pattern: `%1${m[2]}`, start }
  }
  m = new RegExp(`^([${CJK_NUMERALS}])\u3001$`).exec(text)
  if (m) {
    const start = CJK_NUMERALS.indexOf(m[1]) + 1
    return { kind: 'ordered', numFmt: 'chineseCountingThousand', pattern: '%1\u3001', start }
  }
  return null
}

function insideProtected($pos: ResolvedPos): boolean {
  for (let d = $pos.depth; d > 0; d--) if ($pos.node(d).type.name === 'docProtected') return true
  return false
}

const autoList = (ctx: Ctx, kind: ListSpec['kind']): Transaction | null => {
  if (ctx.text !== ' ') return null
  const { $from, state } = ctx
  const listType = state.schema.nodes.docListItem
  if (!listType || $from.parent.type.name !== 'docParagraph') return null
  if ($from.parentOffset !== $from.parent.content.size || insideProtected($from)) return null
  const spec = parseListPrefix(ctx.before)
  if (!spec || spec.kind !== kind) return null
  // allocated up front like the ribbon does; the save drops any pending
  // definition no list item references (Backspace revert, undo)
  const numId =
    spec.kind === 'bullet'
      ? (ctx.hooks?.bulletNumId() ?? null)
      : (ctx.hooks?.numberedNumId(spec.numFmt, spec.pattern, spec.start) ?? null)
  const start = $from.start()
  const tr = state.tr.delete(start, ctx.from)
  tr.setBlockType(start, start, listType, { kind: spec.kind, numId, ilvl: 0 })
  return tr
}

// ---- replace text as you type ----

const WORD_CHAR = /[\p{L}\p{N}]/u
const DELIMITER = /^[\s.,;:!?)\]"']$/

const endsWithFold = (text: string, tail: string): boolean =>
  text.length >= tail.length && text.slice(-tail.length).toLowerCase() === tail.toLowerCase()

/** the character in front of the last `length` characters of `text` */
const charBefore = (text: string, length: number): string => text.slice(-length - 1, -length)

/**
 * The longest table entry the typed text completes. An entry ending in a letter
 * or digit ("teh") waits for the delimiter after the word; any other ("(c)",
 * "-->") fires on its own last character.
 */
interface ReplaceMatch {
  from: string
  to: string
  word: boolean
  /** the replacement kept as formatted text */
  html?: string
}

export function matchReplaceEntry(
  entries: ReplaceEntry[],
  before: string,
  text: string,
): ReplaceMatch | null {
  const typed = before + text
  let best: ReplaceMatch | null = null
  for (const [from, to, html] of entries) {
    if (!from || (best && from.length <= best.from.length)) continue
    const word = WORD_CHAR.test(from.slice(-1))
    if (word) {
      if (!DELIMITER.test(text) || !endsWithFold(before, from)) continue
      if (WORD_CHAR.test(charBefore(before, from.length))) continue
      const matched = before.slice(-from.length)
      const capital = from === from.toLowerCase() && matched[0] !== matched[0].toLowerCase()
      best = { from, to: capital ? to[0].toUpperCase() + to.slice(1) : to, word, html }
    } else {
      if (!endsWithFold(typed, from)) continue
      // "...." and "--->" are not an ellipsis or an arrow with a stray character
      if (charBefore(typed, from.length) === from[0]) continue
      best = { from, to, word, html }
    }
  }
  return best
}

/** the inline content of clipboard HTML that holds one paragraph; null for anything longer */
function inlineContent(state: EditorState, html: string): Fragment | null {
  const holder = document.createElement('div')
  holder.innerHTML = html
  const { content } = PmDOMParser.fromSchema(state.schema).parseSlice(holder, {
    preserveWhitespace: true,
  })
  if (content.childCount !== 1) return content.firstChild?.isInline ? content : null
  const only = content.firstChild!
  return only.isTextblock ? only.content : only.isInline ? content : null
}

const replaceText: Rule = (ctx) => {
  const hit = matchReplaceEntry(autocorrectLists().replace, ctx.before, ctx.text)
  if (!hit) return null
  const count = hit.word ? hit.from.length : hit.from.length - ctx.text.length
  const formatted = hit.html ? inlineContent(ctx.state, hit.html) : null
  const tr = formatted?.size
    ? ctx.state.tr.replaceWith(ctx.from - count, ctx.from, formatted)
    : replaceTail(ctx, count, hit.to)
  return hit.word ? insertTyped(tr, ctx, tr.mapping.map(ctx.from)) : tr
}

// ---- suggestions from the spelling checker ----

const SPELL_WORD = /(?<![\p{L}\p{N}'\u2019])\p{L}[\p{L}'\u2019]*$/u

/**
 * The word a delimiter just ended becomes the spelling checker's suggestion
 * when it is misspelled and the checker offers exactly one. As in Word it
 * follows "check spelling as you type"; words in capitals, deleted text and
 * text marked as not to be proofed are left alone.
 */
const spellSuggest: Rule = (ctx) => {
  if (!DELIMITER.test(ctx.text) || !spellcheckEnabled()) return null
  const word = SPELL_WORD.exec(ctx.before)?.[0]
  if (!word || word.length < 2 || word === word.toUpperCase()) return null
  const unchecked = ctx.$from
    .marks()
    .some((m) => m.type.name === 'del' || (m.type.name === 'docTextStyle' && m.attrs.noProof))
  if (unchecked) return null
  const api = window.desktop
  if (!api?.spellIsMisspelled?.(word)) return null
  const offered = api.spellSuggestions?.(word) ?? []
  if (offered.length !== 1 || !offered[0] || offered[0] === word) return null
  const tr = replaceTail(ctx, word.length, offered[0])
  return insertTyped(tr, ctx, tr.mapping.map(ctx.from))
}

// "\alpha " in body text; the AutoCorrect table above wins a conflict, as in Word
const mathOutside: Rule = (ctx) => {
  if (ctx.text !== ' ') return null
  const token = /(?:^|\s)(\S+)$/.exec(ctx.before)?.[1]
  if (!token) return null
  const to = autocorrectLists().math.find(([from]) => from === token)?.[1]
  if (to === undefined) return null
  const tr = replaceTail(ctx, token.length, to)
  return insertTyped(tr, ctx, tr.mapping.map(ctx.from))
}

// ---- fractions ----

const FRACTIONS: Record<string, string> = { '1/2': '\u00bd', '1/4': '\u00bc', '3/4': '\u00be' }

const fractions: Rule = (ctx) => {
  if (ctx.text !== ' ') return null
  const m = /(^|[^\d/])([13]\/[24])$/.exec(ctx.before)
  const glyph = m ? FRACTIONS[m[2]] : undefined
  if (!glyph) return null
  const tr = replaceTail(ctx, 3, glyph)
  return insertTyped(tr, ctx, tr.mapping.map(ctx.from))
}

// ---- *bold* and _italic_ ----

const boldItalic: Rule = (ctx) => {
  if (ctx.text !== '*' && ctx.text !== '_') return null
  const markType = ctx.state.schema.marks[ctx.text === '*' ? 'bold' : 'italic']
  if (!markType) return null
  const ch = `\\${ctx.text}`
  const m = new RegExp(
    `(^|\\s)${ch}([^${ch}\\s\ufffc](?:[^${ch}\ufffc]*[^${ch}\\s\ufffc])?)$`,
  ).exec(ctx.before)
  if (!m) return null
  const open = ctx.from - m[2].length - 1
  const tr = ctx.state.tr.delete(open, open + 1)
  tr.addMark(open, open + m[2].length, markType.create())
  // what is typed next is plain again
  return tr.setStoredMarks(caretMarks(ctx).filter((mark) => mark.type !== markType))
}

// ---- ordinals ----

function ordinalSuffix(n: number): string {
  const mod100 = n % 100
  if (mod100 >= 11 && mod100 <= 13) return 'th'
  const mod10 = n % 10
  return mod10 === 1 ? 'st' : mod10 === 2 ? 'nd' : mod10 === 3 ? 'rd' : 'th'
}

const ordinals: Rule = (ctx) => {
  if (ctx.text !== ' ') return null
  const m = /(^|[^\w])(\d+)(st|nd|rd|th)$/.exec(ctx.before)
  if (!m || ordinalSuffix(Number(m[2])) !== m[3]) return null
  const styleType = ctx.state.schema.marks.docTextStyle
  if (!styleType) return null
  const marks = caretMarks(ctx)
  const existing = marks.find((mk) => mk.type === styleType)
  const sup = styleType.create({ ...(existing?.attrs ?? {}), vertAlign: 'superscript' })
  const tr = ctx.state.tr.addMark(ctx.from - 2, ctx.from, sup)
  return insertTyped(tr, ctx, ctx.from)
}

// ---- capitals ----

const lowerSet = (words: string[]): Set<string> => new Set(words.map((w) => w.toLowerCase()))

function capitalizeAt(ctx: Ctx, at: number, letter: string): Transaction {
  const tr = ctx.state.tr.insertText(letter.toUpperCase(), at, at + 1)
  return insertTyped(tr, ctx, ctx.from)
}

/** the first paragraph of a table cell: its first word answers to its own switch */
function firstInCell($pos: ResolvedPos): boolean {
  const cell = $pos.depth > 1 ? $pos.node($pos.depth - 1).type.name : ''
  return (cell === 'docTableCell' || cell === 'docTableHeader') && $pos.index($pos.depth - 1) === 0
}

/** the lowercase word the typed delimiter ends, and the text in front of it */
function endedWord(ctx: Ctx): { word: string; prefix: string } | null {
  if (!/^[ .,;:!?]$/.test(ctx.text)) return null
  const word = /(\S+)$/.exec(ctx.before)?.[1]
  if (!word || !/^[a-z]/.test(word) || /[@/.:\\]/.test(word)) return null
  return { word, prefix: ctx.before.slice(0, -word.length) }
}

const capitalize: Rule = (ctx) => {
  const ended = endedWord(ctx)
  if (!ended) return null
  const { word, prefix } = ended
  if (prefix === '') {
    if (firstInCell(ctx.$from)) return null
  } else {
    const end = /(\S+)([.!?])\s+$/.exec(prefix)
    if (!end) return null
    const token = end[1]
    // "e.g. apples", "3. apples", "J. smith": a period that isn't a sentence end
    if (
      end[2] === '.' &&
      (lowerSet(autocorrectLists().firstLetter).has(`${token.toLowerCase()}.`) ||
        /^\d+$/.test(token) ||
        token.length === 1)
    )
      return null
  }
  return capitalizeAt(ctx, ctx.from - word.length, word[0])
}

const capitalizeCells: Rule = (ctx) => {
  const ended = endedWord(ctx)
  if (!ended || ended.prefix !== '' || !firstInCell(ctx.$from)) return null
  return capitalizeAt(ctx, ctx.from - ended.word.length, ended.word[0])
}

const DAY_NAME = /(^|[^\p{L}\p{N}])(monday|tuesday|wednesday|thursday|friday|saturday|sunday)$/u

const capitalizeDays: Rule = (ctx) => {
  if (!DELIMITER.test(ctx.text)) return null
  const m = DAY_NAME.exec(ctx.before)
  return m ? capitalizeAt(ctx, ctx.from - m[2].length, m[2][0]) : null
}

// "THe", but a plural acronym ("PCs") is spelled that way on purpose
const TWO_CAPS = /(^|[^\p{L}\p{N}])([A-Z]{2}(?!s$)[a-z]+)$/u

const twoInitialCaps: Rule = (ctx) => {
  if (!DELIMITER.test(ctx.text)) return null
  const word = TWO_CAPS.exec(ctx.before)?.[2]
  if (!word || autocorrectLists().initialCaps.includes(word)) return null
  const at = ctx.from - word.length + 1
  const tr = ctx.state.tr.insertText(word[1].toLowerCase(), at, at + 1)
  return insertTyped(tr, ctx, ctx.from)
}

// ---- East Asian typing ----

const EAST_ASIAN = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}]/u
const WESTERN = /[A-Za-z0-9]/

// the page already opens a gap between the two scripts; a typed space doubles it
const asianSpaces: Rule = (ctx) => {
  const m = /(\S) $/u.exec(ctx.before)
  if (!m) return null
  const apart =
    (EAST_ASIAN.test(m[1]) && WESTERN.test(ctx.text)) ||
    (WESTERN.test(m[1]) && EAST_ASIAN.test(ctx.text))
  if (!apart) return null
  const tr = ctx.state.tr.delete(ctx.from - 1, ctx.from)
  return insertTyped(tr, ctx, tr.mapping.map(ctx.from))
}

const BRACKETS: Array<[open: string, close: string, wideOpen: string, wideClose: string]> = [
  ['(', ')', '\uff08', '\uff09'],
  ['[', ']', '\uff3b', '\uff3d'],
  ['{', '}', '\uff5b', '\uff5d'],
]

/** the closer in the width of the bracket it closes, when the typed one is the other width */
export function matchingCloser(before: string, typed: string): string | null {
  const pair = BRACKETS.find((b) => typed === b[1] || typed === b[3])
  if (!pair) return null
  const [open, close, wideOpen, wideClose] = pair
  let depth = 0
  for (let i = before.length - 1; i >= 0; i--) {
    const ch = before[i]
    if (ch === close || ch === wideClose) depth++
    else if (ch === open || ch === wideOpen) {
      if (depth > 0) depth--
      else {
        const want = ch === open ? close : wideClose
        return want === typed ? null : want
      }
    }
  }
  return null
}

const parens: Rule = (ctx) => {
  const closer = matchingCloser(ctx.before, ctx.text)
  return closer
    ? ctx.state.tr.insert(ctx.from, ctx.state.schema.text(closer, caretMarks(ctx)))
    : null
}

// salutation -> closing of a Japanese letter
const CLOSINGS: Record<string, string> = {
  '\u62dd\u5553': '\u656c\u5177',
  '\u8b39\u5553': '\u8b39\u767d',
  '\u524d\u7565': '\u8349\u3005',
  '\u62dd\u5fa9': '\u656c\u5177',
  '\u6025\u5553': '\u8349\u3005',
  '\u51a0\u7701': '\u8349\u3005',
}

/** the closing goes in right-aligned below the paragraph that is exactly a salutation */
function closingBelow(tr: Transaction, $from: ResolvedPos, salutation: string): Transaction | null {
  const closing = CLOSINGS[salutation]
  const para = $from.parent
  if (!closing || para.type.name !== 'docParagraph' || para.textContent !== salutation) return null
  const container = $from.node($from.depth - 1)
  const index = $from.index($from.depth - 1)
  for (let i = index + 1; i < Math.min(container.childCount, index + 3); i++)
    if (container.child(i).textContent === closing) return null
  const styleId = para.attrs.styleId as string | null
  return tr.insert(
    tr.mapping.map($from.after()),
    para.type.create({ styleId, align: 'right' }, tr.doc.type.schema.text(closing)),
  )
}

const closingPhrase: Rule = (ctx) => {
  if (ctx.text !== ' ' && ctx.text !== '\u3000') return null
  if (ctx.$from.parentOffset !== ctx.$from.parent.content.size) return null
  return closingBelow(insertTyped(ctx.state.tr, ctx, ctx.from), ctx.$from, ctx.before)
}

// ---- list items start like the one before ----

const LEAD_MARKS = ['bold', 'italic', 'underline']

/** "Term: text" with the term formatted: the marks, and the character that ends the term */
function leadIn(item: PmNode): { marks: Mark[]; delimiter: string } | null {
  const first = item.firstChild
  if (!first?.isText) return null
  const marks = first.marks.filter((m) => LEAD_MARKS.includes(m.type.name))
  if (marks.length === 0) return null
  let text = ''
  let i = 0
  for (; i < item.childCount; i++) {
    const child = item.child(i)
    if (!child.isText || !marks.every((m) => m.isInSet(child.marks))) break
    text += child.text
  }
  if (i === item.childCount) return null
  const delimiter = text.trimEnd().slice(-1)
  return delimiter && !WORD_CHAR.test(delimiter) ? { marks, delimiter } : null
}

const listStart: Rule = (ctx) => {
  const { $from } = ctx
  const item = $from.parent
  if (item.type.name !== 'docListItem') return null
  const index = $from.index($from.depth - 1)
  const prev = index > 0 ? $from.node($from.depth - 1).child(index - 1) : null
  if (!prev || prev.type !== item.type || prev.attrs.numId !== item.attrs.numId) return null
  const lead = leadIn(prev)
  if (!lead) return null
  const marks = caretMarks(ctx)
  const formatted = lead.marks.every((m) => m.isInSet(marks))
  if (ctx.before === '') {
    if (formatted) return null
    const all = lead.marks.reduce((set, m) => m.addToSet(set), marks)
    return ctx.state.tr.insert(ctx.from, ctx.state.schema.text(ctx.text, all))
  }
  if (!formatted || ctx.text !== lead.delimiter || ctx.before.includes(lead.delimiter)) return null
  // the term is over: what follows is plain again
  return insertTyped(ctx.state.tr, ctx, ctx.from).setStoredMarks(
    marks.filter((m) => !lead.marks.some((l) => l.type === m.type)),
  )
}

// the lists must see the bare prefix before the other space rules touch it
const RULES: Array<[AutocorrectRule, Rule]> = [
  ['autoBullets', (ctx) => autoList(ctx, 'bullet')],
  ['autoNumbers', (ctx) => autoList(ctx, 'ordered')],
  ['replaceText', replaceText],
  ['spellSuggest', spellSuggest],
  ['mathOutside', mathOutside],
  ['fractions', fractions],
  ['ordinals', ordinals],
  ['dashes', dashes],
  ['smartQuotes', smartQuotes],
  ['parens', parens],
  ['asianSpaces', asianSpaces],
  ['closingPhrase', closingPhrase],
  ['boldItalic', boldItalic],
  ['twoInitialCaps', twoInitialCaps],
  ['capitalizeDays', capitalizeDays],
  ['capitalizeCells', capitalizeCells],
  ['capitalize', capitalize],
  ['listStart', listStart],
]

// ---- on Enter: border lines and tables ----

const BORDER_LINES: Record<string, { val: string; szPt: number }> = {
  '-': { val: 'single', szPt: 0.75 },
  _: { val: 'single', szPt: 1.5 },
  '=': { val: 'double', szPt: 0.75 },
  '*': { val: 'dotted', szPt: 3 },
}

/** the whole text of the plain paragraph the caret ends, or null */
function lineAtCaret(state: EditorState): { text: string; $from: ResolvedPos } | null {
  const { $from, empty } = state.selection
  if (!empty || $from.parent.type.name !== 'docParagraph') return null
  if ($from.parentOffset !== $from.parent.content.size || insideProtected($from)) return null
  const text = $from.parent.textBetween(0, $from.parent.content.size, undefined, '\ufffc')
  return text.length === $from.parent.content.size ? { text, $from } : null
}

/** "---" + Enter: the line becomes the paragraph's bottom border, the caret goes below it */
function borderLine(state: EditorState): Transaction | null {
  const line = lineAtCaret(state)
  const m = line && /^([-_=*])\1{2,}$/.exec(line.text)
  if (!line || !m) return null
  const { $from } = line
  const para = $from.parent
  const tr = state.tr.delete($from.start(), $from.end())
  tr.setNodeMarkup($from.before(), null, {
    ...para.attrs,
    borders: 'b',
    borderLines: JSON.stringify({ b: { ...BORDER_LINES[m[1]], spacePt: 1 } }),
  })
  const after = tr.mapping.map($from.after())
  tr.insert(after, para.type.create({ styleId: para.attrs.styleId as string | null }))
  return tr.setSelection(TextSelection.create(tr.doc, after + 1))
}

/** "+---+---+" + Enter: columns of a new table */
function tableColumns(state: EditorState): number | null {
  const line = lineAtCaret(state)
  return line && /^\+(?:-+\+)+$/.test(line.text) ? line.text.split('+').length - 2 : null
}

/** a short line with no closing punctuation, blank above, Enter twice: a heading; each leading tab a level down */
function headingFromLine(state: EditorState, styles?: Map<string, StyleInfo>): Transaction | null {
  const { $from, empty } = state.selection
  if (!empty || $from.depth !== 1 || $from.parent.type.name !== 'docParagraph') return null
  const index = $from.index(0)
  if ($from.parent.content.size !== 0 || index === 0) return null
  const line = state.doc.child(index - 1)
  const above = index > 1 ? state.doc.child(index - 2) : null
  if (line.type.name !== 'docParagraph' || (above && above.content.size !== 0)) return null
  const m = /^(\t*)(\S.{0,79})$/u.exec(line.textContent)
  if (!m || line.content.size !== m[0].length) return null
  if (/[.,;:!?\u3002\uff0c\uff1b\uff1a\uff01\uff1f]\s*$/.test(m[2])) return null
  const level = m[1].length + 1
  const styleId = level <= 9 ? headingStyleId(styles, level) : undefined
  const heading = state.schema.nodes.docHeading
  if (!styleId || !heading) return null
  const at = $from.before() - line.nodeSize
  const tr = state.tr.setNodeMarkup(at, heading, {
    ...line.attrs,
    styleId,
    level,
    outlineOnly: null,
  })
  return m[1] ? tr.delete(at + 1, at + 1 + m[1].length) : tr
}

/** leading spaces become the paragraph's first-line indent: two spaces, or one full-width one, a character */
function indentFromSpaces(state: EditorState, sizeHalfPoints: number): Transaction | null {
  const line = lineAtCaret(state)
  const m = line && /^([ \u3000]+)\S/.exec(line.text)
  if (!line || !m || line.$from.parent.attrs.indentFirstLine != null) return null
  const halves = [...m[1]].reduce((n, ch) => n + (ch === ' ' ? 1 : 2), 0)
  if (halves < 2) return null
  const { $from } = line
  const chars = halves * 50
  const tr = state.tr.delete($from.start(), $from.start() + m[1].length)
  return tr.setNodeMarkup($from.before(), null, {
    ...$from.parent.attrs,
    indentFirstLineChars: chars,
    indentFirstLine: Math.round((chars / 100) * sizeHalfPoints * 10),
  })
}

const TAB_INDENT_TWIPS = 720

/** the caret at the very start of a body paragraph that has text */
function paragraphStart(state: EditorState): ResolvedPos | null {
  const { $from, empty } = state.selection
  if (!empty || $from.parentOffset !== 0 || $from.parent.type.name !== 'docParagraph') return null
  if ($from.parent.content.size === 0 || insideProtected($from)) return null
  for (let d = $from.depth - 1; d > 0; d--)
    if (/^docTable(Cell|Header)$/.test($from.node(d).type.name)) return null
  return $from
}

/** Tab there indents the first line, then the whole paragraph; Backspace right after takes it back */
function tabIndent(state: EditorState): Transaction | null {
  const $from = paragraphStart(state)
  if (!$from) return null
  const attrs = $from.parent.attrs
  const change =
    (Number(attrs.indentFirstLine) || 0) === 0
      ? { indentFirstLine: TAB_INDENT_TWIPS, indentFirstLineChars: null }
      : { indentLeft: (Number(attrs.indentLeft) || 0) + TAB_INDENT_TWIPS }
  return state.tr.setNodeMarkup($from.before(), null, { ...attrs, ...change })
}

// ---- AutoComplete tip ----

const MONTHS = [
  'January',
  'February',
  'March',
  'April',
  'August',
  'September',
  'October',
  'November',
  'December',
]

/** the longest run of typed words (four characters or more) that starts one of the names */
export function autoCompleteMatch(
  before: string,
  names: string[],
): { length: number; name: string } | null {
  const from = Math.max(0, before.length - 60)
  for (let i = from; i <= before.length - 4; i++) {
    if (i > 0 && !/\s/.test(before[i - 1])) continue
    const typed = before.slice(i).toLowerCase()
    const name = names.find((n) => n.length > typed.length && n.toLowerCase().startsWith(typed))
    if (name) return { length: typed.length, name }
  }
  return null
}

interface Tip {
  from: number
  to: number
  text: string
  entry?: AutoTextEntry
}

export const autoCompleteKey = new PluginKey<Tip | null>('autoComplete')

function tipFor(state: EditorState, entries: AutoTextEntry[]): Tip | null {
  const { $from, empty } = state.selection
  if (!empty || !$from.parent.isTextblock) return null
  const before = $from.parent.textBetween(0, $from.parentOffset, undefined, '\ufffc')
  if (before.length !== $from.parentOffset) return null
  const hit = autoCompleteMatch(before, [...entries.map((e) => e.name), ...MONTHS])
  if (!hit) return null
  const entry = entries.find((e) => e.name === hit.name)
  return { from: $from.pos - hit.length, to: $from.pos, text: entry?.text ?? hit.name, entry }
}

function tipElement(tip: Tip): HTMLElement {
  const el = document.createElement('span')
  el.className = 'doc-autocomplete-tip'
  el.contentEditable = 'false'
  const text = tip.text.replace(/\s+/g, ' ').trim()
  el.textContent = `${text.length > 48 ? `${text.slice(0, 48)}\u2026` : text} ${t('appAcAutoCompleteHint')}`
  return el
}

/** Enter on a tip: the typed start becomes the whole entry */
function acceptTip(editor: Editor, tip: Tip, hooks: AutocorrectListHooks | null): void {
  const { state, view } = editor
  const content = tip.entry ? inlineContent(state, tip.entry.html) : null
  if (!tip.entry || content?.size) {
    const marks = state.doc.resolve(tip.to).marks()
    view.dispatch(
      closeHistory(state.tr).replaceWith(
        tip.from,
        tip.to,
        content?.size ? content : state.schema.text(tip.text, marks),
      ),
    )
    return
  }
  view.dispatch(closeHistory(state.tr).delete(tip.from, tip.to))
  hooks?.insertAutoText?.(tip.entry)
}

// ---- smart button: what was corrected, to change back or stop ----

/** a correction still standing in the text */
export interface Fix {
  /** tells two corrections of the same word apart */
  id: number
  from: number
  to: number
  /** what was typed */
  original: string
  /** what it became */
  corrected: string
  rule: AutocorrectRule
}

export const fixesKey = new PluginKey<Fix[]>('autocorrectFixes')
const FIXES_KEPT = 30
let nextFixId = 1

/** the typed text a rule changed inside its paragraph, widened to whole words */
function fixOf(ctx: Ctx, tr: Transaction, rule: AutocorrectRule): Fix | null {
  if (!tr.selection.empty) return null
  const $to = tr.selection.$from
  if ($to.parent.type !== ctx.$from.parent.type || $to.start() !== ctx.$from.start()) return null
  const expected = ctx.before + ctx.text
  const actual = $to.parent.textBetween(0, $to.parentOffset, undefined, '\ufffc')
  if (actual.length !== $to.parentOffset || actual === expected) return null
  const shorter = Math.min(expected.length, actual.length)
  let head = 0
  while (head < shorter && expected[head] === actual[head]) head++
  let tail = 0
  while (
    tail < shorter - head &&
    expected[expected.length - 1 - tail] === actual[actual.length - 1 - tail]
  )
    tail++
  const wordy = (text: string) => text !== '' && [...text].every((ch) => WORD_CHAR.test(ch))
  if (wordy(expected.slice(head, expected.length - tail))) {
    while (head > 0 && WORD_CHAR.test(expected[head - 1])) head--
    while (tail > 0 && WORD_CHAR.test(expected[expected.length - tail])) tail--
  }
  const original = expected.slice(head, expected.length - tail)
  const corrected = actual.slice(head, actual.length - tail)
  if (!original || !corrected) return null
  const from = $to.start() + head
  return { id: nextFixId++, from, to: from + corrected.length, original, corrected, rule }
}

function mapFixes(fixes: Fix[], tr: Transaction): Fix[] {
  if (!tr.docChanged) return fixes
  const out: Fix[] = []
  for (const fix of fixes) {
    const from = tr.mapping.map(fix.from, 1)
    const to = tr.mapping.map(fix.to, -1)
    // typed into or deleted: no longer the correction
    if (to - from !== fix.corrected.length || to > tr.doc.content.size) continue
    if (tr.doc.textBetween(from, to, '\n', '\ufffc') !== fix.corrected) continue
    out.push({ ...fix, from, to })
  }
  return out
}

/** Change Back: the typed text returns in place of the correction */
export function changeBackFix(state: EditorState, fix: Fix): Transaction {
  const marks = state.doc.resolve(fix.to).marks()
  return state.tr
    .replaceWith(fix.from, fix.to, state.schema.text(fix.original, marks))
    .setMeta(fixesKey, { remove: fix.from })
}

/** Stop Automatically Correcting: the entry goes, the word becomes an exception, or the rule is switched off */
export function stopCorrecting(fix: Fix): void {
  const lists = autocorrectLists()
  // the replace table matches whatever the case; math names are exact (\delta is not \Delta)
  const same = ([from]: ReplaceEntry) => from.toLowerCase() === fix.original.toLowerCase()
  const exact = ([from]: ReplaceEntry) => from === fix.original
  if (fix.rule === 'replaceText' && lists.replace.some(same))
    setAutocorrectList(
      'replace',
      lists.replace.filter((entry) => !same(entry)),
    )
  else if (fix.rule === 'mathOutside' && lists.math.some(exact))
    setAutocorrectList(
      'math',
      lists.math.filter((entry) => !exact(entry)),
    )
  else if (fix.rule === 'twoInitialCaps')
    setAutocorrectList('initialCaps', [...lists.initialCaps, fix.original].sort())
  else setAutocorrectPref(fix.rule, false)
}

/** the button that appears over a correction under the pointer, and its menu */
class SmartButtonView {
  private button = document.createElement('button')
  private menu: HTMLElement | null = null
  private fix: Fix | null = null
  private hideTimer = 0

  constructor(
    private view: EditorView,
    private storage: AutocorrectStorage,
  ) {
    this.button.type = 'button'
    this.button.className = 'ac-smart-button'
    this.button.setAttribute('aria-label', t('appAcTitle'))
    this.button.textContent = '\u26a1'
    this.button.addEventListener('mousedown', (e) => e.preventDefault())
    this.button.addEventListener('click', () => this.openMenu())
    this.button.addEventListener('mouseenter', () => window.clearTimeout(this.hideTimer))
    this.button.addEventListener('mouseleave', this.onOut)
    view.dom.addEventListener('mouseover', this.onOver)
    view.dom.addEventListener('mouseout', this.onOut)
    document.addEventListener('mousedown', this.onDown, true)
    window.addEventListener('scroll', this.hide, true)
  }

  private onOver = (event: Event) => {
    const el = (event.target as Element | null)?.closest?.('.doc-ac-fix')
    if (!el || this.menu) return
    // by id: two corrections side by side share a boundary position
    const id = Number(el.getAttribute('data-ac-fix'))
    const fix = (fixesKey.getState(this.view.state) ?? []).find((f) => f.id === id)
    if (!fix) return
    window.clearTimeout(this.hideTimer)
    this.fix = fix
    const rect = el.getBoundingClientRect()
    this.button.style.left = `${rect.left}px`
    this.button.style.top = `${rect.bottom + 2}px`
    document.body.appendChild(this.button)
  }

  private onOut = () => {
    if (this.menu) return
    window.clearTimeout(this.hideTimer)
    this.hideTimer = window.setTimeout(this.hide, 500)
  }

  private onDown = (event: Event) => {
    const target = event.target as Node
    if (!this.button.contains(target) && !this.menu?.contains(target)) this.hide()
  }

  private hide = () => {
    window.clearTimeout(this.hideTimer)
    this.button.remove()
    this.menu?.remove()
    this.menu = null
    this.fix = null
  }

  private openMenu() {
    const fix = this.fix
    if (!fix || this.menu) return
    const menu = document.createElement('div')
    menu.className = 'ac-smart-menu'
    menu.setAttribute('role', 'menu')
    const item = (label: string, run: () => void) => {
      const el = document.createElement('button')
      el.type = 'button'
      el.setAttribute('role', 'menuitem')
      el.textContent = label
      el.addEventListener('mousedown', (e) => e.preventDefault())
      el.addEventListener('click', () => {
        this.hide()
        run()
      })
      menu.appendChild(el)
    }
    const live = () => (fixesKey.getState(this.view.state) ?? []).find((f) => f.id === fix.id)
    const changeBack = () => {
      const now = live()
      if (now && this.view.editable) this.view.dispatch(changeBackFix(this.view.state, now))
    }
    item(t('appAcChangeBack', { text: fix.original }), changeBack)
    item(t('appAcStopCorrecting', { text: fix.original }), () => {
      stopCorrecting(fix)
      changeBack()
    })
    item(t('appAcControlOptions'), () => this.storage.hooks?.openOptions?.())
    menu.style.left = this.button.style.left
    menu.style.top = `${parseFloat(this.button.style.top) + 22}px`
    document.body.appendChild(menu)
    this.menu = menu
  }

  update() {
    // the correction under the button was typed over or changed back
    if (this.fix && !(fixesKey.getState(this.view.state) ?? []).some((f) => f.id === this.fix?.id))
      this.hide()
  }

  destroy() {
    this.hide()
    this.view.dom.removeEventListener('mouseover', this.onOver)
    this.view.dom.removeEventListener('mouseout', this.onOut)
    document.removeEventListener('mousedown', this.onDown, true)
    window.removeEventListener('scroll', this.hide, true)
  }
}

function undoAutocorrect(state: EditorState, dispatch?: (tr: Transaction) => void): boolean {
  const record = autocorrectKey.getState(state)
  if (!record) return false
  if (dispatch) {
    const tr = state.tr
    const { steps, docs } = record.transform
    for (let j = steps.length - 1; j >= 0; j--) tr.step(steps[j].invert(docs[j]))
    if (record.text) {
      const marks = tr.doc.resolve(record.from).marks()
      tr.replaceWith(record.from, record.to, state.schema.text(record.text, marks))
    } else {
      tr.delete(record.from, record.to)
    }
    // its own history event, so Cmd+Z after the revert restores exactly the correction
    dispatch(closeHistory(tr.setMeta(autocorrectKey, null)))
  }
  return true
}

export const AutoCorrectExtension = Extension.create<Record<string, never>, AutocorrectStorage>({
  name: 'autocorrect',
  // ahead of the core keymap: Backspace must revert a correction before joinBackward runs
  priority: 1001,

  addStorage() {
    return { hooks: null }
  },

  addProseMirrorPlugins() {
    const storage = this.storage
    const editor = this.editor
    return [
      new Plugin<Undoable | null>({
        key: autocorrectKey,
        state: {
          init: () => null,
          apply(tr, prev) {
            const meta = tr.getMeta(autocorrectKey) as Undoable | null | undefined
            if (meta !== undefined) return meta
            return tr.docChanged || tr.selectionSet ? null : prev
          },
        },
        props: {
          handleTextInput(view, from, to, text) {
            if (view.composing || !view.editable || from !== to) return false
            const { state } = view
            const $from = state.doc.resolve(from)
            if (!$from.parent.isTextblock) return false
            const before = $from.parent.textBetween(0, $from.parentOffset, undefined, '\ufffc')
            // inline nodes with content would skew offsets; nothing to correct there
            if (before.length !== $from.parentOffset) return false
            const ctx: Ctx = { state, $from, from, text, before, hooks: storage.hooks }
            for (const [pref, rule] of RULES) {
              if (!autocorrectEnabled(pref)) continue
              const tr = rule(ctx)
              if (!tr) continue
              tr.setMeta(autocorrectKey, { transform: tr, from, to, text })
              const fix = fixOf(ctx, tr, pref)
              if (fix) tr.setMeta(fixesKey, { add: fix })
              view.dispatch(tr)
              return true
            }
            return false
          },
          handleKeyDown(view, event) {
            if (event.shiftKey || event.altKey || event.ctrlKey || event.metaKey) return false
            if (view.composing || !view.editable) return false
            const { state } = view
            const record = (tr: Transaction) => {
              const at = state.selection.from
              view.dispatch(
                tr.setMeta(autocorrectKey, { transform: tr, from: at, to: at, text: '' }),
              )
              return true
            }
            const tip = autoCompleteKey.getState(state)
            if (tip && event.key === 'Escape') {
              view.dispatch(state.tr.setMeta(autoCompleteKey, null))
              return true
            }
            if (event.key === 'Tab' && autocorrectEnabled('tabIndents')) {
              const tr = tabIndent(state)
              return tr ? record(tr) : false
            }
            if (event.key !== 'Enter') return false
            if (tip && tip.to === state.selection.from) {
              acceptTip(editor, tip, storage.hooks)
              return true
            }
            const numbering = editor.storage.listNumbering
            if (autocorrectEnabled('headingStyles')) {
              const tr = headingFromLine(state, numbering?.styles)
              if (tr) return record(tr)
            }
            if (autocorrectEnabled('borderLines')) {
              const tr = borderLine(state)
              if (tr) return record(tr)
            }
            const insertTable = storage.hooks?.insertTable
            if (insertTable && autocorrectEnabled('tables')) {
              const cols = tableColumns(state)
              if (cols) {
                insertTable(cols)
                return true
              }
            }
            const defineStyle = storage.hooks?.defineStyle
            if (
              autocorrectEnabled('dateStyle') &&
              dateStyleOnEnter(editor, numbering?.styles, defineStyle)
            )
              return true
            if (
              autocorrectEnabled('defineStyles') &&
              definedStyleOnEnter(editor, numbering?.styles, numbering?.docDefaults, defineStyle)
            )
              return true
            // these two change the paragraph and let Enter go on to split it
            if (autocorrectEnabled('closingPhrase')) {
              const { $from } = state.selection
              const tr = state.selection.empty
                ? closingBelow(state.tr, $from, $from.parent.textContent)
                : null
              if (tr) view.dispatch(tr)
            }
            if (autocorrectEnabled('indentSpaces')) {
              const size =
                effectiveSizeHalfPoints(editor, numbering?.styles, numbering?.docDefaults) ?? 20
              const tr = indentFromSpaces(view.state, size)
              if (tr) view.dispatch(tr)
            }
            return false
          },
        },
      }),
      new Plugin<Fix[]>({
        key: fixesKey,
        state: {
          init: () => [],
          apply(tr, fixes) {
            const meta = tr.getMeta(fixesKey) as { add?: Fix; remove?: number } | undefined
            let next = mapFixes(fixes, tr)
            if (meta?.remove !== undefined) {
              const gone = tr.mapping.map(meta.remove, 1)
              next = next.filter((f) => f.from !== gone)
            }
            if (meta?.add) next = [...next, meta.add].slice(-FIXES_KEPT)
            return next
          },
        },
        props: {
          decorations(state) {
            const fixes = fixesKey.getState(state) ?? []
            if (fixes.length === 0 || !autocorrectEnabled('smartButton')) return null
            return DecorationSet.create(
              state.doc,
              fixes.map((f) =>
                Decoration.inline(f.from, f.to, {
                  class: 'doc-ac-fix',
                  'data-ac-fix': String(f.id),
                }),
              ),
            )
          },
        },
        view: (view) => new SmartButtonView(view, storage),
      }),
      new Plugin<Tip | null>({
        key: autoCompleteKey,
        state: {
          init: () => null,
          apply(tr, prev, _old, state) {
            const meta = tr.getMeta(autoCompleteKey) as Tip | null | undefined
            if (meta !== undefined) return meta
            if (!tr.docChanged) return tr.selectionSet ? null : prev
            if (tr.getMeta('paste') || !autocorrectEnabled('autoComplete')) return null
            return tipFor(state, storage.hooks?.autoText?.() ?? [])
          },
        },
        props: {
          decorations(state) {
            const tip = autoCompleteKey.getState(state)
            if (!tip) return null
            return DecorationSet.create(state.doc, [
              Decoration.widget(tip.to, () => tipElement(tip), {
                side: 1,
                key: `autocomplete:${tip.text}`,
                ignoreSelection: true,
              }),
            ])
          },
        },
      }),
      keymap({ Backspace: undoAutocorrect }),
    ]
  },
})
