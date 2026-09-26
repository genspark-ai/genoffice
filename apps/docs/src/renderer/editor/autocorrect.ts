/**
 * Word's AutoCorrect / AutoFormat-as-you-type: smart quotes, "--" to dashes,
 * automatic bulleted and numbered lists, (c) (r) (tm) ... fractions and
 * arrows, superscript ordinals and (opt-in) sentence capitalization.
 *
 * Only real typed text reaches handleTextInput, so paste, drop, AI edits and
 * IME composition never trigger a rule. Each correction records an undo entry
 * that one Backspace right afterwards reverts, as in Word.
 */
import { Extension } from '@tiptap/core'
import type { Mark, ResolvedPos } from '@tiptap/pm/model'
import { Plugin, PluginKey, type EditorState, type Transaction } from '@tiptap/pm/state'
import { keymap } from '@tiptap/pm/keymap'
import { closeHistory } from '@tiptap/pm/history'
import { autocorrectEnabled, type AutocorrectRule } from '../autocorrect-pref'

export interface AutocorrectListHooks {
  /** numId for a typed "- " / "* " bullet (document definition or a new one) */
  bulletNumId: () => string | null
  /** numId for a typed "1." / "a)" / "(1)" list, starting at the typed value */
  numberedNumId: (numFmt: string, pattern: string, start: number) => string | null
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

const autoLists: Rule = (ctx) => {
  if (ctx.text !== ' ') return null
  const { $from, state } = ctx
  const listType = state.schema.nodes.docListItem
  if (!listType || $from.parent.type.name !== 'docParagraph') return null
  if ($from.parentOffset !== $from.parent.content.size || insideProtected($from)) return null
  const spec = parseListPrefix(ctx.before)
  if (!spec) return null
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

// ---- symbols ----

const CLOSING_SYMBOLS: Array<[RegExp, string]> = [
  [/\(c$/i, '\u00a9'],
  [/\(r$/i, '\u00ae'],
  [/\(tm$/i, '\u2122'],
]
const FRACTIONS: Record<string, string> = { '1/2': '\u00bd', '1/4': '\u00bc', '3/4': '\u00be' }

const symbols: Rule = (ctx) => {
  const { before, text } = ctx
  if (text === ')') {
    for (const [re, glyph] of CLOSING_SYMBOLS) {
      const m = re.exec(before)
      if (m) return replaceTail(ctx, m[0].length, glyph)
    }
    return null
  }
  if (text === '.' && /(^|[^.])\.\.$/.test(before)) return replaceTail(ctx, 2, '\u2026')
  if (text === '>' && /(^|[^-])--$/.test(before)) return replaceTail(ctx, 2, '\u2192')
  if (text === '-' && /(^|[^<])<-$/.test(before)) return replaceTail(ctx, 2, '\u2190')
  if (text === ' ') {
    const m = /(^|[^\d/])([13]\/[24])$/.exec(before)
    const glyph = m ? FRACTIONS[m[2]] : undefined
    if (glyph) {
      const tr = replaceTail(ctx, 3, glyph)
      return insertTyped(tr, ctx, tr.mapping.map(ctx.from))
    }
  }
  return null
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

// ---- capitalize first letter of sentences ----

const ABBREVIATIONS =
  /^(e\.g|i\.e|etc|vs|cf|approx|no|fig|inc|ltd|co|mr|mrs|ms|dr|prof|sr|jr|st|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec|a\.m|p\.m|u\.s|u\.k)$/i

const capitalize: Rule = (ctx) => {
  if (!/^[ .,;:!?]$/.test(ctx.text)) return null
  const word = /(\S+)$/.exec(ctx.before)?.[1]
  if (!word || !/^[a-z]/.test(word) || /[@/.:\\]/.test(word)) return null
  const prefix = ctx.before.slice(0, -word.length)
  if (prefix !== '') {
    const end = /(\S+)([.!?])\s+$/.exec(prefix)
    if (!end) return null
    const token = end[1]
    // "e.g. apples", "3. apples", "J. smith": a period that isn't a sentence end
    if (end[2] === '.' && (ABBREVIATIONS.test(token) || /^\d+$/.test(token) || token.length === 1))
      return null
  }
  const at = ctx.from - word.length
  const tr = ctx.state.tr.insertText(word[0].toUpperCase(), at, at + 1)
  return insertTyped(tr, ctx, ctx.from)
}

// autoLists must see the bare prefix before the other space rules touch it
const RULES: Array<[AutocorrectRule, Rule]> = [
  ['autoLists', autoLists],
  ['symbols', symbols],
  ['ordinals', ordinals],
  ['dashes', dashes],
  ['smartQuotes', smartQuotes],
  ['capitalize', capitalize],
]

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
              view.dispatch(tr)
              return true
            }
            return false
          },
        },
      }),
      keymap({ Backspace: undoAutocorrect }),
    ]
  },
})
