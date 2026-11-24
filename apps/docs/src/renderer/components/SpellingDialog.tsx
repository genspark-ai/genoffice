/**
 * Word's Spelling and Grammar dialog (Tools ▸ Spelling and Grammar): walks the misspelled words
 * and, with Check grammar ticked, the grammar findings of an AI check, from the caret on,
 * around the end of the document and back to where it started.
 */
import { useEffect, useRef, useState } from 'react'
import type { Editor } from '@tiptap/core'
import { useI18n } from '../i18n/locale'
import { checkGrammar, type GrammarFinding } from '../ai/grammar'
import {
  applyGrammarSuggestion,
  grammarState,
  ignoreGrammarFinding,
  setGrammarFindings,
} from '../editor/grammar-marks'
import { ignoreSpellingOnce, spellingDismissed } from '../editor/spell-menu'
import { nextMisspelling, spellWords, type SpellWord } from '../editor/spell-walk'
import { checkGrammarEnabled, setCheckGrammarPref } from '../grammar-pref'
import { useModalKeys } from './modal-keys'

const CONTEXT_CHARS = 60

type Issue = SpellWord & { grammar?: GrammarFinding }

interface GrammarNote {
  error?: string
  capped?: { checked: number; total: number }
}

const isMisspelled = (word: string) => window.desktop.spellIsMisspelled?.(word) ?? false

export function SpellingDialog({
  editor,
  onClose,
  grammarAt,
}: {
  editor: Editor
  onClose: () => void
  /** the finding to open at, from Grammar… on its right-click menu: no new check is made */
  grammarAt?: string | null
}) {
  const { t, lang } = useI18n()
  const keys = useModalKeys(onClose)
  const opening = useRef(
    (grammarAt && grammarState(editor.state)?.findings.find((f) => f.id === grammarAt)) || null,
  )
  const start = opening.current?.from ?? editor.state.selection.from
  /** where the check started; the walk wraps around the end and stops here */
  const stop = useRef(start)
  const skipped = useRef(new Set<string>())
  const [at, setAt] = useState({ pos: start, wrapped: false })
  const [found, setFound] = useState<Issue | 'done' | null>(null)
  const [suggestions, setSuggestions] = useState<string[]>([])
  const [picked, setPicked] = useState<string | null>(null)
  // Grammar… on a finding walks the findings whatever the box was left at
  const [withGrammar, setWithGrammar] = useState(() => !!opening.current || checkGrammarEnabled())
  const [grammar, setGrammar] = useState<'off' | 'checking' | 'ready'>(() =>
    opening.current ? 'ready' : withGrammar ? 'checking' : 'off',
  )
  const [note, setNote] = useState<GrammarNote>({})

  useEffect(() => {
    if (grammar !== 'checking') return
    let cancelled = false
    void checkGrammar(editor.state.doc, {
      lang,
      isMisspelled,
      ignored: grammarState(editor.state)?.ignored,
      cancelled: () => cancelled,
    }).then((result) => {
      if (cancelled || editor.isDestroyed) return
      setGrammarFindings(editor, result.findings)
      setNote({
        error: result.error,
        capped:
          result.checked < result.total
            ? { checked: result.checked, total: result.total }
            : undefined,
      })
      setGrammar('ready')
    })
    return () => {
      cancelled = true
    }
  }, [editor, grammar, lang])

  useEffect(() => {
    if (grammar === 'checking') return setFound(null)
    const { doc } = editor.state
    const misspelled = (word: string) => !skipped.current.has(word) && isMisspelled(word)
    const findings = grammar === 'ready' ? (grammarState(editor.state)?.findings ?? []) : []
    const nextIssue = (from: number, to: number): Issue | null => {
      const word = nextMisspelling(doc, from, to, misspelled, (w) =>
        spellingDismissed(editor.state, w),
      )
      const finding = findings.find((f) => f.to > from && f.to <= to)
      if (finding && (!word || finding.from < word.from))
        return { from: finding.from, to: finding.to, word: finding.quote, grammar: finding }
      return word
    }
    let wrapped = at.wrapped
    let next = nextIssue(at.pos, wrapped ? stop.current : doc.content.size)
    if (!next && !wrapped) {
      wrapped = true
      next = nextIssue(0, stop.current)
    }
    if (!next) return setFound('done')
    if (wrapped !== at.wrapped) return setAt({ pos: next.from, wrapped })
    const offered = next.grammar
      ? next.grammar.suggestions
      : (window.desktop.spellSuggestions?.(next.word) ?? [])
    setFound(next)
    setSuggestions(offered)
    setPicked(offered[0] ?? null)
    editor.commands.setTextSelection({ from: next.from, to: next.to })
    const dom = editor.view.domAtPos(next.from).node
    ;(dom instanceof Element ? dom : dom.parentElement)?.scrollIntoView({ block: 'center' })
  }, [editor, at, grammar])

  const word = found && found !== 'done' ? found : null
  const finding = word?.grammar ?? null
  const resume = (pos: number) => setAt((cur) => ({ pos, wrapped: cur.wrapped }))
  const toggleGrammar = (on: boolean) => {
    setCheckGrammarPref(on)
    setWithGrammar(on)
    setNote({})
    setGrammar(on ? 'checking' : 'off')
    if (found === 'done') setAt({ pos: stop.current, wrapped: false })
  }
  const ignoreOnce = () => {
    if (!word) return
    if (finding) ignoreGrammarFinding(editor, finding.id)
    else ignoreSpellingOnce(editor, word)
    resume(word.to)
  }
  const skipAll = (remember?: (word: string) => Promise<unknown> | undefined) => {
    if (!word || finding) return
    skipped.current.add(word.word)
    void remember?.(word.word)
    resume(word.to)
  }
  const change = (all: boolean) => {
    if (!word || picked === null || !editor.isEditable) return
    if (finding) {
      const before = editor.state.doc.content.size
      if (!applyGrammarSuggestion(editor, finding.id, picked)) return resume(word.to)
      const grown = editor.state.doc.content.size - before
      if (stop.current >= word.to) stop.current += grown
      return resume(word.to + grown)
    }
    const { doc, schema } = editor.state
    const targets = all ? spellWords(doc).filter((w) => w.word === word.word) : [word]
    const tr = editor.state.tr
    for (const target of targets.reverse())
      tr.replaceWith(
        target.from,
        target.to,
        schema.text(picked, doc.nodeAt(target.from)?.marks ?? []),
      )
    stop.current = tr.mapping.map(stop.current)
    editor.view.dispatch(tr)
    resume(tr.mapping.map(word.to))
  }

  let context: [string, string, string] | null = null
  if (word) {
    const $from = editor.state.doc.resolve(word.from)
    const offset = word.from - $from.start()
    const text = $from.parent.textBetween(0, $from.parent.content.size, undefined, ' ')
    context = [
      text.slice(Math.max(0, offset - CONTEXT_CHARS), offset),
      text.slice(offset, offset + word.to - word.from),
      text.slice(offset + word.to - word.from, offset + word.to - word.from + CONTEXT_CHARS),
    ]
  }
  const checkedGrammar = grammar === 'ready' && note.error === undefined

  return (
    <div
      className="modal-backdrop manual-hyphenation-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
    >
      <form
        className="modal gs-form tco-dialog spelling-dialog"
        role="dialog"
        aria-label={t('appSpellingTitle')}
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          if (found === 'done') onClose()
          else change(false)
        }}
      >
        <h2>{t('appSpellingTitle')}</h2>
        {found === 'done' ? (
          <p className="spelling-done">
            {t(checkedGrammar ? 'appSpellingGrammarComplete' : 'appSpellingComplete')}
          </p>
        ) : (
          <div className="spelling-body">
            <div className="spelling-fields">
              <span className="spelling-problem">
                {grammar === 'checking'
                  ? t('appGrammarChecking')
                  : finding
                    ? finding.message
                    : t('appSpellingNotInDictionary')}
              </span>
              <div className={`spelling-context${finding ? ' grammar' : ''}`} aria-busy={!word}>
                {context && (
                  <>
                    {context[0]}
                    <b>{context[1]}</b>
                    {context[2]}
                  </>
                )}
              </div>
              <span>{t('appSpellingSuggestions')}</span>
              <div className="spelling-suggestions" role="listbox">
                {suggestions.length === 0 && word ? (
                  <div className="spelling-none">{t('appSpellNoSuggestions')}</div>
                ) : (
                  suggestions.map((s) => (
                    <div
                      key={s}
                      role="option"
                      aria-selected={s === picked}
                      className={s === picked ? 'active' : undefined}
                      onClick={() => setPicked(s)}
                    >
                      {s}
                    </div>
                  ))
                )}
              </div>
            </div>
            <div className="modal-actions spelling-buttons">
              <button type="button" disabled={!word} onClick={ignoreOnce}>
                {t('appSpellingIgnoreOnce')}
              </button>
              <button
                type="button"
                disabled={!word || !!finding}
                onClick={() => skipAll(window.desktop.spellIgnoreWord)}
              >
                {t('appSpellIgnoreAll')}
              </button>
              <button
                type="button"
                disabled={!word || !!finding}
                onClick={() => skipAll(window.desktop.spellAddWord)}
              >
                {t('appSpellAddToDictionary')}
              </button>
              <button
                type="submit"
                className="primary"
                disabled={!word || picked === null || !editor.isEditable}
              >
                {t('appSpellingChange')}
              </button>
              <button
                type="button"
                disabled={!word || !!finding || picked === null || !editor.isEditable}
                onClick={() => change(true)}
              >
                {t('appSpellingChangeAll')}
              </button>
            </div>
          </div>
        )}
        {note.error !== undefined && (
          <p className="spelling-grammar-note" role="status">
            {t('appGrammarUnavailable')}
            {note.error ? ` (${note.error})` : ''}
          </p>
        )}
        {note.capped && (
          <p className="spelling-grammar-note" role="status">
            {t('appGrammarCapped', note.capped)}
          </p>
        )}
        <div className="modal-actions">
          <label className="tco-check spelling-check-grammar">
            <input
              type="checkbox"
              name="checkGrammar"
              checked={withGrammar}
              onChange={(e) => toggleGrammar(e.target.checked)}
            />
            {t('appGrammarCheck')}
          </label>
          {found === 'done' ? (
            <button type="submit" className="primary" autoFocus>
              {t('appOk')}
            </button>
          ) : (
            <button type="button" onClick={onClose}>
              {t('appClose')}
            </button>
          )}
        </div>
      </form>
    </div>
  )
}
