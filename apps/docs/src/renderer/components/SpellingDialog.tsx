/**
 * Word's Spelling and Grammar dialog (Tools ▸ Spelling and Grammar): walks the misspelled words
 * from the caret on, around the end of the document and back to where it started.
 */
import { useEffect, useRef, useState } from 'react'
import type { Editor } from '@tiptap/core'
import { useI18n } from '../i18n/locale'
import { nextMisspelling, spellWords, type SpellWord } from '../editor/spell-walk'
import { useModalKeys } from './modal-keys'

const CONTEXT_CHARS = 60

export function SpellingDialog({ editor, onClose }: { editor: Editor; onClose: () => void }) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  /** where the check started; the walk wraps around the end and stops here */
  const stop = useRef(editor.state.selection.from)
  const skipped = useRef(new Set<string>())
  const [at, setAt] = useState({ pos: editor.state.selection.from, wrapped: false })
  const [found, setFound] = useState<SpellWord | 'done' | null>(null)
  const [suggestions, setSuggestions] = useState<string[]>([])
  const [picked, setPicked] = useState<string | null>(null)

  useEffect(() => {
    const { doc } = editor.state
    const misspelled = (word: string) =>
      !skipped.current.has(word) && (window.desktop.spellIsMisspelled?.(word) ?? false)
    let wrapped = at.wrapped
    let next = nextMisspelling(doc, at.pos, wrapped ? stop.current : doc.content.size, misspelled)
    if (!next && !wrapped) {
      wrapped = true
      next = nextMisspelling(doc, 0, stop.current, misspelled)
    }
    if (!next) return setFound('done')
    if (wrapped !== at.wrapped) return setAt({ pos: next.from, wrapped })
    const offered = window.desktop.spellSuggestions?.(next.word) ?? []
    setFound(next)
    setSuggestions(offered)
    setPicked(offered[0] ?? null)
    editor.commands.setTextSelection({ from: next.from, to: next.to })
    const dom = editor.view.domAtPos(next.from).node
    ;(dom instanceof Element ? dom : dom.parentElement)?.scrollIntoView({ block: 'center' })
  }, [editor, at])

  const word = found && found !== 'done' ? found : null
  const resume = (pos: number) => setAt((cur) => ({ pos, wrapped: cur.wrapped }))
  const skipAll = (remember?: (word: string) => Promise<unknown> | undefined) => {
    if (!word) return
    skipped.current.add(word.word)
    void remember?.(word.word)
    resume(word.to)
  }
  const change = (all: boolean) => {
    if (!word || picked === null || !editor.isEditable) return
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
      text.slice(offset, offset + word.word.length),
      text.slice(offset + word.word.length, offset + word.word.length + CONTEXT_CHARS),
    ]
  }

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
          <p className="spelling-done">{t('appSpellingComplete')}</p>
        ) : (
          <div className="spelling-body">
            <div className="spelling-fields">
              <span>{t('appSpellingNotInDictionary')}</span>
              <div className="spelling-context" aria-busy={!word}>
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
              <button type="button" disabled={!word} onClick={() => word && resume(word.to)}>
                {t('appSpellingIgnoreOnce')}
              </button>
              <button
                type="button"
                disabled={!word}
                onClick={() => skipAll(window.desktop.spellIgnoreWord)}
              >
                {t('appSpellIgnoreAll')}
              </button>
              <button
                type="button"
                disabled={!word}
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
                disabled={!word || picked === null || !editor.isEditable}
                onClick={() => change(true)}
              >
                {t('appSpellingChangeAll')}
              </button>
            </div>
          </div>
        )}
        <div className="modal-actions">
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
