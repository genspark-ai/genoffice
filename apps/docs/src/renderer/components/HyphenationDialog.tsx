import { useEffect, useState } from 'react'
import type { Editor } from '@tiptap/core'
import { useI18n } from '../i18n/locale'
import {
  hyphenationReady,
  nextManualCandidate,
  skipManualWord,
  type HyphenationSettings,
  type ManualCandidate,
} from '../editor/hyphenation'
import { LengthInput } from './LengthInput'
import { useModalKeys } from './modal-keys'

type HyphenationChoice = HyphenationSettings & { auto: boolean }

const LIMITS = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]

/** Word's Tools > Hyphenation: whether and how words break at line ends through the whole document */
export function HyphenationDialog({
  auto,
  settings,
  hasPatterns,
  zoneApplies,
  canEdit,
  onApply,
  onManual,
  onClose,
}: {
  auto: boolean
  settings: HyphenationSettings
  /** false when no hyphenation patterns exist for the document's language: Chromium breaks the words by its own rule */
  hasPatterns: boolean
  /** false in a document laid out the Word 2013+ way, which has no zone */
  zoneApplies: boolean
  canEdit: boolean
  onApply: (next: HyphenationChoice) => void
  onManual: (next: HyphenationChoice) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [on, setOn] = useState(auto)
  const [caps, setCaps] = useState(settings.caps)
  const [limit, setLimit] = useState(settings.limit)
  const [zoneTwips, setZoneTwips] = useState(settings.zoneTwips)
  const choice = { auto: on, caps, limit, zoneTwips }
  const noPatterns = hasPatterns ? undefined : t('ribbonHyphNoPatterns')
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form tco-dialog hlc-dialog hyphenation-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          onClose()
          onApply(choice)
        }}
      >
        <h2>{t('ribbonHyphenation')}</h2>
        <label className="tco-check">
          <input type="checkbox" checked={on} onChange={(e) => setOn(e.target.checked)} />
          {t('ribbonHyphAuto')}
        </label>
        <label className="tco-check" title={noPatterns}>
          <input
            type="checkbox"
            checked={caps}
            disabled={!hasPatterns}
            onChange={(e) => setCaps(e.target.checked)}
          />
          {t('ribbonHyphCaps')}
        </label>
        <div className="modal-row margin-row">
          <label title={noPatterns ?? (zoneApplies ? undefined : t('ribbonHyphZoneUnused'))}>
            {t('ribbonHyphZone')}
            <LengthInput
              value={zoneTwips}
              onCommit={(twips) => twips !== null && setZoneTwips(twips)}
              live
              max={31680}
              disabled={!hasPatterns || !zoneApplies}
            />
          </label>
        </div>
        <div className="modal-row margin-row">
          <label title={noPatterns}>
            {t('ribbonHyphLimit')}
            <select
              disabled={!hasPatterns}
              value={limit}
              onChange={(e) => setLimit(Number(e.target.value))}
            >
              <option value={0}>{t('ribbonHyphNoLimit')}</option>
              {(LIMITS.includes(limit) || limit === 0 ? LIMITS : [...LIMITS, limit]).map((n) => (
                <option key={n} value={n}>
                  {n}
                </option>
              ))}
            </select>
          </label>
        </div>
        <div className="modal-actions">
          <button
            type="button"
            disabled={!hasPatterns || !canEdit}
            title={noPatterns}
            onClick={() => {
              onClose()
              onManual(choice)
            }}
          >
            {t('ribbonHyphManual')}
          </button>
          <div className="tco-actions-right">
            <button type="button" onClick={onClose}>
              {t('ribbonCancel')}
            </button>
            <button type="submit" className="primary">
              {t('ribbonOk')}
            </button>
          </div>
        </div>
      </form>
    </div>
  )
}

const frame = () => new Promise<void>((resolve) => requestAnimationFrame(() => resolve()))

/** the page has the break points of the current answers laid out */
async function probeSettled(editor: Editor): Promise<void> {
  for (let i = 0; i < 180; i++) {
    await frame()
    const dict = editor.storage.hyphenation.dict
    if (dict && hyphenationReady(dict)) break
  }
  await frame()
  await frame()
}

/** Word's Manual Hyphenation: stops at each word that could break at a line end and asks where */
export function ManualHyphenationDialog({
  editor,
  onClose,
}: {
  editor: Editor
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [after, setAfter] = useState(0)
  const [found, setFound] = useState<ManualCandidate | 'done' | null>(null)
  const [point, setPoint] = useState(0)

  useEffect(() => {
    let stale = false
    setFound(null)
    void probeSettled(editor).then(() => {
      if (stale || editor.isDestroyed) return
      const next = nextManualCandidate(editor, after)
      setFound(next ?? 'done')
      if (next) {
        setPoint(next.chosen)
        next.el.scrollIntoView({ block: 'center' })
      }
    })
    return () => {
      stale = true
    }
  }, [editor, after])

  const word = found && found !== 'done' ? found : null
  const answer = (yes: boolean) => {
    if (!word) return
    if (yes) {
      editor.view.dispatch(editor.state.tr.insertText('\u00ad', word.from + point))
      setAfter(word.to + 1)
    } else {
      skipManualWord(editor, word.from)
      setAfter(word.to)
    }
  }
  const step = (by: 1 | -1) => {
    if (!word) return
    const i = word.points.indexOf(point) + by
    if (i >= 0 && i < word.points.length) setPoint(word.points[i])
  }
  const pieces: string[] = []
  if (word) {
    let from = 0
    for (const p of word.points) {
      pieces.push(word.text.slice(from, p))
      from = p
    }
    pieces.push(word.text.slice(from))
  }

  return (
    <div
      className="modal-backdrop manual-hyphenation-backdrop"
      ref={keys.ref}
      onKeyDown={(e) => {
        if (e.key === 'ArrowLeft') step(-1)
        else if (e.key === 'ArrowRight') step(1)
        else keys.onKeyDown(e)
      }}
    >
      <form
        className="modal gs-form tco-dialog hlc-dialog manual-hyphenation-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          if (word) answer(true)
          else if (found === 'done') onClose()
        }}
      >
        <h2>{t('ribbonHyphManualTitle')}</h2>
        {found === 'done' ? (
          <p className="manual-hyphenation-done">{t('ribbonHyphDone')}</p>
        ) : (
          <div className="manual-hyphenation-row">
            <span>{t('ribbonHyphAt')}</span>
            <div className="manual-hyphenation-word" aria-busy={!word}>
              {pieces.map((piece, i) => (
                <span key={i}>
                  {i > 0 && (
                    <button
                      type="button"
                      className={
                        word!.points[i - 1] === point
                          ? 'manual-hyphenation-point active'
                          : 'manual-hyphenation-point'
                      }
                      aria-pressed={word!.points[i - 1] === point}
                      onClick={() => setPoint(word!.points[i - 1])}
                    >
                      -
                    </button>
                  )}
                  {piece}
                </span>
              ))}
            </div>
          </div>
        )}
        <div className="modal-actions">
          {found === 'done' ? (
            <button type="submit" className="primary" autoFocus>
              {t('ribbonOk')}
            </button>
          ) : (
            <>
              <button type="submit" className="primary" disabled={!word}>
                {t('ribbonPropYes')}
              </button>
              <button type="button" disabled={!word} onClick={() => answer(false)}>
                {t('ribbonPropNo')}
              </button>
              <button type="button" onClick={onClose}>
                {t('ribbonCancel')}
              </button>
            </>
          )}
        </div>
      </form>
    </div>
  )
}
