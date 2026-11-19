import { useState } from 'react'
import type { RevisionDisplayMode } from '../editor/revision-view'
import { useI18n } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

/** markup shows on screen in the two views that draw it */
export const highlightsOnScreen = (mode: RevisionDisplayMode): boolean =>
  mode === 'all' || mode === 'simple'

/** the view to switch to for the box as ticked, or null when the current one already matches */
export function displayForHighlight(
  current: RevisionDisplayMode,
  onScreen: boolean,
): RevisionDisplayMode | null {
  if (highlightsOnScreen(current) === onScreen) return null
  return onScreen ? 'all' : 'none'
}

/**
 * The view a print runs in: ticked prints the markup (only All Markup does), unticked
 * prints the clean sheet, null prints what the screen shows.
 */
export function displayForPrint(
  current: RevisionDisplayMode,
  print: boolean | null,
): RevisionDisplayMode {
  if (print === null) return current
  if (print) return 'all'
  return current === 'all' ? 'none' : current
}

interface Props {
  trackChanges: boolean
  /** protection decides tracking: the box shows the state and cannot be changed */
  trackLocked: boolean
  display: RevisionDisplayMode
  /** whether a print highlights changes; null follows the screen */
  print: boolean | null
  trackMoves: boolean
  onApply: (next: {
    track: boolean
    moves: boolean
    display: RevisionDisplayMode | null
    print: boolean | null
  }) => void
  onOptions: () => void
  onClose: () => void
}

/** Word's Tools > Track Changes > Highlight Changes */
export function HighlightChangesDialog({
  trackChanges,
  trackLocked,
  display,
  print: printNow,
  trackMoves,
  onApply,
  onOptions,
  onClose,
}: Props) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [track, setTrack] = useState(trackChanges)
  const [onScreen, setOnScreen] = useState(highlightsOnScreen(display))
  const [print, setPrint] = useState(printNow)
  const [moves, setMoves] = useState(trackMoves)

  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form tco-dialog hlc-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          onApply({ track, moves, display: displayForHighlight(display, onScreen), print })
          onClose()
        }}
      >
        <h2>{t('ribbonHlcTitle')}</h2>
        <label className="tco-check">
          <input
            type="checkbox"
            name="track-changes"
            checked={track}
            disabled={trackLocked}
            onChange={(e) => setTrack(e.target.checked)}
          />
          {t('ribbonHlcTrack')}
        </label>
        <label className="tco-check">
          <input
            type="checkbox"
            name="track-moves"
            checked={moves}
            onChange={(e) => setMoves(e.target.checked)}
          />
          {t('ribbonHlcMoves')}
        </label>
        <label className="tco-check">
          <input
            type="checkbox"
            checked={onScreen}
            onChange={(e) => setOnScreen(e.target.checked)}
          />
          {t('ribbonHlcScreen')}
        </label>
        {/* until it is ticked or cleared here, a print draws the markup the screen shows */}
        <label className="tco-check">
          <input
            type="checkbox"
            name="highlight-print"
            checked={print ?? onScreen}
            onChange={(e) => setPrint(e.target.checked)}
          />
          {t('ribbonHlcPrint')}
        </label>
        <div className="modal-actions">
          <button
            type="button"
            onClick={() => {
              onClose()
              onOptions()
            }}
          >
            {t('ribbonHlcOptions')}
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
