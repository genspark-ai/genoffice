import { useState, type ReactNode } from 'react'
import { formatNumber, type CustomNumberingLevel } from '@genoffice/docx-engine'
import { useI18n, type StringKey } from '../i18n/locale'
import {
  libraryCardCustomized,
  listLibrary,
  previewLevelText,
  resetLibraryCard,
  type ListPresetKind,
} from '../list-presets'
import { NO_LIST_STYLE_ID, listStylePreview, type ListStyleEntry } from '../list-styles'
import type { ListNumberingMode } from '../numbering-actions'
import type { ListDialogKind } from './ListDialogs'
import { useModalKeys } from './modal-keys'

/** the three libraries, one per tab, with the dialog that customizes each */
export const LIST_TABS: Array<{
  kind: ListPresetKind
  labelKey: StringKey
  customize: ListDialogKind
}> = [
  { kind: 'bullets', labelKey: 'ribbonBnBulleted', customize: 'bullet' },
  { kind: 'numbers', labelKey: 'ribbonBnNumbered', customize: 'number' },
  { kind: 'multi', labelKey: 'ribbonBnOutline', customize: 'multi' },
]

/** what OK asks for: no list, or one of the tab's presets */
export type ListChoice = { kind: ListPresetKind; levels: CustomNumberingLevel[] | null }

function preview(kind: ListPresetKind, levels: CustomNumberingLevel[]): ReactNode {
  if (kind === 'bullets')
    return (
      <span
        style={{
          fontFamily: levels[0].font,
          color: levels[0].color ? `#${levels[0].color}` : undefined,
        }}
      >
        {levels[0].lvlText}
      </span>
    )
  if (kind === 'multi')
    return [0, 1, 2].map((lvl) => (
      <span key={lvl} style={{ paddingLeft: lvl * 10 }}>
        {previewLevelText(levels, lvl)} ———
      </span>
    ))
  return [0, 1, 2].map((k) => (
    <span key={k} className="list-gallery-preview-row">
      <span className="list-gallery-preview-prefix">
        {levels[0].lvlText.replace(
          /%1/g,
          formatNumber((levels[0].start ?? 1) + k, levels[0].numFmt),
        )}
      </span>
      <span className="list-gallery-preview-line" />
    </span>
  ))
}

const CARD_LOOK: Record<ListPresetKind, string> = {
  bullets: 'list-gallery-glyph',
  numbers: 'list-gallery-preview',
  multi: 'list-gallery-card-multi',
}

const STYLES_TAB = LIST_TABS.length

/** Word's Format > Bullets and Numbering: a library per tab, None first, Customize for the rest; the last tab lists the named list styles */
export function BulletsNumberingDialog({
  onApply,
  onCustomize,
  listStyles,
  listStylesInUse,
  initialStyleId,
  onApplyListStyle,
  onAddListStyle,
  onModifyListStyle,
  onDeleteListStyle,
  numbering,
  onNumbering,
  onClose,
}: {
  onApply: (choice: ListChoice) => void
  /** opens the dialog that defines a bullet, number format or multilevel list, in this one's place; with a card, it changes that card */
  onCustomize: (kind: ListDialogKind, card?: number) => void
  listStyles: ListStyleEntry[]
  listStylesInUse: { used: Set<string>; current: string | null }
  /** opens on the List Styles tab with this style picked (a style just added) */
  initialStyleId?: string
  onApplyListStyle: (styleId: string) => void
  /** opens the dialog that names and defines a new list style, in this one's place */
  onAddListStyle: () => void
  /** opens Modify Style on the list style, in this one's place */
  onModifyListStyle: (styleId: string) => void
  onDeleteListStyle: (styleId: string) => void
  /** how the caret's numbered item counts now; null when it is not in a numbered list */
  numbering: { mode: ListNumberingMode; canContinue: boolean } | null
  onNumbering: (mode: ListNumberingMode) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [tab, setTab] = useState(initialStyleId ? STYLES_TAB : 0)
  // -1 is None; null until the user picks, and then OK changes nothing
  const [picked, setPicked] = useState<number | null>(null)
  const [styleId, setStyleId] = useState<string | null>(initialStyleId ?? listStylesInUse.current)
  // Reset rewrites the stored library; the cards are read again
  const [, setLibraryEpoch] = useState(0)
  const [showInUse, setShowInUse] = useState(false)
  // null until the user picks: the radios then show how the item counts now
  const [countingPick, setCounting] = useState<ListNumberingMode | null>(null)
  const counting = countingPick ?? numbering?.mode ?? 'restart'
  const onStyles = tab === STYLES_TAB
  const { kind, customize } = LIST_TABS[onStyles ? 0 : tab]
  const presets = listLibrary(kind)
  const card = picked !== null && picked >= 0 ? picked : undefined
  const shownStyles = showInUse
    ? listStyles.filter(
        (s) => s.styleId === NO_LIST_STYLE_ID || listStylesInUse.used.has(s.styleId),
      )
    : listStyles
  const style = shownStyles.find((s) => s.styleId === styleId)
  const styleLabel = (s: ListStyleEntry) => (s.labelKey ? t(s.labelKey) : (s.label ?? s.name))
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form tco-dialog bn-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          onClose()
          if (onStyles) {
            if (style) onApplyListStyle(style.styleId)
          } else if (picked !== null) onApply({ kind, levels: picked < 0 ? null : presets[picked] })
          else if (numbering && kind !== 'bullets' && counting !== numbering.mode)
            onNumbering(counting)
        }}
      >
        <h2>{t('ribbonBnTitle')}</h2>
        <div className="dlg-tabs" role="tablist">
          {LIST_TABS.map((entry, i) => (
            <button
              key={entry.kind}
              type="button"
              role="tab"
              aria-selected={tab === i}
              className={tab === i ? 'active' : ''}
              onClick={() => {
                setTab(i)
                setPicked(null)
              }}
            >
              {t(entry.labelKey)}
            </button>
          ))}
          <button
            type="button"
            role="tab"
            aria-selected={onStyles}
            className={onStyles ? 'active' : ''}
            onClick={() => setTab(STYLES_TAB)}
          >
            {t('ribbonBnListStyles')}
          </button>
        </div>
        <div className="bn-styles" hidden={!onStyles}>
          <div className="bn-styles-col">
            <div className="doc-dialog-caption">{t('ribbonPbStyle')}</div>
            <div className="bn-styles-list" role="listbox" aria-label={t('ribbonBnListStyles')}>
              {shownStyles.map((s) => (
                <button
                  key={s.styleId}
                  type="button"
                  role="option"
                  aria-selected={s === style}
                  className={s === style ? 'active' : ''}
                  onClick={() => setStyleId(s.styleId)}
                >
                  {styleLabel(s)}
                </button>
              ))}
            </div>
            <div className="bn-styles-tools">
              <button
                type="button"
                aria-label={t('ribbonAdd')}
                onClick={() => {
                  onClose()
                  onAddListStyle()
                }}
              >
                +
              </button>
              <button
                type="button"
                aria-label={t('ribbonFitRemove')}
                // Word's built-in list styles cannot be deleted
                disabled={!style?.custom}
                onClick={() => {
                  if (!style) return
                  if (!window.confirm(t('appStyleDeleteConfirm', { name: styleLabel(style) })))
                    return
                  onDeleteListStyle(style.styleId)
                  setStyleId(null)
                }}
              >
                −
              </button>
              <button
                type="button"
                name="list-style-modify"
                disabled={!style || style.styleId === NO_LIST_STYLE_ID || style.numberless}
                onClick={() => {
                  if (!style) return
                  onClose()
                  onModifyListStyle(style.styleId)
                }}
              >
                {t('ribbonStyleModify')}
              </button>
            </div>
            <label className="bn-styles-show">
              {t('ribbonGroupShow')}
              <select
                value={showInUse ? 'inUse' : 'all'}
                onChange={(e) => {
                  const inUse = e.target.value === 'inUse'
                  setShowInUse(inUse)
                  // a pick the filter hides is no longer on screen, so OK must not apply it
                  if (
                    inUse &&
                    styleId !== NO_LIST_STYLE_ID &&
                    !listStylesInUse.used.has(styleId ?? '')
                  )
                    setStyleId(null)
                }}
              >
                <option value="all">{t('ribbonStylesShowAll')}</option>
                <option value="inUse">{t('ribbonStylesShowInUse')}</option>
              </select>
            </label>
          </div>
          <div className="bn-styles-col">
            <div className="doc-dialog-caption">{t('ribbonBnListPreview')}</div>
            <div className="bn-styles-preview" role="img" aria-label={t('ribbonBnListPreview')}>
              {listStylePreview(style?.levels ?? []).map((marker, lvl) => (
                <span key={lvl} style={{ paddingInlineStart: lvl * 10 }}>
                  <span>{marker}</span>
                  <span className="bn-styles-preview-line" />
                </span>
              ))}
            </div>
          </div>
        </div>
        <div className="list-gallery list-gallery-word bn-grid" role="listbox" hidden={onStyles}>
          <button
            type="button"
            role="option"
            aria-selected={picked === -1}
            className={`list-gallery-card list-gallery-none${picked === -1 ? ' selected' : ''}`}
            onClick={() => setPicked(-1)}
          >
            {t('ribbonListNone')}
          </button>
          {presets.map((levels, i) => (
            <button
              key={`${kind}-${i}`}
              type="button"
              role="option"
              aria-selected={picked === i}
              className={`list-gallery-card ${CARD_LOOK[kind]}${picked === i ? ' selected' : ''}`}
              onClick={() => setPicked(i)}
            >
              {preview(kind, levels)}
            </button>
          ))}
        </div>
        {!onStyles && kind !== 'bullets' && (
          <fieldset className="bn-numbering">
            <legend>{t('ribbonBnListNumbering')}</legend>
            {(['restart', 'continue'] as const).map((mode) => (
              <label key={mode} className="tco-check">
                <input
                  type="radio"
                  name="listNumbering"
                  value={mode}
                  checked={counting === mode}
                  // a format picked above makes a new list, which starts over by itself
                  disabled={
                    !numbering || picked !== null || (mode === 'continue' && !numbering.canContinue)
                  }
                  onChange={() => setCounting(mode)}
                />
                {t(mode === 'restart' ? 'ribbonBnRestart' : 'ribbonBnContinue')}
              </label>
            ))}
          </fieldset>
        )}
        <div className="modal-actions">
          {/* Word's List Styles tab has no Customize; the span keeps OK and Cancel on the right */}
          {onStyles ? (
            <span />
          ) : (
            <div className="tco-actions-right">
              <button
                type="button"
                onClick={() => {
                  onClose()
                  onCustomize(customize, card)
                }}
              >
                {t('ribbonBnCustomize')}
              </button>
              <button
                type="button"
                name="reset"
                disabled={card === undefined || !libraryCardCustomized(kind, card)}
                onClick={() => {
                  if (card === undefined) return
                  resetLibraryCard(kind, card)
                  setLibraryEpoch((n) => n + 1)
                }}
              >
                {t('ribbonBnReset')}
              </button>
            </div>
          )}
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
