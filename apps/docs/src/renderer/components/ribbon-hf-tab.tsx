import { useI18n } from '../i18n/locale'
import type { HfView } from '../doc-state'
import {
  IconCaret,
  IconCheckbox,
  IconCheckboxChecked,
  IconClock,
  IconClose,
  IconFooter,
  IconHeader,
  IconPageNumber,
} from './icons'
import { BIG, toggleDropdown, type TabProps } from './ribbon-tabs'
import { LengthInput } from './LengthInput'

/** what the contextual Header & Footer tab is looking at */
export interface HfEditingInfo {
  kind: 'header' | 'footer'
  /** section index of the strip being edited */
  section: number
  sectionCount: number
  variant: HfView
  /** Word's Link to Previous: null for the first section (nothing to link to) */
  linked: boolean | null
  headerDistTwips: number
  footerDistTwips: number
  titlePg: boolean
  evenOddHf: boolean
}

export type HfAction =
  | { type: 'edit'; kind: 'header' | 'footer' }
  | { type: 'remove'; kind: 'header' | 'footer' }
  | { type: 'goto'; kind: 'header' | 'footer' }
  | { type: 'close' }
  | { type: 'pageNumber'; kind: 'header' | 'footer'; align: 'left' | 'center' | 'right' }
  | { type: 'pageNumberHere' }
  | { type: 'numPages' }
  | { type: 'pageNumberFormat' }
  | { type: 'removePageNumbers' }
  | { type: 'field'; instr: 'DATE' | 'TIME' }
  | { type: 'linkToPrevious'; on: boolean }
  | { type: 'distance'; kind: 'header' | 'footer'; twips: number }
  | { type: 'titlePg'; on: boolean }
  | { type: 'evenOddHf'; on: boolean }

/** Word's Header from Top / Footer from Bottom range: 0–22" */
const MAX_HF_DISTANCE_TWIPS = 31680

export function HeaderFooterTab({
  info,
  dropdown,
  setDropdown,
  onAction,
}: {
  info: HfEditingInfo
  dropdown: string | null
  setDropdown: TabProps['setDropdown']
  onAction: (action: HfAction) => void
}) {
  const { t } = useI18n()
  const act = (action: HfAction) => {
    setDropdown(() => null)
    onAction(action)
  }
  const other = info.kind === 'header' ? 'footer' : 'header'
  const distanceInput = (kind: 'header' | 'footer', twips: number, label: string) => (
    <label className="layout-num hf-distance" data-tip={label}>
      <span>{label}</span>
      <LengthInput
        key={kind}
        value={twips}
        min={0}
        max={MAX_HF_DISTANCE_TWIPS}
        ariaLabel={label}
        onCommit={(next) => onAction({ type: 'distance', kind, twips: next ?? 0 })}
        onEnter={() => (document.activeElement as HTMLElement | null)?.blur()}
      />
    </label>
  )
  const hfMenu = (kind: 'header' | 'footer') => (
    <div data-rb-panel="" className="layout-menu">
      <button onClick={() => act({ type: 'edit', kind })}>
        {t(kind === 'header' ? 'ribbonHfEditHeader' : 'ribbonHfEditFooter')}
      </button>
      <button onClick={() => act({ type: 'remove', kind })}>
        {t(kind === 'header' ? 'ribbonHfRemoveHeader' : 'ribbonHfRemoveFooter')}
      </button>
    </div>
  )
  const positions = [
    [t('ribbonPnTopLeft'), 'header', 'left'],
    [t('ribbonPnTopCenter'), 'header', 'center'],
    [t('ribbonPnTopRight'), 'header', 'right'],
    [t('ribbonPnBottomLeft'), 'footer', 'left'],
    [t('ribbonPnBottomCenter'), 'footer', 'center'],
    [t('ribbonPnBottomRight'), 'footer', 'right'],
  ] as Array<[string, 'header' | 'footer', 'left' | 'center' | 'right']>
  return (
    <div className="hf-ribbon-body">
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <div className="rb-split-wrap">
            <button
              className={`rb-big ${info.kind === 'header' ? 'active' : ''}`}
              data-tip={t('ribbonHeaderTip')}
              onClick={() => toggleDropdown(setDropdown, 'hf-header')}
            >
              <span className="rb-big-icon">
                <IconHeader size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonHeader')}</span>
            </button>
            {dropdown === 'hf-header' && hfMenu('header')}
          </div>
          <div className="rb-split-wrap">
            <button
              className={`rb-big ${info.kind === 'footer' ? 'active' : ''}`}
              data-tip={t('ribbonFooterTip')}
              onClick={() => toggleDropdown(setDropdown, 'hf-footer')}
            >
              <span className="rb-big-icon">
                <IconFooter size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonFooter')}</span>
            </button>
            {dropdown === 'hf-footer' && hfMenu('footer')}
          </div>
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              data-tip={t('ribbonPageNumber')}
              onClick={() => toggleDropdown(setDropdown, 'hf-pagenum')}
            >
              <span className="rb-big-icon">
                <IconPageNumber size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonPageNumber')}</span>
            </button>
            {dropdown === 'hf-pagenum' && (
              <div data-rb-panel="" className="layout-menu">
                {positions.map(([label, kind, align]) => (
                  <button key={label} onClick={() => act({ type: 'pageNumber', kind, align })}>
                    {label}
                  </button>
                ))}
                <button onClick={() => act({ type: 'pageNumberHere' })}>
                  {t('ribbonHfPnCurrent')}
                </button>
                <button onClick={() => act({ type: 'numPages' })}>{t('ribbonHfNumPages')}</button>
                <button onClick={() => act({ type: 'pageNumberFormat' })}>
                  {t('ribbonPnFormat')}
                </button>
                <button onClick={() => act({ type: 'removePageNumbers' })}>
                  {t('ribbonPnRemove')}
                </button>
              </div>
            )}
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupHeaderFooter')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              data-tip={t('ribbonHfDateTimeTip')}
              onClick={() => toggleDropdown(setDropdown, 'hf-datetime')}
            >
              <span className="rb-big-icon">
                <IconClock size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonHfDateTime')}</span>
            </button>
            {dropdown === 'hf-datetime' && (
              <div data-rb-panel="" className="layout-menu">
                <button onClick={() => act({ type: 'field', instr: 'DATE' })}>
                  {t('ribbonHfDate')}
                </button>
                <button onClick={() => act({ type: 'field', instr: 'TIME' })}>
                  {t('ribbonHfTime')}
                </button>
              </div>
            )}
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupHfInsert')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className="rb-big"
            data-tip={t(other === 'header' ? 'ribbonHfGoToHeader' : 'ribbonHfGoToFooter')}
            onClick={() => act({ type: 'goto', kind: other })}
          >
            <span className="rb-big-icon">
              {other === 'header' ? <IconHeader size={BIG} /> : <IconFooter size={BIG} />}
            </span>
            <span>{t(other === 'header' ? 'ribbonHfGoToHeader' : 'ribbonHfGoToFooter')}</span>
          </button>
          <div className="rb-col">
            <button
              className={`rb-small ${info.linked ? 'active' : ''}`}
              disabled={info.linked === null}
              data-tip={t('ribbonHfLinkPrevTip')}
              onClick={() => act({ type: 'linkToPrevious', on: !info.linked })}
            >
              {info.linked ? <IconCheckboxChecked /> : <IconCheckbox />} {t('ribbonHfLinkPrev')}
            </button>
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupHfNavigation')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <div className="rb-col">
            <button
              className={`rb-small ${info.titlePg ? 'active' : ''}`}
              data-tip={t('ribbonDiffFirstPageTip')}
              onClick={() => act({ type: 'titlePg', on: !info.titlePg })}
            >
              {info.titlePg ? <IconCheckboxChecked /> : <IconCheckbox />} {t('ribbonDiffFirstPage')}
            </button>
            <button
              className={`rb-small ${info.evenOddHf ? 'active' : ''}`}
              data-tip={t('ribbonDiffOddEvenTip')}
              onClick={() => act({ type: 'evenOddHf', on: !info.evenOddHf })}
            >
              {info.evenOddHf ? <IconCheckboxChecked /> : <IconCheckbox />} {t('ribbonDiffOddEven')}
            </button>
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupHfOptions')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <div className="rb-col hf-distances">
            {distanceInput('header', info.headerDistTwips, t('ribbonHfHeaderFromTop'))}
            {distanceInput('footer', info.footerDistTwips, t('ribbonHfFooterFromBottom'))}
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupHfPosition')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className="rb-big hf-close"
            data-tip={t('ribbonHfCloseTip')}
            onClick={() => act({ type: 'close' })}
          >
            <span className="rb-big-icon">
              <IconClose size={BIG} />
            </span>
            <span>{t('ribbonHfClose')}</span>
          </button>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupHfClose')}</div>
      </div>
    </div>
  )
}
