import { useState } from 'react'
import { useI18n } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

export type PageNumberPlace = 'header' | 'footer'
export type PageNumberAlign = 'left' | 'center' | 'right' | 'inside' | 'outside'

/** Word's Insert > Page Numbers: where on the page the number goes */
export function PageNumbersDialog({
  showOnFirst,
  firstLocked,
  onApply,
  onFormat,
  onClose,
}: {
  /** the first page shows the number: false when the section has a first page of its own */
  showOnFirst: boolean
  /** the first-page choice cannot be changed from here for this section */
  firstLocked: boolean
  onApply: (place: PageNumberPlace, align: PageNumberAlign, showOnFirst: boolean) => void
  /** opens Format Page Numbers in this dialog's place */
  onFormat: () => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  // Word's own defaults
  const [place, setPlace] = useState<PageNumberPlace>('footer')
  const [align, setAlign] = useState<PageNumberAlign>('right')
  const [first, setFirst] = useState(showOnFirst)
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form tco-dialog hlc-dialog page-numbers-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          onClose()
          onApply(place, align, first)
        }}
      >
        <h2>{t('ribbonPnDialogTitle')}</h2>
        <div className="modal-row margin-row">
          <label>
            {t('ribbonPosition')}
            <select value={place} onChange={(e) => setPlace(e.target.value as PageNumberPlace)}>
              <option value="header">{t('ribbonPnTopOfPage')}</option>
              <option value="footer">{t('ribbonPnBottomOfPage')}</option>
            </select>
          </label>
        </div>
        <div className="modal-row margin-row">
          <label>
            {t('appAlignment')}
            <select
              name="align"
              value={align}
              onChange={(e) => setAlign(e.target.value as PageNumberAlign)}
            >
              <option value="left">{t('appAlignLeft')}</option>
              <option value="center">{t('appTabCenter')}</option>
              <option value="right">{t('appAlignRight')}</option>
              <option value="inside">{t('ribbonMarginInside')}</option>
              <option value="outside">{t('ribbonMarginOutside')}</option>
            </select>
          </label>
        </div>
        <label className="tco-check">
          <input
            type="checkbox"
            name="showFirst"
            checked={first}
            disabled={firstLocked}
            onChange={(e) => setFirst(e.target.checked)}
          />
          {t('ribbonPnShowFirst')}
        </label>
        <div className="modal-actions">
          <button
            type="button"
            onClick={() => {
              onClose()
              onFormat()
            }}
          >
            {t('ribbonPnFormat')}
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
