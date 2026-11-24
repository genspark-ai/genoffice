import { useState } from 'react'
import { COMBINE_BRACKETS, type CombineBrackets } from '@genoffice/docx-engine'
import {
  COMBINE_BRACKET_PAIRS,
  twoLinesWidthEm,
  type EaLayoutAttr,
  type EaLayoutChoice,
} from '../editor/asian-layout'
import { useI18n } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

interface Props {
  text: string
  /** the layout the selected text has now */
  layout: EaLayoutAttr | null
  /** null takes the layout off */
  onApply: (choice: EaLayoutChoice | null) => void
  onClose: () => void
}

function Frame({
  title,
  className,
  removable,
  choice,
  onApply,
  onClose,
  children,
}: {
  title: string
  className: string
  removable: boolean
  choice: EaLayoutChoice
  onApply: Props['onApply']
  onClose: () => void
  children: React.ReactNode
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const finish = (next: EaLayoutChoice | null) => {
    onClose()
    onApply(next)
  }
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className={`modal gs-form tco-dialog hlc-dialog eq-layout-dialog ${className}`}
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          finish(choice)
        }}
      >
        <h2>{title}</h2>
        {children}
        <div className="modal-actions">
          <button type="button" name="remove" disabled={!removable} onClick={() => finish(null)}>
            {t('ribbonFitRemove')}
          </button>
          <div className="tco-actions-right">
            <button type="button" onClick={onClose}>
              {t('ribbonCancel')}
            </button>
            <button type="submit" className="primary" autoFocus>
              {t('ribbonOk')}
            </button>
          </div>
        </div>
      </form>
    </div>
  )
}

/** Word's Format > Asian Layout > Two Lines in One: the text as two half-size lines, in brackets or not */
export function TwoLinesInOneDialog({ text, layout, onApply, onClose }: Props) {
  const { t } = useI18n()
  const [enclose, setEnclose] = useState(!!layout?.brackets)
  const [style, setStyle] = useState<CombineBrackets>(layout?.brackets ?? 'round')
  return (
    <Frame
      title={t('ribbonTwoLinesInOne')}
      className="two-lines-dialog"
      removable={!!layout?.combine}
      choice={{ kind: 'combine', brackets: enclose ? style : null }}
      onApply={onApply}
      onClose={onClose}
    >
      <label className="ref-check">
        <input
          type="checkbox"
          name="two-lines-brackets"
          checked={enclose}
          onChange={(e) => setEnclose(e.target.checked)}
        />
        {t('ribbonTwoLinesBrackets')}
      </label>
      <div className="modal-row margin-row">
        <label>
          {t('ribbonTwoLinesBracketStyle')}
          <select
            name="two-lines-bracket-style"
            value={style}
            disabled={!enclose}
            onChange={(e) => setStyle(e.target.value as CombineBrackets)}
          >
            {COMBINE_BRACKETS.map((b) => (
              <option key={b} value={b}>
                {COMBINE_BRACKET_PAIRS[b].join(' ')}
              </option>
            ))}
          </select>
        </label>
      </div>
      <div className="eq-layout-preview" aria-label={t('ribbonPbPreview')}>
        <span
          className="ea-two-lines-sample"
          data-ea-brackets={enclose ? style : undefined}
          style={{ width: `${twoLinesWidthEm(text)}em` }}
        >
          {text}
        </span>
      </div>
    </Frame>
  )
}

/** Word's Format > Asian Layout > Horizontal in Vertical: the text upright where the line runs down the page */
export function HorizontalInVerticalDialog({ text, layout, onApply, onClose }: Props) {
  const { t } = useI18n()
  // Word offers the fitting ticked; text that already stands upright shows what it has
  const [fit, setFit] = useState(layout?.vert ? !!layout.vertCompress : true)
  return (
    <Frame
      title={t('ribbonHorizontalInVertical')}
      className="hiv-dialog"
      removable={!!layout?.vert}
      choice={{ kind: 'vert', fit }}
      onApply={onApply}
      onClose={onClose}
    >
      <label className="ref-check">
        <input
          type="checkbox"
          name="hiv-fit"
          checked={fit}
          onChange={(e) => setFit(e.target.checked)}
        />
        {t('ribbonHivFit')}
      </label>
      <div className="eq-layout-preview ea-hiv-preview" aria-label={t('ribbonPbPreview')}>
        <span className={fit ? 'ea-hiv-sample ea-hiv-fit' : 'ea-hiv-sample'}>{text}</span>
      </div>
    </Frame>
  )
}
