import { useI18n, type StringKey } from '../i18n/locale'
import { useMeasurement } from '../use-measurement'
import { useModalKeys } from './modal-keys'
import { FALLBACK_STYLES } from '../style-gallery'
import type { RibbonFormatState } from './ribbon-format-state'

const ALIGN_LABELS: Record<string, StringKey> = {
  left: 'appAlignLeft',
  center: 'appAlignCenter',
  right: 'appAlignRight',
  justify: 'appAlignJustify',
  distribute: 'ribbonDistributed',
}

/** the font style as the Font dialog names it */
export function fontStyleKey(format: Pick<RibbonFormatState, 'bold' | 'italic'>): StringKey {
  if (format.bold && format.italic) return 'appFontBoldItalic'
  if (format.bold) return 'appFontBold'
  return format.italic ? 'appFontItalic' : 'appFontRegular'
}

/** a style's name: the document's, else the built-in one's, else its id (a file without styles.xml) */
export function revealedStyleName(
  styleId: string,
  styles: ReadonlyMap<string, { name: string }> | null | undefined,
): string {
  return styles?.get(styleId)?.name ?? FALLBACK_STYLES.get(styleId)?.name ?? styleId
}

/** the formatting at the caret as label and value pairs: the font, then the paragraph */
export function revealedFormatting(
  format: RibbonFormatState,
  styleName: string,
  length: (twips: number) => string,
): {
  font: Array<[StringKey, string | StringKey[]]>
  paragraph: Array<[StringKey, string | StringKey[]]>
} {
  const align = format.align ?? (format.bidi ? 'right' : 'left')
  return {
    font: [
      ['appFontDialogTitle', format.fontFamily],
      ['appFontSizeLabel', `${format.fontSizePt} pt`],
      ['appFontStyleLabel', [fontStyleKey(format)]],
      ...(format.underline ? ([['appUnderline', '✓']] as Array<[StringKey, string]>) : []),
      ...(format.textColor
        ? ([['appFontColor', `#${format.textColor}`]] as Array<[StringKey, string]>)
        : []),
    ],
    paragraph: [
      ['ribbonPbStyle', styleName],
      ['appAlignment', [ALIGN_LABELS[align] ?? 'appAlignLeft']],
      ['appIndentLeft', length(format.indentLeft)],
      ['appIndentRight', length(format.indentRight)],
      ['appSpaceBefore', `${format.spaceBefore / 20} pt`],
      ['appSpaceAfter', `${format.spaceAfter / 20} pt`],
      ...(format.lineSpacing
        ? ([['appLineSpacingLabel', `${format.lineSpacing}`]] as Array<[StringKey, string]>)
        : []),
    ],
  }
}

/** Word's View > Reveal Formatting: what formatting the text at the caret has */
export function RevealFormattingDialog({
  format,
  styleName,
  onClose,
}: {
  format: RibbonFormatState
  /** the name of the paragraph's style, as the Styles gallery shows it */
  styleName: string
  onClose: () => void
}) {
  const { t } = useI18n()
  const measure = useMeasurement()
  const keys = useModalKeys(onClose)
  const shown = revealedFormatting(format, styleName, measure.format)
  const rows = (list: Array<[StringKey, string | StringKey[]]>) =>
    list.map(([labelKey, value]) => (
      <div key={labelKey} className="reveal-formatting-row">
        <span>{t(labelKey)}</span>
        <span>{typeof value === 'string' ? value : value.map((k) => t(k)).join(' ')}</span>
      </div>
    ))
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal gs-form tco-dialog hlc-dialog reveal-formatting-dialog">
        <h2>{t('ribbonRevealFormatting')}</h2>
        <div className="doc-dialog-caption">{t('appFontDialogTitle')}</div>
        {rows(shown.font)}
        <div className="doc-dialog-caption">{t('appParagraph')}</div>
        {rows(shown.paragraph)}
        <div className="modal-actions">
          <span />
          <div className="tco-actions-right">
            <button type="button" className="primary" onClick={onClose}>
              {t('ribbonClose')}
            </button>
          </div>
        </div>
      </div>
    </div>
  )
}
