import type { SectionScope } from '../commands/menu-commands'
import { useI18n } from '../i18n/locale'

/** what a section dialog acts on when Apply to is left alone */
export const defaultScope = (multiSection: boolean): SectionScope =>
  multiSection ? 'section' : 'whole'

/** Word's Apply to: the caret's section (when there are several), from the caret on, or every section */
export function ApplyToSelect({
  scope,
  multiSection,
  canApplyForward,
  onChange,
}: {
  scope: SectionScope
  multiSection: boolean
  canApplyForward: boolean
  onChange: (scope: SectionScope) => void
}) {
  const { t } = useI18n()
  return (
    <div className="modal-row margin-row">
      <label>
        {t('ribbonPbApplyTo')}
        <select
          name="applyTo"
          value={scope}
          onChange={(e) => onChange(e.target.value as SectionScope)}
        >
          {multiSection && <option value="section">{t('ribbonPbApplySection')}</option>}
          <option value="forward" disabled={!canApplyForward}>
            {t('ribbonPbApplyForward')}
          </option>
          <option value="whole">{t('ribbonPbApplyWhole')}</option>
        </select>
      </label>
    </div>
  )
}
