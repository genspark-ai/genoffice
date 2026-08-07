import { useI18n } from '../i18n/locale'

interface Props {
  /** inner YAML text, without the --- fences */
  value: string
  onChange: (value: string) => void
  readOnly?: boolean
}

/**
 * Raw-text properties panel: edits the YAML between the frontmatter fences
 * verbatim. No YAML parsing — whatever the user types is what lands in the
 * file, so exotic YAML survives untouched.
 */
export function FrontmatterPanel({ value, onChange, readOnly }: Props) {
  const { t } = useI18n()
  const rows = Math.min(12, Math.max(2, value.split('\n').length))
  return (
    <div className="fm-panel">
      <div className="fm-title">{t('fmProperties')}</div>
      <textarea
        className="fm-textarea"
        value={value}
        rows={rows}
        spellCheck={false}
        readOnly={readOnly}
        placeholder="key: value"
        onChange={(e) => onChange(e.target.value)}
      />
    </div>
  )
}
