import { useEffect, useRef } from 'react'
import { useI18n } from '../i18n/locale'

interface Props {
  /** the exact file text: frontmatter block plus body, as a save would write it */
  value: string
  onChange: (value: string) => void
  /** fired on both focus edges so the caller can flush and re-sync */
  onFocusChange: () => void
}

/** Raw Markdown surface: a bare textarea (native undo, no highlighting) over the exact file text. */
export function SourcePane({ value, onChange, onFocusChange }: Props) {
  const { t } = useI18n()
  const ref = useRef<HTMLTextAreaElement>(null)

  useEffect(() => {
    ref.current?.focus()
  }, [])

  return (
    <div className="source-pane">
      <div className="source-head">
        <span className="source-title">{t('sourceView')}</span>
        <span className="source-hint">{t('sourceViewHint')}</span>
      </div>
      <textarea
        ref={ref}
        className="source-textarea"
        value={value}
        spellCheck={false}
        wrap="off"
        onFocus={onFocusChange}
        onBlur={onFocusChange}
        onChange={(e) => onChange(e.target.value)}
      />
    </div>
  )
}
