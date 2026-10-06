import { useEffect, useRef } from 'react'
import { useI18n } from '../i18n/locale'

interface Props {
  /** the exact file text: frontmatter block plus body, as a save would write it */
  value: string
  onChange: (value: string) => void
  /**
   * The pane can only go stale while it sits unfocused, so the caller re-syncs
   * on either edge. Focusing it therefore refreshes before the user can read.
   */
  onFocusChange: () => void
}

/**
 * Raw Markdown surface for the source view: a plain textarea over the exact
 * file text, so constructs the editor has no node for — footnote definitions,
 * reference-style links, raw HTML, comments — can be hand-edited in place.
 *
 * Deliberately a bare textarea: no syntax highlighting, no line numbers and no
 * custom undo stack. Native undo covers the pane while it has focus, and every
 * keystroke is handed straight to the editor, so the pane is never a draft the
 * editor has not seen.
 */
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
