/**
 * Story bar's writing-font picker: the bundled faces, the writing faces the
 * machine has, and a free-text field for any installed family. The choice is
 * a view preference — the document's own fonts stay untouched. The open state
 * lives in the app so the story view's Escape guard closes the picker before
 * it would leave the view.
 */
import { useEffect, useRef, useState } from 'react'
import { t } from '../i18n/locale'
import { STORY_FACE_CHOICES, faceInstalled, sanitizeFaceName } from '../story-font'

export function StoryFontPicker({
  value,
  open,
  onOpenChange,
  onChange,
}: {
  value: string
  open: boolean
  onOpenChange: (open: boolean) => void
  onChange: (family: string) => void
}) {
  const inputRef = useRef<HTMLInputElement>(null)
  const [draft, setDraft] = useState('')

  useEffect(() => {
    if (open) {
      // reopen from scratch: the free-text field ready for a name the list
      // does not carry
      setDraft('')
      inputRef.current?.focus()
    }
  }, [open])

  const current =
    value === ''
      ? t('storyFontAuto')
      : (STORY_FACE_CHOICES.find((c) => c.family === value)?.family ?? value)

  const pick = (family: string): void => {
    onChange(family)
    onOpenChange(false)
  }

  return (
    <>
      <button
        type="button"
        className="story-bar-btn"
        aria-haspopup="true"
        aria-expanded={open}
        onClick={() => onOpenChange(!open)}
      >
        {value === '' ? t('storyBarFont') : `Aa ${current}`}
      </button>
      {open && (
        <>
          <div className="story-font-backdrop" onMouseDown={() => onOpenChange(false)} />
          <div className="story-font-pop" role="menu">
            <button
              type="button"
              role="menuitem"
              className="story-font-item"
              onClick={() => pick('')}
            >
              <span className="story-font-face">{t('storyFontAuto')}</span>
              <span className="story-font-hint">{t('storyFontAutoHint')}</span>
            </button>
            {STORY_FACE_CHOICES.map((choice) => (
              <button
                key={choice.family}
                type="button"
                role="menuitem"
                className={`story-font-item${value === choice.family ? ' on' : ''}`}
                onClick={() => pick(choice.family)}
              >
                <span className="story-font-face">{choice.family}</span>
                {choice.bundled ? (
                  <span className="story-font-bundled">{t('storyFontBundled')}</span>
                ) : (
                  <span className={`story-font-dot${faceInstalled(choice.family) ? ' has' : ''}`} />
                )}
              </button>
            ))}
            <div className="story-font-any">
              <input
                ref={inputRef}
                value={draft}
                placeholder={t('storyFontPlaceholder')}
                spellCheck={false}
                onChange={(e) => setDraft(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' && sanitizeFaceName(draft)) {
                    pick(sanitizeFaceName(draft))
                  }
                }}
              />
              <button
                type="button"
                className="story-bar-btn"
                disabled={!sanitizeFaceName(draft)}
                onClick={() => {
                  if (sanitizeFaceName(draft)) pick(sanitizeFaceName(draft))
                }}
              >
                {t('storyFontUse')}
              </button>
            </div>
          </div>
        </>
      )}
    </>
  )
}
