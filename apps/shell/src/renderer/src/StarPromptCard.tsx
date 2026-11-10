import { useI18n } from './locale'
import './star-prompt.css'

/**
 * One-time "star us on GitHub" invitation, shown over the home screen after
 * the main process decides the user has gotten real value out of the app
 * (see main/star-prompt.ts). Non-modal: a small bottom-right card that never
 * blocks work. Any reaction resolves it via starPromptAction.
 */

/** with at least this many opens, the title reflects the user's own usage
 * ("you've opened N documents") — the strongest-converting copy per industry
 * data; below it (e.g. upgrade-launch prompts) fall back to the generic title */
const PERSONALIZED_MIN_OPENS = 5

interface StarPromptCardProps {
  /** lifetime documents opened, from the main process's prompt state */
  docOpens: number
  /** called after the user reacts, whatever the reaction — unmounts the card */
  onClose: () => void
}

function StarIcon() {
  return (
    <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
      <path d="M12 2.5l2.9 5.9 6.5.9-4.7 4.6 1.1 6.5L12 17.3l-5.8 3.1 1.1-6.5L2.6 9.3l6.5-.9L12 2.5z" />
    </svg>
  )
}

export function StarPromptCard({ docOpens, onClose }: StarPromptCardProps) {
  const { t } = useI18n()

  const react = (action: 'starred' | 'later') => {
    void window.aiOffice.starPromptAction(action).catch(() => {})
    onClose()
  }

  const title =
    docOpens >= PERSONALIZED_MIN_OPENS
      ? t('starPromptTitleN', { n: docOpens })
      : t('starPromptTitle')

  return (
    <div className="star-prompt" role="dialog" aria-label={title}>
      <button
        className="star-prompt-close"
        aria-label={t('starPromptLater')}
        onClick={() => react('later')}
      >
        <svg width="12" height="12" viewBox="0 0 12 12" aria-hidden="true">
          <path
            d="M2 2l8 8M10 2L2 10"
            stroke="currentColor"
            strokeWidth="1.4"
            strokeLinecap="round"
          />
        </svg>
      </button>
      <div className="star-prompt-head">
        <span className="star-prompt-icon">
          <StarIcon />
        </span>
        <h3 className="star-prompt-title">{title}</h3>
      </div>
      <p className="star-prompt-body">{t('starPromptBody')}</p>
      <div className="star-prompt-actions">
        <button
          className="star-prompt-go"
          onClick={() => {
            void window.aiOffice.openGitHubRepo().catch(() => {})
            react('starred')
          }}
        >
          <StarIcon />
          {t('starPromptGo')}
        </button>
        <button className="star-prompt-done" onClick={() => react('starred')}>
          {t('starPromptDone')}
        </button>
      </div>
    </div>
  )
}
