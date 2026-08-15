import { useEffect, useRef, useState } from 'react'
import appIcon from './assets/app-icon.png'
import { useI18n } from './locale'
import type { StringKey } from './locale'
import './onboarding.css'

interface OnboardingProps {
  /** called when the user finishes the last slide or clicks skip */
  onDone: () => void
}

interface Slide {
  titleKey: StringKey
  /** 18px dark line right under the title */
  subtitleKey: StringKey
  /** 16px muted paragraph below the title block */
  bodyKey?: StringKey
  /** render the body in the dimmer footnote gray (slide 3's credits disclaimer) */
  bodyDim?: boolean
  /** community slide shows the credits offer panel with the "Join GenTeam" call-to-action */
  showOffer?: boolean
  /** closing slide shows the "star us on GitHub" hint */
  showStar?: boolean
  art: 'logo' | 'gift' | 'check'
}

const SLIDES: readonly Slide[] = [
  { titleKey: 'onbTitle1', subtitleKey: 'onbSubtitle1', bodyKey: 'onbBody1', art: 'logo' },
  { titleKey: 'onbTitle2', subtitleKey: 'onbBody2', showOffer: true, art: 'gift' },
  {
    titleKey: 'onbTitle3',
    subtitleKey: 'onbBody3',
    bodyKey: 'onbNote3',
    bodyDim: true,
    showStar: true,
    art: 'check',
  },
]

/** render `**emphasized**` segments of a localized string as <strong> */
function renderEmphasis(text: string) {
  return text
    .split('**')
    .map((part, i) => (i % 2 === 1 ? <strong key={part}>{part}</strong> : part))
}

/* exact vectors from the design spec:
 * 60px canvas, 4px strokes — same visual mass as the 60px app icon */
function SlideArt({ kind }: { kind: Slide['art'] }) {
  if (kind === 'logo') {
    return <img className="onb-art onb-art-logo" src={appIcon} alt="" />
  }
  if (kind === 'gift') {
    // hand-drawn gift kept over the spec vector deliberately; 48 canvas at
    // strokeWidth 3.2 renders the same 4px strokes at 60px as the check icon
    return (
      <span className="onb-art onb-art-badge onb-art-gift" aria-hidden="true">
        <svg
          viewBox="0 0 48 48"
          fill="none"
          stroke="currentColor"
          strokeWidth="3.2"
          strokeLinecap="round"
          strokeLinejoin="round"
        >
          <rect x="6" y="14" width="36" height="9" rx="2" />
          <path d="M9.5 23v15a4 4 0 0 0 4 4h21a4 4 0 0 0 4-4V23" />
          <path d="M24 14v28" />
          <path d="M24 14c-7 0-10.5-2.6-10.5-6 0-2.5 2-4.5 4.5-4.5 4.2 0 6 5.3 6 10.5Z" />
          <path d="M24 14c7 0 10.5-2.6 10.5-6 0-2.5-2-4.5-4.5-4.5-4.2 0-6 5.3-6 10.5Z" />
        </svg>
      </span>
    )
  }
  return (
    <span className="onb-art onb-art-badge onb-art-check" aria-hidden="true">
      <svg viewBox="0 0 60 60" fill="none" stroke="currentColor" strokeWidth="4">
        <path
          d="M29.9883 5.5C43.5194 5.5 54.4883 16.469 54.4883 30C54.4883 43.5311 43.5194 54.5 29.9883 54.5C16.4573 54.5 5.48828 43.5311 5.48828 30C5.48828 16.469 16.4573 5.5 29.9883 5.5Z"
          strokeLinecap="round"
          strokeLinejoin="round"
        />
        <path
          d="M18.125 33.75L24.7764 40.4014C25.8727 41.4977 27.6924 41.342 28.5865 40.0753L41.875 21.25"
          strokeLinecap="round"
        />
      </svg>
    </span>
  )
}

export function Onboarding({ onDone }: OnboardingProps) {
  const { t } = useI18n()
  const [index, setIndex] = useState(0)
  const cardRef = useRef<HTMLDivElement>(null)
  const slide = SLIDES[index]
  const isLast = index === SLIDES.length - 1

  const next = () => {
    if (isLast) onDone()
    else setIndex(index + 1)
  }

  // move focus into the dialog on mount so keyboard users start inside it
  // (the container, not a button, so no focus ring shows on open)
  useEffect(() => {
    cardRef.current?.focus()
  }, [])

  // slide changes can strip focus from the active control (leaving slide 2
  // makes its GenTeam button inert, which blurs it) — pull focus back onto the
  // card so it never drops to body
  useEffect(() => {
    const card = cardRef.current
    const active = document.activeElement
    if (card && (!(active instanceof HTMLElement) || !card.contains(active))) card.focus()
  }, [index])

  // keyboard handling: Escape skips, Enter / ArrowRight advance, ArrowLeft goes
  // back, and Tab is trapped inside the dialog (aria-modal). Enter is ignored
  // when a button is focused so the native click doesn't double-fire.
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        onDone()
        return
      }
      if (event.key === 'Tab') {
        const card = cardRef.current
        if (!card) return
        // inactive slides stay mounted (stacked for the height lock) but are
        // inert — their buttons must not enter the tab cycle
        const focusables = Array.from(card.querySelectorAll<HTMLElement>('button')).filter(
          (el) => !el.closest('[inert]'),
        )
        if (focusables.length === 0) return
        const first = focusables[0]
        const last = focusables[focusables.length - 1]
        const active = document.activeElement
        // the card itself holds focus on open/slide change; from there, Tab in
        // either direction must land on a dialog control, never behind the modal
        const onButton = active instanceof HTMLElement && focusables.includes(active)
        if (!onButton) {
          event.preventDefault()
          ;(event.shiftKey ? last : first).focus()
        } else if (event.shiftKey && active === first) {
          event.preventDefault()
          last.focus()
        } else if (!event.shiftKey && active === last) {
          event.preventDefault()
          first.focus()
        }
        return
      }
      const buttonFocused =
        event.target instanceof HTMLElement && event.target.closest('button') !== null
      if ((event.key === 'Enter' && !buttonFocused) || event.key === 'ArrowRight') next()
      if (event.key === 'ArrowLeft') setIndex((i) => Math.max(0, i - 1))
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  })

  return (
    <div className="onb-overlay" role="dialog" aria-modal="true" aria-label={t(slide.titleKey)}>
      <div className="onb-card" ref={cardRef} tabIndex={-1}>
        {/* all slides stay mounted, stacked in one grid cell: the card locks to
            the tallest slide's height for the language, so the footer and its
            buttons never move between steps. Inactive slides are inert. */}
        <div className="onb-stage">
          {SLIDES.map((s, i) => (
            <div
              className={`onb-slide${i === index ? ' active' : ''}`}
              key={s.titleKey}
              inert={i !== index}
            >
              <SlideArt kind={s.art} />
              <h2 className="onb-title">{t(s.titleKey)}</h2>
              <p className="onb-subtitle">{t(s.subtitleKey)}</p>
              {s.bodyKey && (
                <p className={`onb-body${s.bodyDim ? ' onb-body-dim' : ''}`}>{t(s.bodyKey)}</p>
              )}
              {s.showStar && (
                <div className="onb-star">
                  <p className="onb-star-hint">{t('onbStarHint')}</p>
                  <button
                    className="onb-star-btn"
                    onClick={() => void window.aiOffice.openGitHubRepo()}
                  >
                    <svg
                      width="13"
                      height="13"
                      viewBox="0 0 24 24"
                      fill="currentColor"
                      aria-hidden="true"
                    >
                      <path d="M12 2.5l2.9 5.9 6.5.9-4.7 4.6 1.1 6.5L12 17.3l-5.8 3.1 1.1-6.5L2.6 9.3l6.5-.9L12 2.5z" />
                    </svg>
                    {t('starOnGitHub')}
                  </button>
                </div>
              )}
              {s.showOffer && (
                <div className="onb-offer">
                  <p className="onb-credits">{renderEmphasis(t('onbCredits'))}</p>
                  <button className="onb-join" onClick={() => void window.aiOffice.openGenTeam()}>
                    {t('onbJoinGenTeam')}
                    <svg width="11" height="11" viewBox="0 0 12 12" fill="none" aria-hidden="true">
                      <path
                        d="M3.5 8.5 8.5 3.5M4.5 3.5h4v4"
                        stroke="currentColor"
                        strokeWidth="1.4"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                      />
                    </svg>
                  </button>
                </div>
              )}
            </div>
          ))}
        </div>

        <div className="onb-footer">
          <div className="onb-dots">
            {SLIDES.map((s, i) => (
              <button
                key={s.titleKey}
                className={`onb-dot${i === index ? ' active' : ''}`}
                aria-label={t('onbStepAria', { n: i + 1, total: SLIDES.length })}
                aria-current={i === index}
                onClick={() => setIndex(i)}
              />
            ))}
          </div>
          <div className="onb-nav">
            <button className="onb-skip" onClick={onDone}>
              {t('onbSkip')}
            </button>
            {index > 0 && (
              <button className="onb-back" onClick={() => setIndex(index - 1)}>
                {t('onbBack')}
              </button>
            )}
            <button className="onb-next" onClick={next}>
              {isLast ? t('onbStart') : t('onbNext')}
            </button>
          </div>
        </div>
      </div>
    </div>
  )
}
