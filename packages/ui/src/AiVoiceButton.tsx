import React, { useEffect } from 'react'
import type { Lang } from '@genoffice/i18n'
import { IconMic } from './icons'
import { AI_VOICE_LABELS } from './strings-ai-voice'
import type { VoiceInput } from './use-voice-input'

/** How long a refusal stays in the footer before it stops taking up room. */
const ERROR_TTL_MS = 6000

/**
 * The composer microphone: one button over a `useVoiceInput` session. While the
 * microphone is open the glyph becomes three bars, so the recording state is visible
 * without a separate overlay; a refusal shows as inline text next to it and clears
 * itself.
 */
export function AiVoiceButton({
  voice,
  lang,
}: {
  readonly voice: VoiceInput
  readonly lang: Lang
}): React.JSX.Element | null {
  const labels = AI_VOICE_LABELS[lang]
  const { error, clearError, phase, supported } = voice

  useEffect(() => {
    if (!error) return
    const timer = setTimeout(clearError, ERROR_TTL_MS)
    return () => clearTimeout(timer)
  }, [error, clearError])

  if (!supported) return null

  const recording = phase === 'recording'
  const label = recording ? labels.stop : labels.start
  const tip = phase === 'idle' ? `${labels.start} · ${labels.localOnly}` : label

  return (
    <span className="ai-voice">
      {error && (
        <span className="ai-voice-error" role="status" title={labels[error]}>
          {labels[error]}
        </span>
      )}
      <button
        type="button"
        className={recording ? 'ai-voice-btn recording' : 'ai-voice-btn'}
        onClick={voice.toggle}
        disabled={phase === 'transcribing' || phase === 'requesting'}
        aria-label={label}
        aria-pressed={recording}
        data-tip={tip}
      >
        {recording ? (
          <span className="ai-voice-wave" aria-hidden>
            <i />
            <i />
            <i />
          </span>
        ) : (
          <IconMic size={16} />
        )}
      </button>
    </span>
  )
}
